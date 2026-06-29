import logging
import re
from pathlib import Path

import pandas as pd
import yaml

from src.models import LancamentoExtrato

logger = logging.getLogger(__name__)


def carregar_unidades_validas(caminho: str | Path) -> set[int]:
    """Carrega lista oficial de unidades válidas de um YAML."""
    caminho = Path(caminho)
    if not caminho.exists():
        logger.warning("Arquivo de unidades não encontrado: %s", caminho)
        return set()
    with open(caminho, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return set(data.get("unidades", []))


def extrair_unidade_com_validacao(
    historico: str,
    unidades_validas: set[int],
) -> tuple[str | None, str | None, str | None]:
    """
    Extrai unidade do histórico com validação contra lista oficial.

    Retorna (unidade, origem_provavel, justificativa).
    - unidade: str do número da unidade se válida, None se inválida
    - origem_provavel: de onde veio o número (UN., RC.UN., etc.)
    - justificativa: por que foi aceito/rejeitado
    """
    # 1. Padrão RC.UN.XXXXXX (mais específico, testar primeiro)
    m = re.search(r"RC\.UN\.(\d{4,6})", historico)
    if m:
        unid = str(int(m.group(1)))
        if int(m.group(1)) in unidades_validas:
            return unid, "RC.UN.", "Padrão RC.UN. confirmado"
        return None, "RC.UN.", f"RC.UN. inválido: {m.group(1)} não está na lista oficial"

    # 2. Padrão explícito UN.XXXXXX
    m = re.search(r"UN\.(\d{4,6})", historico)
    if m:
        unid = str(int(m.group(1)))
        if int(m.group(1)) in unidades_validas:
            return unid, "UN.explicito", "Padrão UN. confirmado"
        return None, "UN.explicito", f"UN. inválido: {m.group(1)} não está na lista oficial"

    # 3. Padrão numérico no final próximo de unidade conhecida
    # (apenas se estiver em contexto claro de cobrança)
    if "RC" in historico.upper() or "UN" in historico.upper():
        nums = re.findall(r"\d{3,4}", historico)
        for n in nums:
            val = int(n)
            if val in unidades_validas:
                return str(val), f"numerico_contexto_RC", f"Número {val} está na lista oficial e próximo de padrão RC"

    return None, "nao_detectado", "Nenhum padrão de unidade encontrado"


def analisar_origem_numero_invalido(
    numero: int,
    historico: str,
    codigo_conta: str | None,
    descricao_conta: str | None,
    valor: float | None,
) -> dict:
    """
    Analisa a origem provável de um número inválido detectado como pseudo-unidade.
    Retorna dict com análise.
    """
    resultado = {
        "origem_provavel": "desconhecida",
        "unidade_sugerida": None,
        "justificativa": "",
        "acao_recomendada": "revisar_parse",
    }

    # Verificar se é número de RB (referência bancária)
    m_rb = re.search(r"RB\.?(\d+)", historico)
    if m_rb and str(numero) in m_rb.group(1):
        resultado["origem_provavel"] = "numero_extraido_do_RB"
        resultado["justificativa"] = f"Número {numero} corresponde a parte do RB {m_rb.group(1)}"
        resultado["acao_recomendada"] = "ignorar_como_unidade"
        return resultado

    # Verificar se é número de RC
    m_rc = re.search(r"RC\.?(\d+)", historico)
    if m_rc and str(numero) in m_rc.group(1):
        resultado["origem_provavel"] = "numero_extraido_do_RC"
        resultado["justificativa"] = f"Número {numero} corresponde a parte do RC {m_rc.group(1)}"
        resultado["acao_recomendada"] = "ignorar_como_unidade"
        return resultado

    # Verificar se é código de conta
    if codigo_conta and str(numero) == str(codigo_conta).strip():
        resultado["origem_provavel"] = "codigo_contabil_confundido_com_unidade"
        resultado["justificativa"] = f"Número {numero} é o código da conta {descricao_conta}"
        resultado["acao_recomendada"] = "ignorar_como_unidade"
        return resultado

    # Verificar se é parte do valor
    if valor and valor > 0:
        valor_int = int(round(valor))
        if abs(valor_int - numero) < 5 or str(numero) in str(valor_int):
            resultado["origem_provavel"] = "valor_monetario_mal_interpretado"
            resultado["justificativa"] = f"Número {numero} próximo ao valor {valor:.2f}"
            resultado["acao_recomendada"] = "ignorar_como_unidade"
            return resultado

    # Verificar se é parcela/ref
    if "PARC" in historico.upper() or "PARCELA" in historico.upper():
        resultado["origem_provavel"] = "numero_de_parcela"
        resultado["justificativa"] = f"Número {numero} pode ser número de parcela no histórico"
        resultado["acao_recomendada"] = "revisar_parse"
        return resultado

    # Verificar transferência entre contas
    if "TRANSF" in historico.upper() or "TRANSP" in historico.upper():
        resultado["origem_provavel"] = "transferencia_entre_contas"
        resultado["justificativa"] = f"Histórico é de transferência, número {numero} pode ser referência"
        resultado["acao_recomendada"] = "ignorar_como_unidade"
        return resultado

    # Verificar se tem unidade válida no histórico próximo
    nums = re.findall(r"\d{3,4}", historico)
    for n in nums:
        val = int(n)
        m = re.search(rf"\b{n}\b", historico)
        if m and val != numero:
            resultado["unidade_sugerida"] = str(val)
            resultado["origem_provavel"] = "numero_proximo_de_unidade_valida"
            resultado["justificativa"] = f"Número {val} próximo a {numero} pode ser unidade correta"
            resultado["acao_recomendada"] = "validar_se_unidade_sugerida"
            return resultado

    resultado["justificativa"] = f"Número {numero} sem origem identificada no histórico"
    resultado["acao_recomendada"] = "analise_manual"
    return resultado


def validar_lancamentos_unidades(
    lancamentos: list[LancamentoExtrato],
    unidades_validas: set[int],
) -> tuple[list[LancamentoExtrato], pd.DataFrame]:
    """
    Valida unidades de todos os lançamentos contra lista oficial.
    Remove unidades inválidas e gera tabela de auditoria.

    Retorna (lancamentos_corrigidos, df_auditoria).
    """
    auditoria = []
    lancamentos_corrigidos = []

    for lanc in lancamentos:
        if lanc.unidade is None:
            lancamentos_corrigidos.append(lanc)
            continue

        try:
            uid = int(lanc.unidade)
        except (ValueError, TypeError):
            lancamentos_corrigidos.append(lanc)
            continue

        if uid in unidades_validas:
            lancamentos_corrigidos.append(lanc)
        else:
            # Unidade inválida: analisar origem
            analise = analisar_origem_numero_invalido(
                uid, lanc.historico_bruto, lanc.codigo_conta,
                lanc.descricao_conta, lanc.valor,
            )
            auditoria.append({
                "competencia": lanc.competencia,
                "numero_detectado": uid,
                "historico_bruto": lanc.historico_bruto,
                "codigo_conta": lanc.codigo_conta,
                "descricao_conta": lanc.descricao_conta,
                "valor": lanc.valor,
                "motivo_invalidacao": "unidade_nao_cadastrada",
                "origem_provavel": analise["origem_provavel"],
                "unidade_sugerida": analise["unidade_sugerida"],
                "acao_recomendada": analise["acao_recomendada"],
                "justificativa": analise["justificativa"],
            })
            # Remover unidade inválida, marcar como None
            lanc.unidade = None
            lanc.observacao_parse = f"unidade_invalida_{uid}"
            lancamentos_corrigidos.append(lanc)

    df_auditoria = pd.DataFrame(auditoria)
    if not df_auditoria.empty:
        df_auditoria = df_auditoria.sort_values(
            ["competencia", "numero_detectado"]
        ).reset_index(drop=True)

    logger.info(
        "Validação de unidades: %d válidas, %d inválidas na auditoria",
        len(lancamentos_corrigidos) - len(auditoria),
        len(auditoria),
    )
    return lancamentos_corrigidos, df_auditoria


def gerar_historico_unidades(
    lancamentos: list[LancamentoExtrato],
    unidades_validas: set[int],
) -> pd.DataFrame:
    """
    Gera histórico consolidado de unidades por competência.
    """
    from collections import defaultdict

    historico: dict[tuple[str, int], dict] = defaultdict(lambda: {
        "total_pago": 0.0, "total_debitos": 0.0,
        "apareceu": False, "tem_atraso": False,
        "tem_acordo": False, "data_pagamento": "",
    })

    for lanc in lancamentos:
        if lanc.unidade is None:
            continue
        try:
            uid = int(lanc.unidade)
        except (ValueError, TypeError):
            continue
        if uid not in unidades_validas:
            continue

        chave = (lanc.competencia, uid)
        h = historico[chave]
        h["apareceu"] = True

        if lanc.categoria in {"ACORDOS"}:
            h["tem_acordo"] = True
        if lanc.categoria in {"ATRASO_COTA_CONDOMINIAL", "ATRASO_BENFEITORIAS",
                               "ATRASO_FUNDO_RESERVA", "ATRASO_SALAO_FESTAS"}:
            h["tem_atraso"] = True

        if lanc.tipo == "credito" and lanc.valor:
            h["total_pago"] += lanc.valor
        elif lanc.tipo == "debito" and lanc.valor:
            h["total_debitos"] += lanc.valor

        if lanc.data_lancamento and not h["data_pagamento"]:
            h["data_pagamento"] = lanc.data_lancamento

    linhas = []
    comps = sorted(set(l.competencia for l in lancamentos if l.competencia))
    unids = sorted(unidades_validas)

    for uid in unids:
        for comp in comps:
            chave = (comp, uid)
            h = historico.get(chave)
            if h:
                linhas.append({
                    "unidade": str(uid),
                    "competencia": comp,
                    "total_pago": round(h["total_pago"], 2),
                    "total_debitos": round(h["total_debitos"], 2),
                    "data_pagamento": h["data_pagamento"],
                    "apareceu": "sim" if h["apareceu"] else "nao",
                    "tem_atraso": "sim" if h["tem_atraso"] else "nao",
                    "tem_acordo": "sim" if h["tem_acordo"] else "nao",
                    "status": _classificar_historico(h),
                })
            else:
                linhas.append({
                    "unidade": str(uid),
                    "competencia": comp,
                    "total_pago": 0.0,
                    "total_debitos": 0.0,
                    "data_pagamento": "",
                    "apareceu": "nao",
                    "tem_atraso": "nao",
                    "tem_acordo": "nao",
                    "status": "nao_localizado",
                })

    return pd.DataFrame(linhas)


def _classificar_historico(h: dict) -> str:
    """Classifica status de uma unidade em uma competência."""
    if h["tem_atraso"]:
        return "em_atraso"
    if h["tem_acordo"]:
        return "em_acordo"
    if h["apareceu"] and h["total_pago"] > 0:
        return "pago_no_mes"
    if h["apareceu"] and h["total_pago"] == 0:
        return "analise_manual"
    return "nao_localizado"
