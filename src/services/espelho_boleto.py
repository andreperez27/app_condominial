import logging
from collections import defaultdict, Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

from src.models import LancamentoExtrato, RecebivelUnidade

try:
    from src.utils import ensure_columns
except ImportError:
    def ensure_columns(df, *args, **kwargs):
        return df

logger = logging.getLogger(__name__)

# ── Verbete categories that belong in a boleto ──
BOLETO_CATEGORIES = {
    "COTA_CONDOMINIAL",
    "AGUA",
    "GAS",
    "FUNDO_RESERVA",
    "BENFEITORIAS",
    "SALAO_FESTAS",
    "MANUT_PINTURA",
    "MULTAS",
    "ACORDOS",
    "TAXA_ADMINISTRACAO",
    "MANUTENCOES_DIVERSAS",
}

# Categories that must NEVER appear in a boleto (non-unit items)
EXCLUDED_CATEGORIES = {
    "TRANSFERENCIA_ENTRE_CONTAS",
    "APLICACAO_BTG",
    "RENDIMENTO_APLICACAO",
    "CORRECAO_MONETARIA",
    "RECEITAS_EVENTUAIS",
    "RECEITAS_FUNDO_RESERVA",
    "PORTA_AUTOMATICA_RECEITAS",
    "TROCA_TUBULACAO_GAS",
    "DIVERSOS",
    "ESTORNO",
    "TRIBUTOS_FOLHA",
    "CAIXA_LOCAL",
    "ATRASO_COTA_CONDOMINIAL",
    "ATRASO_BENFEITORIAS",
    "ATRASO_FUNDO_RESERVA",
    "ATRASO_SALAO_FESTAS",
    "CONDOMINOS_EM_ATRASO",
}

COD_VERBAS = {
    "1002": "Cota Condominial",
    "1010": "Consumo de Gás",
    "1052": "Consumo de Água",
    "1050": "Manutenções Diversas",
    "15002": "Manut. Pintura PC",
    "3002": "Fundo de Reserva",
    "1102": "Taxa de Administração",
    "2002": "Benfeitorias",
    "19002": "Salao de Festas",
    "1098": "Multas",
    "2998": "Multas",
    "3998": "Multas",
    "15998": "Multas",
    "19998": "Multas",
    "1006": "Acordos",
    "2005": "Acordos",
    "3005": "Acordos",
    "15005": "Acordos",
}

OBS_PADRAO = {
    "COTA_CONDOMINIAL": "Despesa fixa mensal",
    "AGUA": "Consumo do período",
    "GAS": "Consumo do período",
    "TAXA_ADMINISTRACAO": "Administração do condomínio",
    "FUNDO_RESERVA": "Poupança para futuras reformas",
    "BENFEITORIAS": "Benfeitoria em andamento",
    "MANUT_PINTURA": "Manutenção programada",
    "SALAO_FESTAS": "Reserva de área comum",
    "MULTAS": "Penalidade por atraso",
    "ACORDOS": "Acordo de parcelamento",
    "MANUTENCOES_DIVERSAS": "Manutenções diversas",
}

CATEGORIA_POR_CODIGO = {
    "1002": "COTA_CONDOMINIAL",
    "1010": "GAS",
    "1052": "AGUA",
    "1050": "MANUTENCOES_DIVERSAS",
    "15002": "MANUT_PINTURA",
    "3002": "FUNDO_RESERVA",
    "1102": "TAXA_ADMINISTRACAO",
    "2002": "BENFEITORIAS",
    "19002": "SALAO_FESTAS",
    "1098": "MULTAS",
    "2998": "MULTAS",
    "3998": "MULTAS",
    "15998": "MULTAS",
    "19998": "MULTAS",
    "1006": "ACORDOS",
    "2005": "ACORDOS",
    "3005": "ACORDOS",
    "15005": "ACORDOS",
}


def _parse_data_br(data_str: str | None) -> str | None:
    """Converte string para YYYY-MM-DD.
    Aceita formatos: DD/MM/YYYY, DD/MM/YY, YYYY-MM-DD, YYYY/MM/DD.
    Retorna None se inválido."""
    if not data_str or not data_str.strip():
        return None
    s = data_str.strip()

    formatos = ["%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%Y/%m/%d"]
    for fmt in formatos:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def _extrair_ref_do_historico(historico: str) -> str | None:
    """Tenta extrair mês/ano de referência do histórico."""
    import re
    m = re.search(r"(?:REF(?:ERENCIA)?[:\s]*)?(\d{2})/(\d{4})", historico, re.IGNORECASE)
    if m:
        return f"{m.group(2)}-{m.group(1)}"
    m = re.search(r"(\d{4})[-/](\d{2})", historico)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return None


def _extrair_vencimento_do_historico(historico: str) -> str | None:
    """Tenta extrair data de vencimento do histórico (VC.DD/MM/AA ou VC.DD/MM/AAAA)."""
    import re
    m = re.search(r"VC\.(\d{2}/\d{2}/\d{2,4})", historico)
    if m:
        return _parse_data_br(m.group(1))
    return None


def _determinar_competencia_referencia(
    data_lancamento: str | None,
    vencimento: str | None,
    historico_bruto: str,
    competencia_pagamento: str,
) -> tuple[str, str]:
    """
    Determina a competência de referência (mês do boleto).
    Retorna (competencia_referencia, observacao).
    """
    obs = ""
    venc_iso = _parse_data_br(vencimento) if vencimento else None
    data_iso = _parse_data_br(data_lancamento) if data_lancamento else None

    # 1. Vencimento do boleto
    if venc_iso:
        mes_ref = venc_iso[:7]
        return mes_ref, f"Vencimento {venc_iso}"

    # 1b. Tentar extrair vencimento do histórico (VC.XX/XX/XX)
    venc_hist = _extrair_vencimento_do_historico(historico_bruto)
    if venc_hist:
        mes_ref = venc_hist[:7]
        return mes_ref, f"Venc. histórico {venc_hist}"

    # 2. Ref explícita no histórico
    ref_hist = _extrair_ref_do_historico(historico_bruto)
    if ref_hist:
        return ref_hist, f"Ref. histórica {ref_hist}"

    # 3. Regra do condomínio: pagamento até dia 10 do mês subsequente = mês anterior
    if data_iso:
        dia = int(data_iso[8:10])
        mes = int(data_iso[5:7])
        ano = int(data_iso[:4])
        if dia <= 10:
            mes_ref = mes - 1
            ano_ref = ano
            if mes_ref == 0:
                mes_ref = 12
                ano_ref -= 1
            comp_ref = f"{ano_ref:04d}-{mes_ref:02d}"
            return comp_ref, f"Pagamento dia {dia} (regra dia 10)"

    # 4. Fallback: usar a própria competência do pagamento
    return competencia_pagamento, "DUVIDOSO - sem referência clara"


def _criar_chave_dedup(row: dict) -> str:
    """Cria chave de deduplicação para um item do espelho."""
    return "|".join([
        str(row.get("competencia_referencia", "")),
        str(row.get("unidade", "")),
        str(row.get("codigo", "")),
        str(row.get("data_pagamento", "")),
        str(f"{row.get('valor', 0):.2f}"),
        str(row.get("historico_bruto", ""))[:50],
        str(row.get("arquivo_origem", "")),
    ])


def _classificar_status_item(
    valor_encontrado: float,
    valor_esperado: float,
    encontrou: bool,
) -> str:
    if not encontrou:
        return "nao_localizado"
    if abs(valor_encontrado - valor_esperado) < 0.01:
        return "pago"
    elif valor_encontrado > valor_esperado:
        return "pago_excedente"
    elif valor_encontrado > 0:
        return "parcial"
    return "nao_localizado"


def _classificar_status_competencia(
    total_esperado: float,
    total_pago: float,
    itens_nao_encontrados: int,
    itens_parciais: int,
) -> str:
    if total_esperado == 0:
        return "sem_lancamentos"
    if itens_nao_encontrados > 0 and total_pago == 0:
        return "inadimplente_presumido"
    if total_pago >= total_esperado:
        if itens_nao_encontrados == 0 and itens_parciais == 0:
            return "pago_no_mes"
        else:
            return "pago_com_divergencia"
    if itens_nao_encontrados > 0 or itens_parciais > 0:
        return "parcial"
    return "nao_localizado"


def _calcular_mediana_historica(
    df_espelho: pd.DataFrame,
    unidade: str,
    codigo: str,
    comp_ref_atual: str,
) -> float:
    """Calcula mediana dos valores pagos para (unidade, codigo) excluindo competência atual."""
    hist = df_espelho[
        (df_espelho["unidade"] == unidade)
        & (df_espelho["codigo"] == codigo)
        & (df_espelho["competencia_referencia"] != comp_ref_atual)
    ]
    if hist.empty:
        return 0.0
    return hist["valor_pago"].median()


def _calcular_confiabilidade(audit_rows: list[dict]) -> str:
    """Alta = sem flags nos itens que entraram, Média = suspeitas menores, Baixa = suspeitas em itens incluídos."""
    entrou = [r for r in audit_rows if r.get("entrou_no_boleto") == "sim"]
    for r in entrou:
        if r.get("flag_competencia_suspeita") == "sim":
            return "baixa"
        if r.get("flag_unidade_suspeita") == "sim":
            return "baixa"

    qualquer = [r for r in audit_rows if r.get("entrou_no_boleto") == "nao"]
    for r in qualquer:
        if r.get("flag_competencia_suspeita") == "sim":
            return "media"
        if r.get("flag_unidade_suspeita") == "sim":
            return "media"
    return "alta"


def gerar_espelho_boleto(
    detalhe_recebiveis: pd.DataFrame,
    unidades_esperadas: list[str] | None = None,
) -> pd.DataFrame:
    """
    Gera espelho do boleto por unidade e competência de referência.
    Inclui auditoria completa e detecção de discrepâncias.
    """
    if detalhe_recebiveis.empty:
        return pd.DataFrame()

    df = detalhe_recebiveis.copy()

    for col in ["data_lancamento", "vencimento", "referencia_rb", "historico_bruto"]:
        if col not in df.columns:
            df[col] = ""

    # ── 1. Filter non-boleto categories ──
    df = ensure_columns(
        df,
        ["competencia", "unidade", "categoria", "valor", "codigo_conta",
         "descricao_conta", "historico_bruto", "arquivo_origem",
         "data_lancamento", "referencia_rb", "vencimento"],
        stage="espelho_boleto_pre_filtro",
    )
    mask_boleto = df["categoria"].isin(BOLETO_CATEGORIES)
    df_excluded = df[~mask_boleto].copy()
    df = df[mask_boleto].copy()

    logger.info(
        "Filtro categorias: %d boleto + %d excluídas (não fazem parte do boleto)",
        len(df), len(df_excluded),
    )

    linhas_espelho = []
    auditoria_linhas = []

    # ── 2. Group by competência + unidade + codigo_conta for dedup ──
    # First pass: build raw espelho rows
    for _, row in df.iterrows():
        comp_pagto = str(row.get("competencia", ""))
        data_lanc = str(row.get("data_lancamento", ""))
        venc = str(row.get("vencimento", "")) if pd.notna(row.get("vencimento")) else ""
        hist = str(row.get("historico_bruto", ""))
        unidade = str(row.get("unidade", ""))
        categoria = str(row.get("categoria", ""))
        valor = float(row.get("valor", 0))
        cod_conta = str(row.get("codigo_conta", ""))
        arquivo = str(row.get("arquivo_origem", ""))

        comp_ref, obs = _determinar_competencia_referencia(
            data_lanc, venc, hist, comp_pagto
        )

        descricao_verba = COD_VERBAS.get(cod_conta, "")
        if not descricao_verba:
            descricao_verba = f"{cod_conta} - {row.get('descricao_conta', '')}"

        obs_padrao = OBS_PADRAO.get(categoria, "")
        data_pagto_iso = _parse_data_br(data_lanc) or ""

        row_data = {
            "competencia_referencia": comp_ref,
            "unidade": unidade,
            "codigo": cod_conta,
            "descricao_da_verba": descricao_verba,
            "categoria": categoria,
            "valor": valor,
            "data_pagamento": data_pagto_iso,
            "competencia_pagamento": comp_pagto,
            "historico_bruto": hist,
            "arquivo_origem": arquivo,
            "observacao_competencia": obs,
            "observacoes": obs_padrao,
        }
        linhas_espelho.append(row_data)

    # ── 3. Dedup ──
    vistos: dict[str, int] = {}
    linhas_dedup = []
    for item in linhas_espelho:
        chave = _criar_chave_dedup(item)
        if chave in vistos:
            vistos[chave] += 1
            item["flag_duplicado"] = "sim"
        else:
            vistos[chave] = 1
            item["flag_duplicado"] = "nao"
        linhas_dedup.append(item)

    duplicatas_encontradas = [r for r in linhas_dedup if r["flag_duplicado"] == "sim"]
    if duplicatas_encontradas:
        logger.warning(
            "Dedup: %d linhas duplicadas removidas", len(duplicatas_encontradas)
        )

    # ── 4. Build audit rows ──
    for item in linhas_dedup:
        comp_suspeita = "nao"
        if "DUVIDOSO" in item.get("observacao_competencia", ""):
            comp_suspeita = "sim"

        unid_suspeita = "nao"
        try:
            uid = int(item["unidade"])
            if uid < 1 or uid > 999:
                unid_suspeita = "sim"
        except (ValueError, TypeError):
            unid_suspeita = "sim"

        entrou = "nao" if item["flag_duplicado"] == "sim" else "sim"
        motivo_exclusao = ""
        motivo_inclusao = ""
        if item["flag_duplicado"] == "sim":
            motivo_exclusao = "duplicata_detectada"
        elif entrou == "sim":
            motivo_inclusao = "categoria_boleto"

        auditoria_linhas.append({
            "unidade": item["unidade"],
            "competencia_referencia": item["competencia_referencia"],
            "codigo_conta": item["codigo"],
            "descricao_conta": item["descricao_da_verba"],
            "valor_item": item["valor"],
            "data_pagamento": item["data_pagamento"],
            "historico_bruto": item["historico_bruto"],
            "arquivo_origem": item["arquivo_origem"],
            "entrou_no_boleto": entrou,
            "motivo_inclusao": motivo_inclusao,
            "motivo_exclusao": motivo_exclusao,
            "flag_duplicado": item["flag_duplicado"],
            "flag_competencia_suspeita": comp_suspeita,
            "flag_unidade_suspeita": unid_suspeita,
        })

    # ── 5. Build final espelho DataFrame (only dedup'd items that entered) ──
    espelho_rows = []
    for item in linhas_dedup:
        if item["flag_duplicado"] == "sim":
            continue
        espelho_rows.append({
            "competencia_referencia": item["competencia_referencia"],
            "unidade": item["unidade"],
            "codigo": item["codigo"],
            "descricao_da_verba": item["descricao_da_verba"],
            "ref": item["competencia_referencia"],
            "valor_esperado": item["valor"],
            "valor_pago": item["valor"],
            "categoria": item["categoria"],
            "observacoes": item["observacoes"],
            "data_pagamento": item["data_pagamento"],
            "competencia_pagamento": item["competencia_pagamento"],
            "status_item": "pago",
            "status_competencia": "",
        })

    df_espelho = pd.DataFrame(espelho_rows)

    if df_espelho.empty:
        return pd.DataFrame()

    # ── 6. Aggregate (same codigo within same comp+unidade -> sum) ──
    agg = df_espelho.groupby(
        ["competencia_referencia", "unidade", "codigo", "descricao_da_verba",
         "categoria", "observacoes"]
    ).agg(
        valor_esperado=("valor_esperado", "sum"),
        valor_pago=("valor_pago", "sum"),
        data_pagamento=("data_pagamento", "first"),
        competencia_pagamento=("competencia_pagamento", "first"),
        ref=("ref", "first"),
    ).reset_index()

    # ── 7. Classify items ──
    agg["status_item"] = agg.apply(
        lambda r: _classificar_status_item(r["valor_pago"], r["valor_esperado"], True),
        axis=1,
    )

    # ── 8. Competência-level status ──
    for (comp_ref, unid), grupo in agg.groupby(["competencia_referencia", "unidade"]):
        total_esp = grupo["valor_esperado"].sum()
        total_pago = grupo["valor_pago"].sum()
        nao_encont = len(grupo[grupo["status_item"] == "nao_localizado"])
        parciais = len(grupo[grupo["status_item"] == "parcial"])
        status_comp = _classificar_status_competencia(total_esp, total_pago, nao_encont, parciais)
        agg.loc[
            (agg["competencia_referencia"] == comp_ref) & (agg["unidade"] == unid),
            "status_competencia"
        ] = status_comp

    # ── 9. Discrepancy detection (compare to historical median) ──
    discrepant_rows = []
    for _, r in agg.iterrows():
        mediana = _calcular_mediana_historica(
            agg, r["unidade"], r["codigo"], r["competencia_referencia"]
        )
        if mediana > 0:
            ratio = r["valor_pago"] / mediana
            if ratio > 1.5 or ratio < 0.5:
                discrepant_rows.append({
                    "unidade": r["unidade"],
                    "competencia_referencia": r["competencia_referencia"],
                    "codigo": r["codigo"],
                    "descricao": r["descricao_da_verba"],
                    "valor_atual": r["valor_pago"],
                    "mediana_historica": round(mediana, 2),
                    "razao": round(ratio, 2),
                    "classificacao": "VALOR_DISCREPANTE" if ratio > 2.0 or ratio < 0.3 else "NECESSITA_REVISAO",
                })

    df_espelho_final = agg.sort_values(
        ["competencia_referencia", "unidade", "codigo"]
    ).reset_index(drop=True)

    # Store audit/discrepancy data as attributes for downstream retrieval
    df_espelho_final.attrs["auditoria"] = auditoria_linhas
    df_espelho_final.attrs["itens_excluidos"] = df_excluded.to_dict("records") if not df_excluded.empty else []
    df_espelho_final.attrs["duplicatas"] = duplicatas_encontradas
    df_espelho_final.attrs["discrepancias"] = discrepant_rows

    # ── 10. Reliability indicator ──
    df_espelho_final.attrs["confiabilidade"] = _calcular_confiabilidade(auditoria_linhas)

    logger.info(
        "Espelho boleto gerado: %d linhas para %d unidades, confiabilidade=%s",
        len(df_espelho_final),
        agg["unidade"].nunique(),
        df_espelho_final.attrs["confiabilidade"],
    )

    return df_espelho_final


def gerar_espelho_consolidado_por_unidade(
    df_espelho: pd.DataFrame,
) -> pd.DataFrame:
    """Visão consolidada (1 linha por unidade/competência)."""
    if df_espelho.empty:
        return pd.DataFrame()

    consol = df_espelho.groupby(
        ["competencia_referencia", "unidade"]
    ).agg(
        total_esperado=("valor_esperado", "sum"),
        total_pago=("valor_pago", "sum"),
        data_pagamento=("data_pagamento", lambda x: "; ".join(sorted(set(str(d) for d in x if d)))),
        status_competencia=("status_competencia", "first"),
    ).reset_index()

    consol["diferenca"] = (consol["total_pago"] - consol["total_esperado"]).round(2)

    status_order = {
        "pago_no_mes": 0,
        "pago_antecipado": 1,
        "pago_com_divergencia": 2,
        "parcial": 3,
        "em_acordo": 4,
        "inadimplente_presumido": 5,
        "nao_localizado": 6,
        "sem_lancamentos": 7,
    }
    consol["_status_order"] = consol["status_competencia"].map(status_order).fillna(99)
    consol = consol.sort_values(
        ["unidade", "_status_order", "competencia_referencia"]
    ).drop(columns=["_status_order"])

    # Forward audit attrs
    if df_espelho.attrs:
        consol.attrs["auditoria"] = df_espelho.attrs.get("auditoria", [])
        consol.attrs["itens_excluidos"] = df_espelho.attrs.get("itens_excluidos", [])
        consol.attrs["duplicatas"] = df_espelho.attrs.get("duplicatas", [])
        consol.attrs["discrepancias"] = df_espelho.attrs.get("discrepancias", [])
        consol.attrs["confiabilidade"] = df_espelho.attrs.get("confiabilidade", "alta")

    return consol
