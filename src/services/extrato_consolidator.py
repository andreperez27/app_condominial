import logging
from collections import defaultdict
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

import pandas as pd

from src.parser_extrato import ParserExtrato

logger = logging.getLogger(__name__)

# Categorias para o resumo mensal (modelo Planilha controle do condominio.xlsx)
CATEGORIAS_RESUMO = [
    "CONTA CONDOMINIO",
    "BENFEITORIAS",
    "FUNDO DE RESERVA",
    "SALAO DE FESTAS",
    "Mant. PINTURA",
    "APLICACAO BTG",
]

# Mapeamento de categorias normalizadas -> nome no resumo
CATEGORIA_PARA_RESUMO = {
    "COTA_CONDOMINIAL": "CONTA CONDOMINIO",
    "AGUA": "CONTA CONDOMINIO",
    "GAS": "CONTA CONDOMINIO",
    "MULTAS": "CONTA CONDOMINIO",
    "ACORDOS": "CONTA CONDOMINIO",
    "TAXA_ADMINISTRACAO": "CONTA CONDOMINIO",
    "RECEITAS_EVENTUAIS": "CONTA CONDOMINIO",
    "CONDOMINOS_EM_ATRASO": "CONTA CONDOMINIO",
    "ATRASO_COTA_CONDOMINIAL": "CONTA CONDOMINIO",
    "ATRASO_BENFEITORIAS": "CONTA CONDOMINIO",
    "ATRASO_FUNDO_RESERVA": "CONTA CONDOMINIO",
    "ATRASO_SALAO_FESTAS": "CONTA CONDOMINIO",
    "RECEITAS_FUNDO_RESERVA": "FUNDO DE RESERVA",
    "FUNDO_RESERVA": "FUNDO DE RESERVA",
    "BENFEITORIAS": "BENFEITORIAS",
    "SALAO_FESTAS": "SALAO DE FESTAS",
    "MANUT_PINTURA": "Mant. PINTURA",
    "CORRECAO_MONETARIA": "APLICACAO BTG",
    "RENDIMENTO_APLICACAO": "APLICACAO BTG",
    "APLICACAO_BTG": "APLICACAO BTG",
    "TRANSFERENCIA_ENTRE_CONTAS": "APLICACAO BTG",
    "MANUTENCOES_DIVERSAS": "CONTA CONDOMINIO",
}

# Contas de despesa que devem ser tratadas como débito no resumo
CONTAS_DESPESA = {
    1561, 1563, 1564, 1575, 1576, 1582, 1583, 1592, 1610, 1621, 1624,
    1634, 1641, 1644, 1647, 1648, 1649, 1653, 1662, 1667, 1677, 1681,
    1685, 1690, 1693, 1702, 1712, 1714, 1721, 1725, 1732, 1733, 1743,
    1747, 1770, 1772, 1803, 1818, 1820, 1821, 2501, 1024, 1099,
}


def _extrair_saldos(resumos: list) -> dict:
    """Extrai saldo_inicial e saldo_final da lista de ResumoMensal."""
    s = {"saldo_inicial": 0.0, "saldo_final": 0.0, "creditos": 0.0, "debitos": 0.0}
    for r in resumos:
        if r.saldo_inicial is not None:
            s["saldo_inicial"] = r.saldo_inicial
        if r.saldo_final is not None:
            s["saldo_final"] = r.saldo_final
        if r.creditos is not None:
            s["creditos"] = r.creditos
        if r.debitos is not None:
            s["debitos"] = r.debitos
    return s


def consolidar_extratos(
    input_dir: str | Path,
    mapa_contas: str | Path,
) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    """
    Lê todos os extratos renomeados e retorna:
      - df_lancamentos: DataFrame com todos os lançamentos consolidados
      - df_resumo: DataFrame com resumo mensal por categoria
      - auditoria: lista de dicts com problemas encontrados
    """
    dst = Path(input_dir)
    padroes = sorted(dst.glob("extrato_condominio_*.xls")) + sorted(dst.glob("extrato_condominio_*.xlsx"))
    # Dedup por competência (último vence)
    arquivos_por_comp: dict[str, Path] = {}
    for f in padroes:
        comp = f.stem.replace("extrato_condominio_", "").split("__")[0]
        arquivos_por_comp[comp] = f

    parser = ParserExtrato(str(mapa_contas))
    auditoria: list[dict] = []

    todos_lancamentos_raw = []
    todos_resumos_raw = []
    saldos_por_comp: dict[str, dict] = {}

    for comp in sorted(arquivos_por_comp):
        arquivo = arquivos_por_comp[comp]
        try:
            lancs, resums, nao_int = parser.parse(str(arquivo))
            saldos_por_comp[comp] = _extrair_saldos(resums)
            for l in lancs:
                l.arquivo_origem = arquivo.name
            todos_lancamentos_raw.extend(lancs)
            todos_resumos_raw.extend(resums)
            logger.info("OK: %s -> %d lancamentos", arquivo.name, len(lancs))
        except Exception as e:
            logger.error("Erro ao processar %s: %s", arquivo.name, e)
            auditoria.append({
                "arquivo_origem": arquivo.name,
                "tipo_problema": "erro_parse",
                "linha_origem": "",
                "observacao": str(e),
                "severidade": "erro",
            })

    # Montar DataFrame de lançamentos consolidados
    linhas = []
    for l in todos_lancamentos_raw:
        data_doc = getattr(l, "vencimento", None) or getattr(l, "data_documento", None)
        data_lanc = getattr(l, "data_lancamento", None)
        if isinstance(data_lanc, str):
            try:
                data_lanc = datetime.strptime(data_lanc, "%Y-%m-%d")
            except ValueError:
                pass
        if isinstance(data_doc, str):
            try:
                data_doc = datetime.strptime(data_doc, "%d/%m/%y")
            except ValueError:
                try:
                    data_doc = datetime.strptime(data_doc, "%d/%m/%Y")
                except ValueError:
                    data_doc = None

        valor = getattr(l, "valor", 0) or 0
        tipo = getattr(l, "tipo", "")
        debito = valor if tipo == "debito" else 0
        credito = valor if tipo == "credito" else 0
        valor_assinado = -valor if tipo == "debito" else valor

        saldos = saldos_por_comp.get(l.competencia, {})
        linhas.append({
            "competencia": l.competencia,
            "data_lancamento": data_lanc,
            "data_documento": data_doc,
            "codigo_conta": l.codigo_conta,
            "descricao_conta": l.descricao_conta,
            "historico": l.historico_bruto,
            "documento": getattr(l, "referencia_rb", ""),
            "tipo_movimento": tipo,
            "debito": debito,
            "credito": credito,
            "valor": round(valor_assinado, 2),
            "saldo_inicial_mes": saldos.get("saldo_inicial", 0),
            "saldo_final_mes": saldos.get("saldo_final", 0),
            "arquivo_origem": l.arquivo_origem,
            "nome_arquivo": l.arquivo_origem,
            "observacao_parse": getattr(l, "observacao_parse", ""),
            "categoria": getattr(l, "categoria", ""),
            "unidade": getattr(l, "unidade", ""),
        })

    df_lancamentos = pd.DataFrame(linhas)
    if not df_lancamentos.empty:
        df_lancamentos = df_lancamentos.sort_values(["competencia", "data_lancamento"]).reset_index(drop=True)

    # Montar resumo mensal por categoria
    resumo_linhas = []
    comps_ordenadas = sorted(saldos_por_comp.keys())

    for comp in comps_ordenadas:
        saldos = saldos_por_comp[comp]
        # Filtrar lançamentos desta competência
        df_comp = df_lancamentos[df_lancamentos["competencia"] == comp]

        # Agrupar por categoria de resumo
        creditos_por_cat: dict[str, float] = defaultdict(float)
        debitos_por_cat: dict[str, float] = defaultdict(float)

        for _, row in df_comp.iterrows():
            cat_normalizada = row.get("categoria", "")
            cat_resumo = CATEGORIA_PARA_RESUMO.get(cat_normalizada, "CONTA CONDOMINIO")
            v = row["valor"]
            if v > 0:
                creditos_por_cat[cat_resumo] += v
            elif v < 0:
                debitos_por_cat[cat_resumo] += abs(v)

        # Para APLICACAO BTG, separar aportes (débito na conta corrente = transferência p/ aplicação)
        # e juros (crédito de rendimento)
        # Mapear códigos de conta para detectar aportes vs juros
        aporte_btg = 0.0
        juros_btg = 0.0
        df_btg = df_comp[df_comp["categoria"].isin(["APLICACAO_BTG", "RENDIMENTO_APLICACAO", "CORRECAO_MONETARIA", "TRANSFERENCIA_ENTRE_CONTAS"])]
        for _, row in df_btg.iterrows():
            if row["tipo_movimento"] == "debito":
                aporte_btg += abs(row["valor"])
            elif row["tipo_movimento"] == "credito":
                juros_btg += abs(row["valor"])

        # Saldo anterior para cada categoria (herdado do mês anterior)
        # Para simplificar, usamos os saldos gerais do extrato
        saldo_anterior_geral = saldos.get("saldo_inicial", 0)
        saldo_final_geral = saldos.get("saldo_final", 0)
        creditos_geral = saldos.get("creditos", 0)
        debitos_geral = saldos.get("debitos", 0)

        # Linha SALDO FINAL
        total_creditos = sum(creditos_por_cat.values())
        total_debitos = sum(debitos_por_cat.values())

        for cat in CATEGORIAS_RESUMO:
            cred = round(creditos_por_cat.get(cat, 0), 2)
            deb = round(debitos_por_cat.get(cat, 0), 2)

            if cat == "APLICACAO BTG":
                # Saldo anterior = saldo da aplicação no mês anterior
                # (estimativa: saldo que estava aplicado)
                resumo_linhas.append({
                    "competencia": comp,
                    "descricao": f"{cat} - Aporte",
                    "saldo_anterior": 0,
                    "creditos": 0,
                    "debitos": round(aporte_btg, 2),
                    "saldo": round(-aporte_btg, 2),
                })
                resumo_linhas.append({
                    "competencia": comp,
                    "descricao": f"{cat} - Juros",
                    "saldo_anterior": 0,
                    "creditos": round(juros_btg, 2),
                    "debitos": 0,
                    "saldo": round(juros_btg, 2),
                })
                continue

            # Para categorias comuns: saldo = cred - deb
            saldo_cat = round(cred - deb, 2)
            resumo_linhas.append({
                "competencia": comp,
                "descricao": cat,
                "saldo_anterior": 0,  # viria do mês anterior se houvesse rastreabilidade
                "creditos": cred,
                "debitos": deb,
                "saldo": saldo_cat,
            })

        # Linha SALDO FINAL do mês
        resumo_linhas.append({
            "competencia": comp,
            "descricao": "SALDO FINAL",
            "saldo_anterior": round(saldo_anterior_geral, 2),
            "creditos": round(creditos_geral, 2),
            "debitos": round(debitos_geral, 2),
            "saldo": round(saldo_final_geral, 2),
        })

        # Validação: saldo_anterior + creditos_geral - debitos_geral == saldo_final_geral
        calc = round(saldo_anterior_geral + creditos_geral - debitos_geral, 2)
        if abs(calc - saldo_final_geral) > 0.01:
            auditoria.append({
                "arquivo_origem": f"extrato_{comp}.xls",
                "tipo_problema": "divergencia_totalizacao",
                "linha_origem": "",
                "observacao": f"{comp}: saldo_ant={saldo_anterior_geral} + cred={creditos_geral} - deb={debitos_geral} = {calc}, mas saldo_final={saldo_final_geral}",
                "severidade": "aviso",
            })

    df_resumo = pd.DataFrame(resumo_linhas)

    # Auditoria: detectar contas sem categoria
    if not df_lancamentos.empty:
        sem_cat = df_lancamentos[df_lancamentos["categoria"] == ""]
        if not sem_cat.empty:
            for _, row in sem_cat.iterrows():
                auditoria.append({
                    "arquivo_origem": row["arquivo_origem"],
                    "tipo_problema": "conta_sem_categoria",
                    "linha_origem": str(row.get("codigo_conta", "")),
                    "observacao": f"Conta {row['codigo_conta']} - {row['descricao_conta']} não mapeada",
                    "severidade": "aviso",
                })

    return df_lancamentos, df_resumo, auditoria
