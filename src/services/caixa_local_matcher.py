import logging
from collections import Counter
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz

from src.parsers.pdf_despesas_locais_parser import (
    LancamentoDespesaLocal,
    extrair_fornecedores_de_paginas,
    parse_pdf_despesas_locais,
)
from src.utils import parse_valor_bruto

logger = logging.getLogger(__name__)

STATUS_MATCH_EXATO = "MATCH_EXATO"
STATUS_MATCH_VALOR_DATA = "MATCH_POR_VALOR_DATA"
STATUS_MATCH_PROVAVEL = "MATCH_PROVAVEL"
STATUS_NAO_ENCONTRADO = "NAO_ENCONTRADO"
STATUS_DUPLICADO = "DUPLICADO"
STATUS_AMBIGUO = "AMBIGUO"


def _normalizar_valor(v: float | None) -> float | None:
    if v is None:
        return None
    return round(v, 2)


def _extrair_data_iso(data_str: str | None) -> str:
    if not data_str:
        return ""
    partes = data_str.split("/")
    if len(partes) == 3:
        return f"{partes[2]}-{partes[1]}-{partes[0]}"
    return data_str


def processar_pdfs_despesas_locais(
    pdf_dir: str | Path,
) -> dict[str, dict]:
    """
    Lê todos os PDFs de despesas locais em pdf_dir.
    Retorna dict: {competencia: {"lancamentos": [...], "pdf_path": str}}
    """
    dst = Path(pdf_dir)
    pdfs = sorted(dst.glob("despesas_locais_*.pdf"))
    resultado = {}

    for pdf_path in pdfs:
        lancamentos, comp, periodo_ini, periodo_fim = parse_pdf_despesas_locais(str(pdf_path))
        if not comp or not lancamentos:
            logger.warning("PDF sem dados válidos: %s", pdf_path.name)
            continue
        lancamentos = extrair_fornecedores_de_paginas(lancamentos, str(pdf_path))
        resultado[comp] = {"lancamentos": lancamentos, "pdf_path": str(pdf_path)}
        logger.info("PDF %s: %d lançamentos (comp=%s)", pdf_path.name, len(lancamentos), comp)

    return resultado


def conciliar_caixa_local(
    lancamentos_extrato: list,
    pdf_data: dict[str, dict],
) -> pd.DataFrame:
    """
    Cruza lançamentos do extrato (conta 1723) com dados do PDF despesas locais.
    Retorna DataFrame com a conciliação detalhada.
    """
    linhas = []

    # Agrupa lançamentos do extrato por competência
    extrato_por_comp: dict[str, list] = {}
    for l in lancamentos_extrato:
        if l.codigo_conta == "1723" and l.valor is not None and l.valor > 0:
            comp = l.competencia or ""
            if comp not in extrato_por_comp:
                extrato_por_comp[comp] = []
            extrato_por_comp[comp].append(l)

    for comp, extratos in extrato_por_comp.items():
        pdf_lancs = pdf_data.get(comp, {}).get("lancamentos", []) if comp in pdf_data else []

        # Separa devedores (despesas) e credores (remessas) do PDF
        debitos_pdf = [p for p in pdf_lancs if p.tipo_movimento == "debito" and p.valor_debito and p.valor_debito > 0]
        creditos_pdf = [p for p in pdf_lancs if p.tipo_movimento == "credito" and p.valor_credito and p.valor_credito > 0]

        # Agrupar PDF por valor para detectar duplicatas
        pdf_valor_count = Counter()
        for p in debitos_pdf:
            v = _normalizar_valor(p.valor_debito)
            if v is not None:
                pdf_valor_count[v] += 1

        # Match cada extrato com PDF
        usados_pdf = set()

        for e in extratos:
            data_e = _extrair_data_iso(e.data_lancamento) if hasattr(e, "data_lancamento") and e.data_lancamento else ""
            valor_e = _normalizar_valor(e.valor)
            hist_e = e.historico_bruto or ""

            melhor_match = None
            melhor_score = -1

            for idx_p, p in enumerate(debitos_pdf):
                if idx_p in usados_pdf:
                    continue
                data_p = _extrair_data_iso(p.data)
                valor_p = _normalizar_valor(p.valor_debito)

                if valor_e is not None and valor_p is not None and abs(valor_e - valor_p) < 0.01:
                    # Se há múltiplos PDF com mesmo valor, exigir data igual
                    if pdf_valor_count.get(valor_p, 0) > 1:
                        if data_e and data_p and data_e == data_p:
                            score = 200
                        else:
                            continue  # pula candidatos sem data match quando há duplicatas
                    else:
                        score = 100
                        if data_e and data_p and data_e == data_p:
                            score = 200

                    # Similaridade textual
                    txt_score = fuzz.token_sort_ratio(hist_e.upper(), p.historico_pdf.upper())
                    score += txt_score

                    if score > melhor_score:
                        melhor_score = score
                        melhor_match = (idx_p, p)

            if melhor_match:
                idx_p, p = melhor_match
                usados_pdf.add(idx_p)

                if melhor_score >= 200:
                    status = STATUS_MATCH_EXATO
                elif melhor_score >= 100:
                    status = STATUS_MATCH_VALOR_DATA
                else:
                    status = STATUS_MATCH_PROVAVEL

                desc_final = f"CAIXA LOCAL - {p.historico_pdf}"
            else:
                status = STATUS_NAO_ENCONTRADO
                p = None
                desc_final = hist_e

            subcat = _classificar_subcategoria(p.historico_pdf if p else hist_e)

            linhas.append({
                "competencia": comp,
                "data": _extrair_data_iso(e.data_lancamento) if hasattr(e, "data_lancamento") and e.data_lancamento else "",
                "valor": valor_e,
                "historico_pdf": p.historico_pdf if p else "",
                "descricao_extrato_original": hist_e,
                "descricao_final_conciliada": desc_final,
                "status_match_pdf": status,
                "subcategoria": subcat,
                "fornecedor_sugerido": p.fornecedor_sugerido if p else None,
                "documento_sugerido": p.documento_sugerido if p else None,
                "cnpj": p.cnpj if p else None,
                "nf_numero": p.nf_numero if p else None,
            })

        # Itens do PDF sem match no extrato
        for idx_p, p in enumerate(debitos_pdf):
            if idx_p in usados_pdf:
                continue
            subcat = _classificar_subcategoria(p.historico_pdf)
            linhas.append({
                "competencia": comp,
                "data": _extrair_data_iso(p.data),
                "valor": _normalizar_valor(p.valor_debito),
                "historico_pdf": p.historico_pdf,
                "descricao_extrato_original": "",
                "descricao_final_conciliada": f"CAIXA LOCAL (somente PDF) - {p.historico_pdf}",
                "status_match_pdf": STATUS_NAO_ENCONTRADO,
                "subcategoria": subcat,
                "fornecedor_sugerido": p.fornecedor_sugerido,
                "documento_sugerido": p.documento_sugerido,
                "cnpj": p.cnpj,
                "nf_numero": p.nf_numero,
            })

    df = pd.DataFrame(linhas)
    if not df.empty:
        df = df.sort_values(["competencia", "data", "valor"]).reset_index(drop=True)
    return df


def _classificar_subcategoria(texto: str) -> str:
    t = texto.upper()
    if "PORTARIA" in t or "PRATELEIRA" in t:
        return "portaria"
    if "PISCINA" in t or "VIDRO" in t or "MESA" in t:
        return "piscina"
    if "TELA" in t or "EXAUST" in t or "ISOLAMENTO" in t:
        return "exaustao"
    if "DISPENSER" in t or "COPO" in t or "AGUA" in t:
        return "copa"
    if "LIMPEZA" in t or "TANQUINHO" in t:
        return "limpeza"
    if "JARDINAGEM" in t or "JARDIM" in t or "PLANTA" in t:
        return "jardinagem"
    if "BOMBA" in t or "BOLA" in t:
        return "lazer"
    if "MANUT" in t or "CONSERTO" in t or "REPARO" in t or "OBRA" in t:
        return "manutencao"
    if "MATERIAL" in t or "FERRAMENTA" in t:
        return "materiais"
    if "CAFE" in t or "ACUCAR" in t or "ALIMENT" in t:
        return "despensa"
    return "outros"


def gerar_resumo_caixa_local(df_conciliado: pd.DataFrame) -> pd.DataFrame:
    """Gera resumo por competência da conciliação de caixa local."""
    if df_conciliado.empty:
        return pd.DataFrame()

    resumo = df_conciliado.groupby("competencia").agg(
        total_gastos_pdf=("valor", lambda v: v[df_conciliado.loc[v.index, "status_match_pdf"] != STATUS_NAO_ENCONTRADO].sum()),
        total_extrato=("valor", lambda v: v[df_conciliado.loc[v.index, "descricao_extrato_original"] != ""].sum()),
        itens_conciliados=("status_match_pdf", lambda s: s.isin([STATUS_MATCH_EXATO, STATUS_MATCH_VALOR_DATA, STATUS_MATCH_PROVAVEL]).sum()),
        itens_nao_conciliados=("status_match_pdf", lambda s: (s == STATUS_NAO_ENCONTRADO).sum()),
        qtde_total=("status_match_pdf", "count"),
    ).reset_index()

    resumo["diferenca"] = resumo["total_gastos_pdf"] - resumo["total_extrato"]
    return resumo
