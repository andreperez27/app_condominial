import logging
from collections import defaultdict
from pathlib import Path

import pandas as pd

from src.models import LancamentoExtrato, RecebivelUnidade, ResumoMensal, LinhaNaoInterpretada
from src.utils import load_mapa_contas

logger = logging.getLogger(__name__)


def consolidar_resumo_mensal(
    resumos: list[ResumoMensal],
) -> pd.DataFrame:
    if not resumos:
        return pd.DataFrame()

    registro_por_chave: dict[tuple[str, str], dict] = {}

    for r in resumos:
        chave = (r.competencia, r.conta_normalizada or r.conta_origem)
        if chave not in registro_por_chave:
            registro_por_chave[chave] = {
                "competencia": r.competencia,
                "conta_origem": r.conta_origem,
                "conta_normalizada": r.conta_normalizada or r.conta_origem,
                "saldo_inicial": None,
                "creditos": None,
                "debitos": None,
                "saldo_final": None,
                "arquivo_origem": r.arquivo_origem,
                "observacao": "",
            }
        reg = registro_por_chave[chave]
        if r.saldo_inicial is not None:
            reg["saldo_inicial"] = r.saldo_inicial
        if r.creditos is not None:
            reg["creditos"] = r.creditos
        if r.debitos is not None:
            reg["debitos"] = r.debitos
        if r.saldo_final is not None:
            reg["saldo_final"] = r.saldo_final
        if r.observacao:
            if reg["observacao"]:
                reg["observacao"] += "; " + r.observacao
            else:
                reg["observacao"] = r.observacao

    df = pd.DataFrame(list(registro_por_chave.values()))
    df = df.sort_values(["competencia", "conta_normalizada"]).reset_index(drop=True)
    logger.info("Resumo mensal consolidado: %d linhas", len(df))
    return df


def separar_recebiveis_por_unidade(
    lancamentos: list[LancamentoExtrato],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    recebiveis: list[RecebivelUnidade] = []

    for lanc in lancamentos:
        if lanc.tipo != "credito":
            continue
        if not lanc.unidade:
            continue
        if not lanc.categoria:
            continue

        recebiveis.append(
            RecebivelUnidade(
                competencia=lanc.competencia,
                unidade=lanc.unidade,
                categoria=lanc.categoria,
                valor=lanc.valor,
                data_lancamento=lanc.data_lancamento,
                codigo_conta=lanc.codigo_conta,
                descricao_conta=lanc.descricao_conta,
                referencia_rb=lanc.referencia_rb,
                vencimento=lanc.vencimento,
                historico_bruto=lanc.historico_bruto,
                arquivo_origem=lanc.arquivo_origem,
            )
        )

    if not recebiveis:
        return pd.DataFrame(), pd.DataFrame()

    df_detalhe = pd.DataFrame([vars(r) for r in recebiveis])

    pivot = df_detalhe.pivot_table(
        index=["competencia", "unidade"],
        columns="categoria",
        values="valor",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    col_order = [
        "competencia", "unidade",
        "COTA_CONDOMINIAL", "AGUA", "GAS",
        "FUNDO_RESERVA", "BENFEITORIAS", "SALAO_FESTAS",
        "MULTAS", "ACORDOS", "MANUT_PINTURA", "MANUTENCOES_DIVERSAS",
        "TAXA_ADMINISTRACAO", "CORRECAO_MONETARIA", "RECEITAS_EVENTUAIS",
        "RECEITAS_FUNDO_RESERVA", "TRANSFERENCIA_ENTRE_CONTAS",
        "RENDIMENTO_APLICACAO", "APLICACAO_BTG", "DIVERSOS",
        "PORTA_AUTOMATICA_RECEITAS", "ATRASO_COTA_CONDOMINIAL",
        "ATRASO_BENFEITORIAS", "ATRASO_FUNDO_RESERVA", "ATRASO_SALAO_FESTAS",
        "TROCA_TUBULACAO_GAS",
    ]

    for col in col_order:
        if col not in pivot.columns:
            pivot[col] = 0.0

    num_cols = pivot.select_dtypes(include="number").columns
    pivot["total_pago"] = pivot[num_cols].sum(axis=1)

    pivot = pivot[[c for c in col_order if c in pivot.columns] + ["total_pago"]]

    logger.info("Recebiveis por unidade: %d linhas detalhe, %d pivotadas", len(df_detalhe), len(pivot))
    return df_detalhe, pivot


def gerar_lancamentos_normalizados(
    lancamentos: list[LancamentoExtrato],
) -> pd.DataFrame:
    if not lancamentos:
        return pd.DataFrame()
    df = pd.DataFrame([vars(l) for l in lancamentos])
    cols = [
        "competencia", "data_lancamento", "unidade", "codigo_conta",
        "descricao_conta", "categoria", "historico_bruto", "referencia_rb",
        "vencimento", "valor", "tipo", "arquivo_origem", "linha_origem",
        "observacao_parse",
    ]
    df = df[[c for c in cols if c in df.columns]]
    logger.info("Lancamentos normalizados: %d linhas", len(df))
    return df


def gerar_resumo_mensal_por_conta(
    lancamentos: list[LancamentoExtrato],
) -> list[ResumoMensal]:
    if not lancamentos:
        return []

    aggr: dict[tuple[str, str], dict] = {}
    for lanc in lancamentos:
        cat = lanc.categoria or "SEM_CATEGORIA"
        chave = (lanc.competencia, cat)
        if chave not in aggr:
            aggr[chave] = {"creditos": 0.0, "debitos": 0.0,
                           "conta_origem": f"{lanc.codigo_conta} - {lanc.descricao_conta}"}
        if lanc.tipo == "credito":
            aggr[chave]["creditos"] += lanc.valor
        else:
            aggr[chave]["debitos"] += lanc.valor

    resumos: list[ResumoMensal] = []
    for (comp, cat), vals in aggr.items():
        resumos.append(ResumoMensal(
            competencia=comp,
            conta_origem=vals["conta_origem"],
            conta_normalizada=cat,
            saldo_inicial=None,
            creditos=round(vals["creditos"], 2),
            debitos=round(vals["debitos"], 2),
            saldo_final=None,
            arquivo_origem=lancamentos[0].arquivo_origem,
            observacao="agregado_dos_lancamentos",
        ))

    logger.info("Resumo mensal por conta gerado: %d categorias", len(resumos))
    return resumos
