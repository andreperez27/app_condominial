import logging

import pandas as pd

logger = logging.getLogger(__name__)

CAT_COBERTURA = {
    "AGUA": {"cobranca": ["AGUA"], "despesa": ["DESPESA_AGUA"]},
    "GAS": {"cobranca": ["GAS"], "despesa": ["DESPESA_GAS"]},
}

COD_ACORDOS = {"ACORDOS"}
COD_ATRASO = {
    "ATRASO_COTA_CONDOMINIAL", "ATRASO_BENFEITORIAS",
    "ATRASO_FUNDO_RESERVA", "ATRASO_SALAO_FESTAS",
}


def analisar_cobertura(lancamentos_df: pd.DataFrame) -> dict:
    if lancamentos_df.empty:
        return {"tabela_mensal": pd.DataFrame(), "detalhe_unidade": pd.DataFrame()}

    linhas_mensal: list[dict] = []
    linhas_unidade: list[dict] = []

    todas_comps = sorted(lancamentos_df["competencia"].unique())

    for comp in todas_comps:
        df_comp = lancamentos_df[lancamentos_df["competencia"] == comp]

        for cat_nome, map_info in CAT_COBERTURA.items():
            cats_cob = map_info["cobranca"]
            cats_desp = map_info["despesa"]

            cobrancas = df_comp[
                df_comp["categoria"].isin(cats_cob) & (df_comp["tipo"] == "credito")
            ]
            despesas = df_comp[
                df_comp["categoria"].isin(cats_desp) & (df_comp["tipo"] == "debito")
            ]

            valor_cobrado = cobrancas["valor"].sum()
            valor_arrecadado = valor_cobrado
            valor_despesa = despesas["valor"].sum()
            valor_em_aberto = valor_cobrado - valor_arrecadado

            acordos_rel = df_comp[
                df_comp["categoria"].isin(COD_ACORDOS) & (df_comp["tipo"] == "credito")
            ]
            valor_recuperado_acordo = acordos_rel["valor"].sum()

            cobertura_arrecadada_pct = (
                (valor_arrecadado / valor_despesa * 100)
                if valor_despesa > 0 else None
            )
            cobertura_cobrada_pct = (
                (valor_cobrado / valor_despesa * 100)
                if valor_despesa > 0 else None
            )

            if valor_despesa == 0:
                status = "sem_despesa"
            elif valor_arrecadado >= valor_despesa:
                status = "cobriu"
            elif valor_cobrado >= valor_despesa:
                status = "cobriria_sem_inadimplencia"
            else:
                status = "nao_cobriu_mesmo_sem_inadimplencia"

            linhas_mensal.append({
                "competencia": comp,
                "categoria": cat_nome,
                "valor_cobrado": round(valor_cobrado, 2),
                "valor_arrecadado": round(valor_arrecadado, 2),
                "valor_despesa": round(valor_despesa, 2),
                "valor_em_aberto": round(valor_em_aberto, 2),
                "valor_recuperado_acordo": round(valor_recuperado_acordo, 2),
                "cobertura_arrecadada_pct": round(cobertura_arrecadada_pct, 2) if cobertura_arrecadada_pct is not None else None,
                "cobertura_cobrada_pct": round(cobertura_cobrada_pct, 2) if cobertura_cobrada_pct is not None else None,
                "status_analise": status,
            })

            # ── Detalhe por unidade ──────────────────────────────────
            unids_cob = cobrancas[cobrancas["unidade"].notna()]
            unids = sorted(unids_cob["unidade"].unique())

            for unid in unids:
                valor_cobrado_unid = unids_cob[unids_cob["unidade"] == unid]["valor"].sum()
                em_atraso = df_comp[
                    df_comp["categoria"].isin(COD_ATRASO)
                    & (df_comp["tipo"] == "credito")
                    & (df_comp["unidade"] == unid)
                ]["valor"].sum()
                em_acordo = df_comp[
                    df_comp["categoria"].isin(COD_ACORDOS)
                    & (df_comp["tipo"] == "credito")
                    & (df_comp["unidade"] == unid)
                ]["valor"].sum()

                qtde_lanc = len(unids_cob[unids_cob["unidade"] == unid])
                tem_atraso = em_atraso > 0
                tem_acordo = em_acordo > 0

                inadimplente = "sim" if (tem_atraso or tem_acordo) else "nao"
                # sem dados de pagamento individualizados, consideramos arrecadado = cobrado
                valor_arrecadado_unid = valor_cobrado_unid
                valor_em_aberto_unid = valor_cobrado_unid - valor_arrecadado_unid

                linhas_unidade.append({
                    "competencia": comp,
                    "categoria": cat_nome,
                    "unidade": unid,
                    "valor_cobrado": round(valor_cobrado_unid, 2),
                    "valor_pago": round(valor_arrecadado_unid, 2),
                    "valor_em_aberto": round(valor_em_aberto_unid, 2),
                    "qtde_lancamentos": qtde_lanc,
                    "valor_em_atraso": round(em_atraso, 2),
                    "valor_em_acordo": round(em_acordo, 2),
                    "indicio_inadimplencia": inadimplente,
                })

    df_mensal = pd.DataFrame(linhas_mensal)
    df_unidade = pd.DataFrame(linhas_unidade)

    logger.info(
        "Cobertura analisada: %d linhas mensais, %d por unidade",
        len(df_mensal), len(df_unidade),
    )
    return {"tabela_mensal": df_mensal, "detalhe_unidade": df_unidade}
