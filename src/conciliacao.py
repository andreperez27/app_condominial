import logging

import pandas as pd

logger = logging.getLogger(__name__)


def conciliar(
    resumo_extrato: pd.DataFrame,
    resumo_planilha: pd.DataFrame,
) -> pd.DataFrame:
    logger.info("Iniciando conciliacao entre extrato e planilha")

    if resumo_extrato is None or resumo_extrato.empty:
        logger.warning("Resumo do extrato vazio")
        return pd.DataFrame()

    if resumo_planilha is None:
        resumo_planilha = pd.DataFrame()

    extrato_norm = resumo_extrato.copy()
    if "conta_normalizada" not in extrato_norm.columns:
        extrato_norm["conta_normalizada"] = extrato_norm.get("conta_origem", "")
    if "competencia" not in extrato_norm.columns:
        extrato_norm["competencia"] = ""

    # Guardar linhas RESUMO antes de filtrar (para consulta de saldo_inicial)
    resumo_rows = extrato_norm[
        extrato_norm["conta_normalizada"] == "RESUMO"
    ].copy()

    extrato_norm = extrato_norm[
        extrato_norm["conta_normalizada"] != "RESUMO"
    ].copy()
    extrato_norm = extrato_norm[
        extrato_norm["conta_normalizada"].notna()
    ].copy()
    extrato_norm["conta_normalizada"] = extrato_norm["conta_normalizada"].astype(str)

    planilha_norm = resumo_planilha.copy()
    if "conta_normalizada" not in planilha_norm.columns:
        planilha_norm["conta_normalizada"] = planilha_norm.get("conta_origem", "")
    if "competencia" not in planilha_norm.columns:
        planilha_norm["competencia"] = ""
    planilha_norm = planilha_norm[
        planilha_norm["conta_normalizada"].notna()
    ].copy()
    planilha_norm["conta_normalizada"] = planilha_norm["conta_normalizada"].astype(str)

    linhas: list[dict] = []

    todas_competencias = sorted(
        set(extrato_norm["competencia"].unique())
        | set(planilha_norm["competencia"].unique())
    )
    todas_contas = sorted(
        set(extrato_norm["conta_normalizada"].unique())
        | set(planilha_norm["conta_normalizada"].unique())
    )

    def _get_saldo_inicial(comp: str) -> float:
        """Busca saldo_inicial da linha RESUMO para a competência."""
        resumo_row = resumo_rows[
            resumo_rows["competencia"] == comp
        ]
        if not resumo_row.empty and "saldo_inicial" in resumo_row.columns:
            v = resumo_row["saldo_inicial"].iloc[0]
            return 0.0 if pd.isna(v) else float(v)
        return 0.0

    for comp in todas_competencias:
        saldo_ini_comp = _get_saldo_inicial(comp)
        for conta in todas_contas:
            if conta == "SALDO_FINAL":
                continue
            e = extrato_norm[
                (extrato_norm["competencia"] == comp)
                & (extrato_norm["conta_normalizada"] == conta)
            ]
            p = planilha_norm[
                (planilha_norm["competencia"] == comp)
                & (planilha_norm["conta_normalizada"] == conta)
            ]

            def _get_val(df_row, col):
                if df_row.empty or col not in df_row.columns:
                    return None
                v = df_row[col].iloc[0]
                return None if pd.isna(v) else v

            saldo_e = _get_val(e, "saldo_final")
            cred_e = _get_val(e, "creditos")
            deb_e = _get_val(e, "debitos")

            saldo_p = _get_val(p, "saldo_final")

            if saldo_e is None and cred_e is not None:
                # Usar saldo_inicial do RESUMO para cálculo correto:
                # saldo_final = saldo_inicial + creditos - debitos
                saldo_e = round(saldo_ini_comp + cred_e - (deb_e or 0.0), 2)

            if saldo_e is None and saldo_p is None:
                continue

            diferenca = None
            status = "sem_dados_suficientes"

            if saldo_e is not None and saldo_p is not None:
                diferenca = round(saldo_e - saldo_p, 2)
                if abs(diferenca) < 0.01:
                    status = "conciliado"
                else:
                    status = "divergente"
            elif saldo_e is not None:
                status = "apenas_no_extrato"
            else:
                status = "apenas_na_planilha"

            obs = ""
            if not e.empty and "observacao" in e.columns:
                obs_e = e["observacao"].iloc[0]
                if pd.notna(obs_e) and obs_e:
                    obs = f"Extrato: {obs_e}"
            if not p.empty and "observacao" in p.columns:
                obs_p = p["observacao"].iloc[0]
                if pd.notna(obs_p) and obs_p:
                    if obs:
                        obs += " | "
                    obs += f"Planilha: {obs_p}"

            linhas.append(
                {
                    "competencia": comp,
                    "conta_normalizada": conta,
                    "saldo_extrato": saldo_e,
                    "saldo_planilha": saldo_p,
                    "diferenca": diferenca,
                    "status_conciliacao": status,
                    "observacao": obs,
                }
            )

    df = pd.DataFrame(linhas)
    logger.info("Conciliacao gerada: %d linhas", len(df))
    return df
