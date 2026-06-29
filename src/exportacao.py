import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)


def exportar_csv(df: pd.DataFrame, caminho: str | Path, nome: str) -> Path:
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    arquivo = caminho / nome
    df.to_csv(arquivo, index=False, encoding="utf-8-sig")
    logger.info("CSV exportado: %s (%d linhas)", arquivo, len(df))
    return arquivo


def exportar_xlsx(
    sheets: dict[str, pd.DataFrame],
    caminho: str | Path,
) -> Path:
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            sheet_name_clean = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name_clean, index=False)

    logger.info("XLSX exportado: %s (%d abas)", caminho, len(sheets))
    return caminho


def exportar_tudo(
    lancamentos_norm: pd.DataFrame,
    recebiveis_detalhe: pd.DataFrame,
    resumo_unidade: pd.DataFrame,
    resumo_mensal: pd.DataFrame,
    conciliacao_df: pd.DataFrame,
    contas_nao_mapeadas: pd.DataFrame,
    linhas_nao_interpretadas: pd.DataFrame,
    output_dir: str | Path,
) -> dict[str, Path]:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    arquivos = {}

    if not lancamentos_norm.empty:
        arquivos["lancamentos_normalizados"] = exportar_csv(
            lancamentos_norm, output_dir, "lancamentos_normalizados.csv"
        )

    if not recebiveis_detalhe.empty:
        arquivos["recebiveis_por_unidade"] = exportar_csv(
            recebiveis_detalhe, output_dir, "recebiveis_por_unidade.csv"
        )

    if not resumo_unidade.empty:
        arquivos["resumo_unidade"] = exportar_csv(
            resumo_unidade, output_dir, "resumo_unidade.csv"
        )

    if not resumo_mensal.empty:
        arquivos["resumo_mensal"] = exportar_csv(
            resumo_mensal, output_dir, "resumo_mensal.csv"
        )

    if not conciliacao_df.empty:
        arquivos["conciliacao"] = exportar_csv(
            conciliacao_df, output_dir, "conciliacao.csv"
        )

    if not contas_nao_mapeadas.empty:
        arquivos["contas_nao_mapeadas"] = exportar_csv(
            contas_nao_mapeadas, output_dir, "contas_nao_mapeadas.csv"
        )

    if not linhas_nao_interpretadas.empty:
        arquivos["linhas_nao_interpretadas"] = exportar_csv(
            linhas_nao_interpretadas, output_dir, "linhas_nao_interpretadas.csv"
        )

    sheets = {}
    if not resumo_mensal.empty:
        sheets["resumo_mensal"] = resumo_mensal
    if not recebiveis_detalhe.empty:
        sheets["recebiveis_por_unidade"] = recebiveis_detalhe
    if not resumo_unidade.empty:
        sheets["resumo_unidade"] = resumo_unidade
    if not conciliacao_df.empty:
        sheets["conciliacao"] = conciliacao_df
    if not contas_nao_mapeadas.empty:
        sheets["contas_nao_mapeadas"] = contas_nao_mapeadas
    if not linhas_nao_interpretadas.empty:
        sheets["linhas_nao_interpretadas"] = linhas_nao_interpretadas

    if sheets:
        arquivos["conciliacao_xlsx"] = exportar_xlsx(sheets, output_dir / "conciliacao.xlsx")

    return arquivos
