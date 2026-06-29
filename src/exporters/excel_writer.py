import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, numbers
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14)
MONTH_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
TOTAL_FONT = Font(bold=True, size=11)
CURRENCY_FMT = '#,##0.00'
DATE_FMT = 'DD/MM/YYYY'


def _ajustar_larguras(ws, min_width=10, max_width=50):
    for col_cells in ws.columns:
        max_len = min_width
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[col_letter].width = max_len


def _estilo_cabecalho(ws, row_num, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _estilo_linha(ws, row_num, ncols, fill=None, font=None):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.border = BORDER
        if fill:
            cell.fill = fill
        if font:
            cell.font = font


def escrever_planilha_consolidada(
    output_path: str | Path,
    df_lancamentos: pd.DataFrame,
    df_resumo: pd.DataFrame,
    auditoria: list[dict],
    input_dir: str = "",
):
    """Gera arquivo Excel com abas: Extrato Consolidado, Resumo Mensal, Auditoria."""
    wb = Workbook()

    _escrever_extrato(wb, df_lancamentos)
    _escrever_resumo(wb, df_resumo)
    if auditoria:
        _escrever_auditoria(wb, auditoria)

    # Remover sheet padrão se não usada
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > (3 if auditoria else 2):
        del wb["Sheet"]

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
    logger.info("Planilha salva: %s", output_path)


def _escrever_extrato(wb: Workbook, df: pd.DataFrame):
    """Aba 1: Extrato Consolidado."""
    ws = wb.active
    ws.title = "Extrato Consolidado"

    # Cabeçalho
    colunas = [
        "competencia", "data_lancamento", "data_documento", "codigo_conta",
        "descricao_conta", "historico", "documento", "tipo_movimento",
        "debito", "credito", "valor",
        "saldo_inicial_mes", "saldo_final_mes",
        "arquivo_origem", "nome_arquivo", "observacao_parse",
        "categoria", "unidade",
    ]
    colunas_existentes = [c for c in colunas if c in df.columns]

    ws.append(colunas_existentes)
    _estilo_cabecalho(ws, 1, len(colunas_existentes))

    # Dados
    for _, row in df.iterrows():
        linha = []
        for col in colunas_existentes:
            val = row.get(col)
            if isinstance(val, (datetime, pd.Timestamp)):
                try:
                    val = val.strftime("%d/%m/%Y")
                except (ValueError, AttributeError):
                    val = ""
            elif isinstance(val, float):
                val = round(val, 2)
            elif val is None or (isinstance(val, float) and pd.isna(val)):
                val = ""
            linha.append(val)
        ws.append(linha)

    # Formatar células de moeda
    moeda_cols = []
    for i, col in enumerate(colunas_existentes, 1):
        if col in ("debito", "credito", "valor", "saldo_inicial_mes", "saldo_final_mes"):
            moeda_cols.append(i)
            for row_num in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_num, column=i)
                cell.number_format = CURRENCY_FMT

    # Formatar datas
    for i, col in enumerate(colunas_existentes, 1):
        if col in ("data_lancamento", "data_documento"):
            for row_num in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_num, column=i)
                if cell.value and isinstance(cell.value, str) and "/" in cell.value:
                    cell.number_format = DATE_FMT

    # Congelar primeira linha
    ws.freeze_panes = "A2"
    _ajustar_larguras(ws)


def _escrever_resumo(wb: Workbook, df: pd.DataFrame):
    """Aba 2: Resumo Mensal no formato da planilha modelo."""
    ws = wb.create_sheet("Resumo Mensal")

    if df.empty:
        ws.append(["Nenhum dado de resumo disponível."])
        return

    comps = sorted(df["competencia"].unique())
    linha_atual = 1

    for comp in comps:
        df_comp = df[df["competencia"] == comp]

        # Título do mês
        try:
            ano, mes = comp.split("-")
            ws.cell(row=linha_atual, column=1, value=f"Saldo referente ao mês {mes}/{ano}")
        except (ValueError, IndexError):
            ws.cell(row=linha_atual, column=1, value=f"Saldo referente ao mês {comp}")
        ws.cell(row=linha_atual, column=1).font = TITLE_FONT
        linha_atual += 1

        # Cabeçalho
        headers = ["DESCRIÇÃO", "SALDO ANT.", "CRÉDITOS", "DÉBITOS", "SALDO"]
        for c, h in enumerate(headers, 1):
            ws.cell(row=linha_atual, column=c, value=h)
        _estilo_cabecalho(ws, linha_atual, len(headers))
        linha_atual += 1

        # Linhas de categorias
        for _, row in df_comp.iterrows():
            desc = row.get("descricao", "")
            ws.cell(row=linha_atual, column=1, value=desc)
            ws.cell(row=linha_atual, column=1).font = Font(size=11)

            for c, col in enumerate(["saldo_anterior", "creditos", "debitos", "saldo"], 2):
                val = row.get(col, 0)
                if pd.isna(val):
                    val = 0
                cell = ws.cell(row=linha_atual, column=c, value=round(float(val), 2))
                cell.number_format = CURRENCY_FMT

            # Destacar SALDO FINAL
            if desc == "SALDO FINAL":
                _estilo_linha(ws, linha_atual, len(headers), fill=TOTAL_FILL, font=TOTAL_FONT)
            else:
                _estilo_linha(ws, linha_atual, len(headers))

            linha_atual += 1

        # Linha em branco entre meses
        linha_atual += 1

    _ajustar_larguras(ws, min_width=15)


def _escrever_auditoria(wb: Workbook, auditoria: list[dict]):
    """Aba 3: Auditoria."""
    ws = wb.create_sheet("Auditoria")

    headers = ["arquivo_origem", "tipo_problema", "linha_origem", "observacao", "severidade"]
    ws.append(headers)
    _estilo_cabecalho(ws, 1, len(headers))

    for item in auditoria:
        ws.append([
            item.get("arquivo_origem", ""),
            item.get("tipo_problema", ""),
            item.get("linha_origem", ""),
            item.get("observacao", ""),
            item.get("severidade", ""),
        ])

    # Colorir por severidade
    cores = {"erro": "FCE4EC", "aviso": "FFF3CD"}
    for row_num in range(2, ws.max_row + 1):
        sev = ws.cell(row=row_num, column=5).value
        fill = PatternFill(start_color=cores.get(sev, "FFFFFF"), end_color=cores.get(sev, "FFFFFF"), fill_type="solid")
        for col in range(1, 6):
            ws.cell(row=row_num, column=col).fill = fill
            ws.cell(row=row_num, column=col).border = BORDER

    ws.freeze_panes = "A2"
    _ajustar_larguras(ws)


def exportar_csvs(output_dir: str | Path, df_lancamentos: pd.DataFrame, df_resumo: pd.DataFrame):
    """Exporta CSVs auxiliares."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    caminho_lanc = output_dir / "extrato_consolidado.csv"
    df_lancamentos.to_csv(caminho_lanc, index=False, encoding="utf-8-sig")
    logger.info("CSV salvo: %s", caminho_lanc)

    caminho_resumo = output_dir / "resumo_mensal.csv"
    df_resumo.to_csv(caminho_resumo, index=False, encoding="utf-8-sig")
    logger.info("CSV salvo: %s", caminho_resumo)
