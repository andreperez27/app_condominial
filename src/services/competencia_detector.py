import logging
import re
from datetime import datetime
from pathlib import Path

import xlrd

logger = logging.getLogger(__name__)

DATA_RE = re.compile(r"(\d{2})/(\d{2})/(\d{4})")
PERIODO_PDF_RE = re.compile(
    r"Per[íi]odo\s*de:\s*(\d{2}/\d{2}/\d{4})\s*at[ée]\s*(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE,
)


def _detectar_xls(caminho: str):
    """Detecta competência de arquivos .xls/.xlsx."""
    wb = xlrd.open_workbook(caminho)
    ws = wb.sheet_by_index(0)

    periodo_ini = None
    periodo_fim = None
    competencia = None

    for row in range(ws.nrows):
        row_vals = [str(ws.cell_value(row, c)).strip() for c in range(ws.ncols)]

        for val in row_vals:
            m = re.search(r"COMPET[ÊE]NCIA\s*:?\s*(\d{2})/(\d{4})", val, re.IGNORECASE)
            if m:
                competencia = f"{m.group(2)}-{m.group(1)}"
                break

        texto_linha = " ".join(row_vals)
        if "PER" in texto_linha.upper() or "PERIODO" in texto_linha.upper():
            datas = [m.group(0) for val in row_vals for m in DATA_RE.finditer(val)]
            if len(datas) >= 2:
                periodo_ini = datas[0]
                periodo_fim = datas[-1]
            elif len(datas) == 1:
                periodo_fim = datas[0]

    return competencia, periodo_ini, periodo_fim


def _detectar_pdf(caminho: str):
    """Detecta competência de arquivos PDF de despesas locais."""
    try:
        import pdfplumber
        with pdfplumber.open(caminho) as pdf:
            if not pdf.pages:
                return None, None, None
            texto = pdf.pages[0].extract_text() or ""
            m = PERIODO_PDF_RE.search(texto)
            if m:
                periodo_ini = m.group(1)
                periodo_fim = m.group(2)
                dia, mes, ano = periodo_fim.split("/")
                competencia = f"{ano}-{mes}"
                return competencia, periodo_ini, periodo_fim
    except ImportError:
        logger.warning("pdfplumber não instalado, não é possível ler PDFs")
    except Exception as e:
        logger.warning("Erro ao processar PDF %s: %s", caminho, e)
    return None, None, None


def detectar_competencia(caminho: str) -> tuple[str | None, str | None, str | None]:
    caminho_p = Path(caminho)
    ext = caminho_p.suffix.lower()

    if ext == ".pdf":
        return _detectar_pdf(caminho)
    elif ext == ".xls":
        comp, periodo_ini, periodo_fim = _detectar_xls(caminho)

        if not comp and periodo_fim:
            try:
                dt = datetime.strptime(periodo_fim, "%d/%m/%Y")
                comp = dt.strftime("%Y-%m")
                logger.info("Competência inferida da data final do período: %s", comp)
            except ValueError:
                pass

        if comp:
            logger.info("Competência detectada: %s em %s", comp, caminho)
        else:
            logger.warning("Não foi possível detectar competência em %s", caminho)

        return comp, periodo_ini, periodo_fim

    logger.warning("Tipo de arquivo não suportado: %s", ext)
    return None, None, None


def detectar_tipo_arquivo(caminho: str) -> str:
    """Detecta o tipo do arquivo pelo conteúdo."""
    caminho_p = Path(caminho)
    ext = caminho_p.suffix.lower()

    if ext == ".pdf":
        try:
            import pdfplumber
            with pdfplumber.open(caminho) as pdf:
                if pdf.pages:
                    texto = pdf.pages[0].extract_text() or ""
                    if "Conta Despesas Locais" in texto:
                        return "PDF_DESPESAS_LOCAIS"
        except Exception:
            pass
        return "PDF_DESCONHECIDO"
    elif ext == ".xls":
        try:
            wb = xlrd.open_workbook(caminho)
            ws = wb.sheet_by_index(0)
            for row in range(min(5, ws.nrows)):
                for col in range(ws.ncols):
                    val = str(ws.cell_value(row, col))
                    if "Conta Corrente" in val or "QUINTA DE NIAGARA" in val:
                        return "EXTRATO_BANCARIO"
        except Exception:
            pass
        return "XLS_DESCONHECIDO"
    elif ext == ".xlsx":
        return "PLANILHA_RESUMO"

    return "DESCONHECIDO"
