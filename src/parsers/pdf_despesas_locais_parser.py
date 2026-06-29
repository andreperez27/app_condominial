"""
Parser for PDF files containing "Conta Despesas Locais" - local cash expenses.
Extracts structured data from the summary page and receipt/invoice pages.
"""

import logging
import re
from dataclasses import dataclass

import pdfplumber

logger = logging.getLogger(__name__)


@dataclass
class LancamentoDespesaLocal:
    competencia: str
    data: str
    historico_pdf: str
    valor_debito: float | None
    valor_credito: float | None
    tipo_movimento: str  # "debito" | "credito" | "saldo"
    numero_pagina: int
    texto_bruto: str
    fornecedor_sugerido: str | None = None
    nf_numero: str | None = None
    cnpj: str | None = None
    documento_sugerido: str | None = None
    origem_comprovante: str | None = None


def _parse_br(val: str) -> float | None:
    if not val or not val.strip():
        return None
    v = val.strip().replace(".", "").replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


def parse_pdf_despesas_locais(
    caminho: str,
) -> tuple[list[LancamentoDespesaLocal], str | None, str | None, str | None]:
    lancamentos: list[LancamentoDespesaLocal] = []
    competencia: str | None = None
    periodo_ini: str | None = None
    periodo_fim: str | None = None

    logger.info("Abrindo PDF: %s", caminho)

    with pdfplumber.open(caminho) as pdf:
        if not pdf.pages:
            logger.warning("PDF vazio: %s", caminho)
            return lancamentos, competencia, periodo_ini, periodo_fim

        texto_pag1 = pdf.pages[0].extract_text() or ""

        # --- detect period ---
        m_periodo = re.search(
            r"Per[íi]odo\s*de:\s*(\d{2}/\d{2}/\d{4})\s*at[ée]\s*(\d{2}/\d{2}/\d{4})",
            texto_pag1,
            re.IGNORECASE,
        )
        if m_periodo:
            periodo_ini = m_periodo.group(1)
            periodo_fim = m_periodo.group(2)
            dia, mes, ano = periodo_fim.split("/")
            competencia = f"{ano}-{mes}"
            logger.info(
                "Período detectado: %s até %s -> competência %s",
                periodo_ini,
                periodo_fim,
                competencia,
            )
        else:
            logger.warning("Período não encontrado na primeira página")

        # --- parse summary table from page 1 ---
        linhas = texto_pag1.split("\n")
        for linha in linhas:
            linha = linha.strip()
            if not linha:
                continue

            if re.search(
                r"Saldo|Pág\.|Conta Despesas|Per[íi]odo|^Data|Hist[óo]rico|D[ée]bito|Cr[ée]dito",
                linha,
                re.IGNORECASE,
            ):
                continue

            m_data = re.match(r"^(\d{2}/\d{2}/\d{4})\s+(.*)", linha)
            if not m_data:
                continue

            data = m_data.group(1)
            resto = m_data.group(2).strip()

            num_re = r"(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}"
            nums = re.findall(num_re, resto)

            debito_val: float | None = None
            credito_val: float | None = None
            historico = resto

            if len(nums) >= 2:
                debito_val = _parse_br(nums[-2])
                credito_val = _parse_br(nums[-1])
                # Remove apenas os dois últimos números (débito/crédito),
                # mantendo eventuais valores R$ no meio do histórico
                penultimo_idx = resto.rfind(nums[-2])
                historico = resto[:penultimo_idx].rstrip()
            elif len(nums) == 1:
                v = _parse_br(nums[0])
                if v is not None:
                    if "REMESSA" in resto.upper():
                        credito_val = v
                    else:
                        debito_val = v
                # Com apenas 1 número, pode fazer parte da descrição
                # ex: "SINAL DE R$ 1.200,00 REFERENTE A..."
                idx = resto.rfind(nums[0])
                if idx >= 0:
                    historico = resto[:idx].rstrip()

            if debito_val is not None and debito_val > 0:
                tipo = "debito"
            elif credito_val is not None and credito_val > 0:
                tipo = "credito"
            else:
                tipo = "debito"

            lanc = LancamentoDespesaLocal(
                competencia=competencia or "",
                data=data,
                historico_pdf=historico,
                valor_debito=debito_val,
                valor_credito=credito_val,
                tipo_movimento=tipo,
                numero_pagina=1,
                texto_bruto=linha,
            )
            lancamentos.append(lanc)

        logger.info("Lançamentos extraídos da página 1: %d", len(lancamentos))

    return lancamentos, competencia, periodo_ini, periodo_fim


def extrair_fornecedores_de_paginas(
    lancamentos: list[LancamentoDespesaLocal], pdf_path: str
) -> list[LancamentoDespesaLocal]:
    logger.info("Extraindo fornecedores de comprovantes: %s", pdf_path)

    comprovantes: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, pagina in enumerate(pdf.pages):
            if i < 2:
                continue
            texto = pagina.extract_text() or ""
            if not texto.strip():
                continue

            fornecedor = _extrair_fornecedor(texto)
            cnpj = _extrair_cnpj(texto)
            nf_numero = _extrair_nf_numero(texto)
            valor = _extrair_valor_comprovante(texto)

            if fornecedor or cnpj or nf_numero or valor is not None:
                comprovantes.append(
                    {
                        "fornecedor": fornecedor,
                        "cnpj": cnpj,
                        "nf_numero": nf_numero,
                        "valor": valor,
                        "pagina": i + 1,
                    }
                )

    if comprovantes:
        logger.info("Comprovantes encontrados: %d", len(comprovantes))
        for lanc in lancamentos:
            if lanc.valor_debito is None or lanc.valor_debito <= 0:
                continue
            for comp in comprovantes:
                if comp["valor"] is None:
                    continue
                if abs(lanc.valor_debito - comp["valor"]) < 0.50:
                    lanc.fornecedor_sugerido = comp["fornecedor"]
                    lanc.cnpj = comp["cnpj"]
                    lanc.nf_numero = comp["nf_numero"]
                    lanc.origem_comprovante = f"página {comp['pagina']}"
                    lanc.documento_sugerido = (
                        f"NF-e {comp['nf_numero']}" if comp["nf_numero"] else "Recibo"
                    )
                    break
    else:
        logger.info("Nenhum comprovante encontrado nas páginas seguintes")

    return lancamentos


def _extrair_fornecedor(texto: str) -> str | None:
    m = re.search(r"RECEBEMOS\s+DE\s+(.+)", texto, re.IGNORECASE)
    if m:
        return m.group(1).strip().split("\n")[0].strip().rstrip(".")
    m = re.search(
        r"(?:FORNECEDOR|DESTINAT[ÁA]RIO|RAZ[ÃA]O SOCIAL)\s*:?\s*(.+)",
        texto,
        re.IGNORECASE,
    )
    if m:
        return m.group(1).strip().split("\n")[0].strip().rstrip(".")
    return None


def _extrair_cnpj(texto: str) -> str | None:
    m = re.search(r"CNPJ\s*:?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if m:
        return m.group(1)
    m = re.search(r"CNPJ[:\s]*(\d{14})", texto)
    if m:
        raw = m.group(1)
        return f"{raw[:2]}.{raw[2:5]}.{raw[5:8]}/{raw[8:12]}-{raw[12:]}"
    return None


def _extrair_nf_numero(texto: str) -> str | None:
    m = re.search(r"N[Ff]\s*[-–—eE]?\s*[eE]?\s*[:\s]*n?[°º]?\s*(\d+)", texto)
    if m:
        return m.group(1)
    m = re.search(r"N[°º]\s*(\d+)", texto)
    if m:
        return m.group(1)
    m = re.search(r"(?:NF[-\s]?[Ee]?|N[Ff][Ee]?)\s*[:\s]*(\d+)", texto)
    if m:
        return m.group(1)
    return None


def _extrair_valor_comprovante(texto: str) -> float | None:
    m = re.search(
        r"(?:VALOR\s*(?:TOTAL)?|Total)\s*(?:[:\s]*R?\$?\s*)?([\d.]+,\d{2})",
        texto,
        re.IGNORECASE,
    )
    if m:
        v = _parse_br(m.group(1))
        if v is not None and v > 0:
            return v
    vals = re.findall(r"([\d.]+,\d{2})", texto)
    for v in reversed(vals):
        parsed = _parse_br(v)
        if parsed is not None and parsed > 0:
            return parsed
    return None
