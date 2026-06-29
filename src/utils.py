import logging
import re
from pathlib import Path

import pandas as pd
import yaml

logger = logging.getLogger(__name__)


def load_mapa_contas(path: str | Path) -> dict:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Mapa de contas nao encontrado: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def parse_unidade(texto: str) -> str | None:
    # Apenas padrões explícitos de unidade — nunca aceitar número solto
    m = re.search(r"RC\.UN\.(\d{4,6})", texto)
    if m:
        return str(int(m.group(1)))
    m = re.search(r"UN\.(\d{4,6})", texto)
    if m:
        return str(int(m.group(1)))
    # UN seguido de número com separador
    m = re.search(r"UN[\.:\s]+(\d{3,4})(?:\s|$)", texto)
    if m:
        return str(int(m.group(1)))
    return None


def parse_referencia_rb(texto: str) -> str | None:
    m = re.search(r"RB\.(\d+)", texto)
    if m:
        return m.group(1)
    m = re.search(r"RC\.(\d+)", texto)
    if m:
        return m.group(1)
    return None


def parse_vencimento(texto: str) -> str | None:
    m = re.search(r"VC\.(\d{2}/\d{2}/\d{2,4})", texto)
    if m:
        return m.group(1)
    m = re.search(r"V\.(\d{2}/\d{2}/\d{2,4})", texto)
    if m:
        return m.group(1)
    m = re.search(r"VC\s+(\d{2}/\d{2}/\d{2,4})", texto)
    if m:
        return m.group(1)
    m = re.search(r"(\d{2}/\d{2}/\d{2,4})", texto)
    if m:
        return m.group(1)
    return None


def parse_valor_bruto(valor_str) -> float | None:
    if valor_str is None:
        return None
    try:
        if isinstance(valor_str, (int, float)):
            return float(valor_str)
        s = str(valor_str).strip().strip('"').strip("'")
        if not s:
            return None
        s = s.replace(" ", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "." in s and s.count(".") > 1:
            s = s.replace(".", "")
        return float(s)
    except (ValueError, TypeError):
        return None


def parse_competencia_from_periodo(texto: str) -> str | None:
    m = re.search(r"(\d{2})/(\d{4})", texto)
    if m:
        mes, ano = m.group(1), m.group(2)
        return f"{ano}-{mes}"
    return None


def parse_competencia_from_arquivo(nome: str) -> str | None:
    m = re.search(r"M(\d{2})[-_]?(\d{2,4})", nome)
    if m:
        mes, ano = m.group(1), m.group(2)
        if len(ano) == 2:
            ano = "20" + ano if int(ano) <= 50 else "19" + ano
        return f"{ano}-{mes}"
    return None


def eh_linha_recebimento_unidade(texto: str) -> bool:
    return bool(re.search(r"RC\.UN\.\d{4,6}", texto)) or bool(
        re.search(r"RC\.MLT\.UN\.\d{4,6}", texto)
    )


# ─────────────────────────────────────────────────────────────────
# Column schema enforcement
# ─────────────────────────────────────────────────────────────────

COLUMN_AUDIT_LOG: list[dict] = []

CANONICAL_ALIASES: dict[str, list[str]] = {
    "descricao_conta": ["descricao", "desc_conta", "conta_descricao", "descricao_da_conta"],
    "codigo_conta": ["cod_conta", "conta_codigo", "cd_conta"],
    "historico_bruto": ["historico", "hist_bruto", "historico_original"],
    "arquivo_origem": ["arquivo", "origem_arquivo", "nome_arquivo", "arquivo_orig"],
    "unidade": ["unidade_identificada", "num_unidade", "numero_unidade", "unidade_detectada"],
    "data_lancamento": ["data", "dt_lancamento", "data_contabil", "data_lancto"],
    "competencia": ["competencia_referencia", "ref_competencia", "periodo"],
    "data_pagamento": ["dt_pagamento", "data_pgto"],
    "categoria": ["categoria_original", "cat"],
}

COLUMN_DEFAULTS: dict[str, object] = {
    "descricao_conta": "",
    "codigo_conta": "",
    "historico_bruto": "",
    "arquivo_origem": "",
    "unidade": "",
    "data_lancamento": "",
    "data_pagamento": "",
    "competencia": "",
    "categoria": "",
    "valor": 0.0,
    "valor_esperado": 0.0,
    "valor_pago": 0.0,
    "observacoes": "",
    "status_validacao": "",
    "incluir_no_calculo": "sim",
    "origem_ajuste_manual": "",
    "chave_deduplicacao": "",
    "flag_duplicado": "",
    "motivo_exclusao": "",
    "entrou_no_boleto": "",
}


def limpar_auditoria_colunas():
    COLUMN_AUDIT_LOG.clear()


def relatorio_colunas_ausentes() -> pd.DataFrame:
    if not COLUMN_AUDIT_LOG:
        return pd.DataFrame()
    return pd.DataFrame(COLUMN_AUDIT_LOG)


def normalizar_nomes_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Mapeia aliases conhecidos para nomes canônicos."""
    df = df.copy()
    rename = {}
    for canonico, aliases in CANONICAL_ALIASES.items():
        for alias in aliases:
            if alias in df.columns and canonico not in df.columns:
                rename[alias] = canonico
                break
    if rename:
        df = df.rename(columns=rename)
        logger.debug("Colunas renomeadas: %s", rename)
    return df


def ensure_columns(
    df: pd.DataFrame,
    colunas_obrigatorias: list[str],
    defaults: dict | None = None,
    stage: str = "",
) -> pd.DataFrame:
    """
    Garante que todas as colunas obrigatórias existam no DataFrame.

    - Normaliza nomes de colunas (aliases → canônico)
    - Adiciona colunas faltantes com valores padrão
    - Mantém auditoria em COLUMN_AUDIT_LOG
    - Retorna DataFrame seguro para uso com as colunas obrigatórias

    Uso típico:
        df = ensure_columns(df, ["competencia", "unidade", "descricao_conta"], stage="tab_visao")
        display_df = df[["competencia", "unidade", "descricao_conta"]]   # seguro
    """
    if df is None:
        logger.warning("ensure_columns recebeu None em stage='%s'", stage)
        cols = {c: pd.Series(dtype="object") for c in colunas_obrigatorias}
        return pd.DataFrame(cols)

    if df.empty:
        cols = {c: pd.Series(dtype="object") for c in colunas_obrigatorias}
        return pd.DataFrame(cols)

    df = df.copy()

    # 1. Normaliza nomes
    df = normalizar_nomes_colunas(df)

    # 2. Identifica faltantes
    faltantes = [c for c in colunas_obrigatorias if c not in df.columns]

    # 3. Auditoria
    for col in faltantes:
        COLUMN_AUDIT_LOG.append({
            "dataframe_origem": stage,
            "coluna_ausente": col,
            "etapa_do_pipeline": stage,
            "acao_aplicada": f"adicionada_com_default",
        })
        logger.info("Coluna '%s' ausente em '%s' — adicionada com default", col, stage)

    # 4. Adiciona faltantes
    resolved_defaults = defaults or {}
    for col in faltantes:
        if col in resolved_defaults:
            df[col] = resolved_defaults[col]
        elif col in COLUMN_DEFAULTS:
            df[col] = COLUMN_DEFAULTS[col]
        else:
            df[col] = ""

    return df


def ensure_schema(
    df: pd.DataFrame,
    schema: dict[str, object],
    stage: str = "",
) -> pd.DataFrame:
    """
    Garante schema mínimo via dicionário {coluna: valor_padrao}.
    Útil quando os defaults diferem dos padrões globais.

    Exemplo:
        df = ensure_schema(df, {
            "competencia": "", "unidade": "", "descricao_conta": "",
            "valor": 0.0, "status_validacao": "pendente",
        }, stage="tab_auditoria")
    """
    return ensure_columns(df, list(schema.keys()), defaults=schema, stage=stage)
