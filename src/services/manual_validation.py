import hashlib
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

COLUNAS_VALIDACAO = [
    "id_registro",
    "competencia_original",
    "numero_detectado",
    "unidade_corrigida",
    "historico_bruto",
    "codigo_conta",
    "descricao_conta",
    "categoria_original",
    "categoria_corrigida",
    "competencia_corrigida",
    "valor",
    "status_suspeita",
    "acao_manual",
    "incluir_no_calculo",
    "justificativa",
    "data_validacao",
    "usuario",
]

ACOES_DISPONIVEIS = [
    "manter_sem_alteracao",
    "corrigir_unidade",
    "corrigir_competencia",
    "corrigir_categoria",
    "excluir_do_calculo",
    "marcar_como_duplicado",
    "marcar_como_nao_vinculado",
]


def _localizar_arquivo() -> Path:
    """Localiza data/manual_validations.csv relativo ao projeto."""
    candidates = [
        Path.cwd() / "data" / "manual_validations.csv",
        Path(__file__).resolve().parent.parent.parent / "data" / "manual_validations.csv",
    ]
    for c in candidates:
        if c.exists():
            return c
    # Default to CWD
    return candidates[0]


VALIDACOES_FILE = _localizar_arquivo()
IMPACTO_PATH = Path("output") / "impacto_validacoes.csv"


def _ensure_file():
    if not VALIDACOES_FILE.exists():
        VALIDACOES_FILE.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(columns=COLUNAS_VALIDACAO).to_csv(
            VALIDACOES_FILE, index=False, encoding="utf-8-sig"
        )
        logger.info("Arquivo de validações criado: %s", VALIDACOES_FILE)


def carregar_validacoes() -> pd.DataFrame:
    _ensure_file()
    try:
        df = pd.read_csv(VALIDACOES_FILE, encoding="utf-8-sig", dtype=str)
        return df.fillna("")
    except (pd.errors.EmptyDataError, FileNotFoundError):
        return pd.DataFrame(columns=COLUNAS_VALIDACAO)


def salvar_validacao(registro: dict) -> bool:
    _ensure_file()
    df = carregar_validacoes()
    for col in COLUNAS_VALIDACAO:
        registro.setdefault(col, "")
    registro["data_validacao"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    registro.setdefault("usuario", "admin")
    df_novo = pd.DataFrame([registro])
    df = pd.concat([df, df_novo], ignore_index=True)
    df.to_csv(VALIDACOES_FILE, index=False, encoding="utf-8-sig")
    logger.info("Validação salva: %s - %s", registro.get("acao_manual"), registro.get("id_registro"))
    return True


def remover_validacao(id_registro: str) -> bool:
    df = carregar_validacoes()
    antes = len(df)
    df = df[df["id_registro"] != id_registro]
    if len(df) < antes:
        df.to_csv(VALIDACOES_FILE, index=False, encoding="utf-8-sig")
        return True
    return False


def gerar_id_registro(competencia, unidade, codigo_conta, data_lancamento) -> str:
    raw = "|".join([
        str(competencia or ""),
        str(unidade or ""),
        str(codigo_conta or ""),
        str(data_lancamento or ""),
    ])
    return hashlib.md5(raw.encode()).hexdigest()[:12]


def _chave_dedup(competencia, unidade, codigo_conta, data_lancamento, valor) -> str:
    return "|".join([
        str(competencia or ""),
        str(unidade or ""),
        str(codigo_conta or ""),
        str(data_lancamento or ""),
        f"{float(valor or 0):.2f}",
    ])


def _id_from_chave(chave: str) -> str:
    return hashlib.md5(chave.encode()).hexdigest()[:12]


def aplicar_validacoes(df_lancamentos: pd.DataFrame) -> pd.DataFrame:
    """Aplica validações manuais sobre o DataFrame de lançamentos."""
    validacoes = carregar_validacoes()
    df = df_lancamentos.copy()

    df["status_validacao"] = "pendente"
    df["origem_ajuste_manual"] = ""
    df["incluir_no_calculo"] = "sim"
    df["chave_deduplicacao"] = ""

    for idx, row in df.iterrows():
        chave = _chave_dedup(
            row.get("competencia"), row.get("unidade"),
            row.get("codigo_conta"), row.get("data_lancamento"),
            row.get("valor"),
        )
        df.at[idx, "chave_deduplicacao"] = chave

    if validacoes.empty:
        return df

    for _, val in validacoes.iterrows():
        acao = str(val.get("acao_manual", ""))
        incluir = str(val.get("incluir_no_calculo", "sim"))
        id_reg = str(val.get("id_registro", ""))
        if not id_reg:
            continue

        mask = df["chave_deduplicacao"].apply(lambda x: _id_from_chave(x) == id_reg)
        if not mask.any():
            comp = str(val.get("competencia_original", ""))
            unid = str(val.get("numero_detectado", ""))
            cod = str(val.get("codigo_conta", ""))
            mask = (
                (df.get("competencia", "").astype(str) == comp)
                & (df.get("unidade", "").astype(str) == unid)
                & (df.get("codigo_conta", "").astype(str) == cod)
            )

        for idx in df[mask].index:
            df.at[idx, "origem_ajuste_manual"] = f"validacao_{acao}"
            df.at[idx, "incluir_no_calculo"] = incluir

            if acao == "corrigir_unidade":
                nova = str(val.get("unidade_corrigida", ""))
                if nova:
                    df.at[idx, "unidade"] = nova
                    df.at[idx, "status_validacao"] = "corrigido"
            elif acao == "corrigir_competencia":
                nova = str(val.get("competencia_corrigida", ""))
                if nova:
                    df.at[idx, "competencia"] = nova
                    df.at[idx, "status_validacao"] = "corrigido"
            elif acao == "corrigir_categoria":
                nova = str(val.get("categoria_corrigida", ""))
                if nova:
                    df.at[idx, "categoria"] = nova
                    df.at[idx, "status_validacao"] = "corrigido"
            elif acao in ("excluir_do_calculo", "marcar_como_duplicado", "marcar_como_nao_vinculado"):
                df.at[idx, "incluir_no_calculo"] = "nao"
                df.at[idx, "status_validacao"] = "excluido"
            elif acao == "manter_sem_alteracao":
                df.at[idx, "status_validacao"] = "mantido"

    logger.info("Validações aplicadas: %d registros, %d validações ativas", len(df), len(validacoes))
    return df


def gerar_impacto(df_antes: pd.DataFrame, df_depois: pd.DataFrame) -> pd.DataFrame:
    """Gera log de impacto comparando antes/depois."""
    validacoes = carregar_validacoes()
    if validacoes.empty:
        IMPACTO_PATH.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame().to_csv(IMPACTO_PATH, index=False, encoding="utf-8-sig")
        return pd.DataFrame()

    registros = []
    for _, val in validacoes.iterrows():
        id_reg = str(val.get("id_registro", ""))
        if not id_reg:
            continue
        f_antes = df_antes["chave_deduplicacao"].apply(lambda x: _id_from_chave(x) == id_reg) if "chave_deduplicacao" in df_antes else pd.Series(False, index=df_antes.index)
        f_depois = df_depois["chave_deduplicacao"].apply(lambda x: _id_from_chave(x) == id_reg) if "chave_deduplicacao" in df_depois else pd.Series(False, index=df_depois.index)

        ra = df_antes[f_antes].iloc[0] if f_antes.any() else None
        rd = df_depois[f_depois].iloc[0] if f_depois.any() else None

        registros.append({
            "id_registro": id_reg,
            "acao_manual": val.get("acao_manual", ""),
            "unidade_antes": str(ra.get("unidade", "")) if ra is not None else "",
            "unidade_depois": str(rd.get("unidade", "")) if rd is not None else "",
            "competencia_antes": str(ra.get("competencia", "")) if ra is not None else "",
            "competencia_depois": str(rd.get("competencia", "")) if rd is not None else "",
            "categoria_antes": str(ra.get("categoria", "")) if ra is not None else "",
            "categoria_depois": str(rd.get("categoria", "")) if rd is not None else "",
            "valor_antes": ra.get("valor", 0) if ra is not None else 0,
            "valor_depois": rd.get("valor", 0) if rd is not None else 0,
            "incluir_no_calculo": val.get("incluir_no_calculo", "sim"),
            "justificativa": val.get("justificativa", ""),
            "data_validacao": val.get("data_validacao", ""),
        })

    df_impacto = pd.DataFrame(registros)
    IMPACTO_PATH.parent.mkdir(parents=True, exist_ok=True)
    df_impacto.to_csv(IMPACTO_PATH, index=False, encoding="utf-8-sig")
    return df_impacto
