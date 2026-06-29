import hashlib
import logging
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

from src.services.competencia_detector import detectar_competencia, detectar_tipo_arquivo

logger = logging.getLogger(__name__)

PREFIXOS = {
    "EXTRATO_BANCARIO": "extrato_condominio",
    "PDF_DESPESAS_LOCAIS": "despesas_locais",
    "PLANILHA_RESUMO": "planilha_resumo",
    "XLS_DESCONHECIDO": "desconhecido",
    "PDF_DESCONHECIDO": "desconhecido",
    "DESCONHECIDO": "desconhecido",
}


@dataclass
class ArquivoIndexado:
    nome_original: str
    nome_renomeado: str
    caminho: str
    tipo: str
    competencia: str
    periodo_inicio: str = ""
    periodo_fim: str = ""
    hash_md5: str = ""
    tamanho_kb: float = 0.0
    data_inclusao: str = ""
    status: str = "ok"  # ok, ignorado_duplicata, erro
    mensagem: str = ""


@dataclass
class BaseIndexada:
    arquivos: list[ArquivoIndexado] = field(default_factory=list)
    competencias: set[str] = field(default_factory=set)
    tipos: set[str] = field(default_factory=set)
    erros: list[ArquivoIndexado] = field(default_factory=list)
    duplicatas: list[ArquivoIndexado] = field(default_factory=list)
    pasta: str = ""
    total_arquivos: int = 0
    ultima_atualizacao: str = ""


def _hash_arquivo(caminho: Path) -> str:
    return hashlib.md5(caminho.read_bytes()).hexdigest()


def escanear_pasta(pasta: str | Path) -> BaseIndexada:
    """
    Escaneia a pasta oficial, indexa todos os arquivos renomeados,
    detecta tipo, competência, hash, duplicatas e erros.
    """
    dst = Path(pasta)
    if not dst.exists():
        logger.warning("Pasta não encontrada: %s", pasta)
        return BaseIndexada(pasta=str(pasta))

    padroes = (
        sorted(dst.glob("extrato_condominio_*.xls"))
        + sorted(dst.glob("extrato_condominio_*.xlsx"))
        + sorted(dst.glob("despesas_locais_*.pdf"))
        + sorted(dst.glob("planilha_resumo_*.xlsx"))
        + sorted(dst.glob("planilha_resumo_*.xls"))
    )

    indexados: list[ArquivoIndexado] = []
    erros: list[ArquivoIndexado] = []
    duplicatas: list[ArquivoIndexado] = []
    seen_hashes: dict[str, Path] = {}
    competencias: set[str] = set()
    tipos: set[str] = set()

    for arquivo in padroes:
        if arquivo.name.startswith("~$"):
            continue

        nome_original = arquivo.name
        tamanho_kb = round(arquivo.stat().st_size / 1024, 1)
        data_ino = datetime.fromtimestamp(arquivo.stat().st_mtime).strftime("%Y-%m-%d %H:%M")

        # Extrair competência e tipo do nome renomeado
        comp = ""
        tipo = "DESCONHECIDO"
        periodo_ini = ""
        periodo_fim = ""

        if "extrato_condominio_" in arquivo.name:
            tipo = "EXTRATO_BANCARIO"
            try:
                comp_part = arquivo.stem.replace("extrato_condominio_", "").split("__")[0]
                if comp_part:
                    comp = comp_part
            except IndexError:
                pass
        elif "despesas_locais_" in arquivo.name:
            tipo = "PDF_DESPESAS_LOCAIS"
            try:
                comp_part = arquivo.stem.replace("despesas_locais_", "").split("__")[0]
                if comp_part:
                    comp = comp_part
            except IndexError:
                pass
        elif "planilha_resumo_" in arquivo.name:
            tipo = "PLANILHA_RESUMO"
            try:
                comp_part = arquivo.stem.replace("planilha_resumo_", "").split("__")[0]
                if comp_part:
                    comp = comp_part
            except IndexError:
                pass
        else:
            # Fallback: detector genérico
            tipo = detectar_tipo_arquivo(str(arquivo))
            comp, periodo_ini, periodo_fim = detectar_competencia(str(arquivo))

        if not comp:
            erros.append(ArquivoIndexado(
                nome_original=nome_original,
                nome_renomeado=arquivo.name,
                caminho=str(arquivo),
                tipo=tipo,
                competencia="",
                status="erro",
                mensagem="Não foi possível detectar competência",
                tamanho_kb=tamanho_kb,
                data_inclusao=data_ino,
            ))
            continue

        try:
            hash_val = _hash_arquivo(arquivo)
        except Exception as e:
            erros.append(ArquivoIndexado(
                nome_original=nome_original,
                nome_renomeado=arquivo.name,
                caminho=str(arquivo),
                tipo=tipo,
                competencia=comp,
                status="erro",
                mensagem=f"Erro ao ler arquivo: {e}",
                tamanho_kb=tamanho_kb,
                data_inclusao=data_ino,
            ))
            continue

        # Detectar duplicatas por hash
        chave_hash = f"{tipo}_{hash_val}"
        if chave_hash in seen_hashes:
            existente = seen_hashes[chave_hash]
            duplicatas.append(ArquivoIndexado(
                nome_original=nome_original,
                nome_renomeado=arquivo.name,
                caminho=str(arquivo),
                tipo=tipo,
                competencia=comp,
                hash_md5=hash_val,
                status="ignorado_duplicata",
                mensagem=f"Hash igual a {existente.name}",
                tamanho_kb=tamanho_kb,
                data_inclusao=data_ino,
            ))
            continue

        seen_hashes[chave_hash] = arquivo
        competencias.add(comp)
        tipos.add(tipo)

        indexados.append(ArquivoIndexado(
            nome_original=nome_original,
            nome_renomeado=arquivo.name,
            caminho=str(arquivo),
            tipo=tipo,
            competencia=comp,
            periodo_inicio=periodo_ini,
            periodo_fim=periodo_fim,
            hash_md5=hash_val,
            tamanho_kb=tamanho_kb,
            data_inclusao=data_ino,
            status="ok",
            mensagem="",
        ))

    base = BaseIndexada(
        arquivos=indexados,
        competencias=sorted(competencias),
        tipos=sorted(tipos),
        erros=erros,
        duplicatas=duplicatas,
        pasta=str(dst),
        total_arquivos=len(indexados),
        ultima_atualizacao=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    logger.info(
        "Pasta indexada: %d arquivos, %d competências, %d erros, %d duplicatas",
        base.total_arquivos, len(base.competencias), len(base.erros), len(base.duplicatas),
    )
    return base


def gerar_relatorio_index(base: BaseIndexada) -> str:
    """Gera relatório texto do estado da base."""
    linhas = [
        f"Pasta: {base.pasta}",
        f"Total de arquivos válidos: {base.total_arquivos}",
        f"Competências: {', '.join(sorted(base.competencias))}",
        f"Tipos: {', '.join(base.tipos)}",
        f"Erros: {len(base.erros)}",
        f"Duplicatas ignoradas: {len(base.duplicatas)}",
        f"Última atualização: {base.ultima_atualizacao}",
        "",
        "Arquivos:",
    ]
    for a in base.arquivos:
        linhas.append(f"  [{a.tipo:20s}] {a.competencia}  {a.nome_renomeado}  ({a.tamanho_kb} KB)")
    if base.erros:
        linhas.append("")
        linhas.append("Erros:")
        for a in base.erros:
            linhas.append(f"  {a.nome_renomeado}: {a.mensagem}")
    return "\n".join(linhas)
