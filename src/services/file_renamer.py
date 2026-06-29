import csv
import hashlib
import logging
import shutil
from pathlib import Path

from src.services.competencia_detector import detectar_competencia, detectar_tipo_arquivo

logger = logging.getLogger(__name__)

LOG_COLUNAS = [
    "nome_original", "nome_novo", "tipo_arquivo", "competencia",
    "periodo_inicio", "periodo_fim", "status", "mensagem",
]

PREFIXOS = {
    "EXTRATO_BANCARIO": "extrato_condominio",
    "PDF_DESPESAS_LOCAIS": "despesas_locais",
    "PLANILHA_RESUMO": "planilha_resumo",
    "XLS_DESCONHECIDO": "desconhecido",
    "PDF_DESCONHECIDO": "desconhecido",
    "DESCONHECIDO": "desconhecido",
}


def _hash_arquivo(caminho: Path) -> str:
    return hashlib.md5(caminho.read_bytes()).hexdigest()


def _proximo_nome_disponivel(dest_dir: Path, base: str, ext: str) -> Path:
    cand = dest_dir / f"{base}{ext}"
    if not cand.exists():
        return cand
    v = 2
    while True:
        cand = dest_dir / f"{base}__v{v}{ext}"
        if not cand.exists():
            return cand
        v += 1


def _buscar_por_hash(dest_dir: Path, hash_orig: str, base: str) -> Path | None:
    for f in dest_dir.iterdir():
        if f.is_file() and f.stem.startswith(base):
            try:
                if _hash_arquivo(f) == hash_orig:
                    return f
            except Exception:
                continue
    return None


def varrer_e_renomear(
    input_dir: str | Path,
    renamed_dir: str | Path,
    log_path: str | Path | None = None,
) -> list[dict]:
    src = Path(input_dir)
    dst = Path(renamed_dir)
    dst.mkdir(parents=True, exist_ok=True)

    if log_path is None:
        log_path = Path("logs") / "log_renomeacao.csv"
    log_path = Path(log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    padroes = (
        sorted(src.glob("*.[xX][lL][sS]"))
        + sorted(src.glob("*.[xX][lL][sS][xX]"))
        + sorted(src.glob("*.[pP][dD][fF]"))
    )

    registros: list[dict] = []

    for arquivo in padroes:
        ext = arquivo.suffix.lower()
        if arquivo.name.startswith("~$"):
            continue

        tipo = detectar_tipo_arquivo(str(arquivo))
        comp, periodo_ini, periodo_fim = detectar_competencia(str(arquivo))

        if not comp:
            registros.append({
                "nome_original": arquivo.name,
                "nome_novo": "",
                "tipo_arquivo": tipo,
                "competencia": "",
                "periodo_inicio": periodo_ini or "",
                "periodo_fim": periodo_fim or "",
                "status": "ERRO",
                "mensagem": f"Não foi possível detectar competência (tipo: {tipo})",
            })
            logger.error("Arquivo %s: não foi possível detectar competência", arquivo.name)
            continue

        prefixo = PREFIXOS.get(tipo, "desconhecido")
        base = f"{prefixo}_{comp}"
        hash_orig = _hash_arquivo(arquivo)

        existente = _buscar_por_hash(dst, hash_orig, base)
        if existente:
            registros.append({
                "nome_original": arquivo.name,
                "nome_novo": existente.name,
                "tipo_arquivo": tipo,
                "competencia": comp,
                "periodo_inicio": periodo_ini or "",
                "periodo_fim": periodo_fim or "",
                "status": "IGNORADO",
                "mensagem": f"Arquivo já existe com mesmo hash ({existente.name})",
            })
            logger.info("Arquivo %s já renomeado (hash igual a %s)", arquivo.name, existente.name)
            continue

        destino = _proximo_nome_disponivel(dst, base, ext)
        shutil.move(str(arquivo), str(destino))

        registros.append({
            "nome_original": arquivo.name,
            "nome_novo": destino.name,
            "tipo_arquivo": tipo,
            "competencia": comp,
            "periodo_inicio": periodo_ini or "",
            "periodo_fim": periodo_fim or "",
            "status": "RENOMEADO",
            "mensagem": f"Movido para {destino.name}",
        })
        logger.info("Arquivo %s -> %s (tipo=%s, comp=%s)", arquivo.name, destino.name, tipo, comp)

    with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=LOG_COLUNAS)
        writer.writeheader()
        writer.writerows(registros)

    logger.info("Renomeação concluída: %d arquivos processados, log em %s", len(registros), log_path)
    return registros
