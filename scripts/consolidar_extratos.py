#!/usr/bin/env python3
"""
Script para consolidar todos os extratos em uma única planilha Excel.

Uso:
  python scripts/consolidar_extratos.py \\
    --input "data/renamed" \\
    --modelo "Planilha controle do condominio.xlsx" \\
    --output "output/extrato_consolidado_condominio.xlsx"

Sem argumentos, usa os defaults do projeto.
"""
import argparse
import logging
import sys
from pathlib import Path

# Garantir que o diretório raiz está no path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.exporters.excel_writer import escrever_planilha_consolidada, exportar_csvs
from src.logging_config import setup_logging
from src.services.extrato_consolidator import consolidar_extratos


def main():
    setup_logging()
    logger = logging.getLogger(__name__)

    parser = argparse.ArgumentParser(
        description="Consolida todos os extratos em uma única planilha Excel"
    )
    parser.add_argument(
        "--input",
        default=str(ROOT / "data" / "renamed"),
        help="Pasta com os extratos renomeados (default: data/renamed/)",
    )
    parser.add_argument(
        "--modelo",
        default=str(ROOT / "Planilha controle do condominio.xlsx"),
        help="Caminho para a planilha modelo (opcional, para referência)",
    )
    parser.add_argument(
        "--output",
        default=str(ROOT / "output" / "extrato_consolidado_condominio.xlsx"),
        help="Caminho de saída para a planilha consolidada",
    )
    parser.add_argument(
        "--mapa",
        default=str(ROOT / "config" / "mapa_contas.yaml"),
        help="Caminho para o mapa de contas YAML",
    )
    parser.add_argument(
        "--csv", action="store_true", default=True,
        help="Exportar também CSVs auxiliares (default: True)",
    )

    args = parser.parse_args()
    input_dir = Path(args.input)
    mapa_contas = Path(args.mapa)
    output_path = Path(args.output)

    if not input_dir.exists():
        logger.error("Pasta de entrada não encontrada: %s", input_dir)
        sys.exit(1)
    if not mapa_contas.exists():
        logger.error("Mapa de contas não encontrado: %s", mapa_contas)
        sys.exit(1)

    logger.info("Lendo extratos de: %s", input_dir)
    logger.info("Mapa de contas: %s", mapa_contas)
    logger.info("Saída: %s", output_path)

    df_lancamentos, df_resumo, auditoria = consolidar_extratos(input_dir, mapa_contas)

    logger.info(
        "Consolidação concluída: %d lançamentos, %d competências no resumo",
        len(df_lancamentos),
        df_resumo["competencia"].nunique() if not df_resumo.empty else 0,
    )
    logger.info("Auditoria: %d registros", len(auditoria))

    escrever_planilha_consolidada(
        output_path=output_path,
        df_lancamentos=df_lancamentos,
        df_resumo=df_resumo,
        auditoria=auditoria,
        input_dir=str(input_dir),
    )

    if args.csv:
        exportar_csvs(output_path.parent, df_lancamentos, df_resumo)

    print(f"\nPlanilha gerada: {output_path}")
    if auditoria:
        print(f"Auditoria: {len(auditoria)} registros (aba Auditoria no Excel)")


if __name__ == "__main__":
    main()
