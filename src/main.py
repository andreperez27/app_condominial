import argparse
import logging
import sys
from pathlib import Path

import pandas as pd

from src.conciliacao import conciliar
from src.exportacao import exportar_tudo
from src.logging_config import setup_logging
from src.normalizacao import (
    consolidar_resumo_mensal,
    gerar_lancamentos_normalizados,
    gerar_resumo_mensal_por_conta,
    separar_recebiveis_por_unidade,
)
from src.parser_extrato import ParserExtrato
from src.parser_planilha import ParserPlanilha
from src.services.caixa_local_matcher import (
    conciliar_caixa_local,
    gerar_resumo_caixa_local,
    processar_pdfs_despesas_locais,
)
from src.services.file_renamer import varrer_e_renomear
from src.services.unidade_classifier import (
    carregar_unidades_esperadas,
    classificar_unidades,
    gerar_cobertura_72,
    gerar_inventario_unidades,
    ler_todos_extratos,
)

logger = logging.getLogger(__name__)

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
MAPA_CONTAS = CONFIG_DIR / "mapa_contas.yaml"
UNIDADES_YAML = CONFIG_DIR / "unidades.yaml"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Sistema de conciliacao financeira de condominio"
    )
    parser.add_argument(
        "--extrato",
        type=str,
        help="Caminho do arquivo .xls do extrato da administradora (modo unico)",
    )
    parser.add_argument(
        "--input-dir",
        type=str,
        help="Diretorio com extratos brutos para processamento em lote",
    )
    parser.add_argument(
        "--planilha",
        type=str,
        help="Caminho do arquivo .xlsx da planilha-resumo de saldos",
    )
    parser.add_argument(
        "--out",
        type=str,
        default="output",
        help="Diretorio de saida para relatorios (default: output)",
    )
    parser.add_argument(
        "--competencia",
        type=str,
        default=None,
        help="Filtrar por competencia (formato YYYY-MM)",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Ativa logging verbose",
    )
    parser.add_argument(
        "--mapa-contas",
        type=str,
        default=str(MAPA_CONTAS),
        help="Caminho do arquivo YAML de mapeamento de contas",
    )
    return parser


def processar_unico(args, mapa_path: str):
    """Pipeline para um único extrato (comportamento original)."""
    extrato_path = Path(args.extrato)
    planilha_path = Path(args.planilha) if args.planilha else None

    output_dir = Path(args.out)
    output_dir.mkdir(parents=True, exist_ok=True)

    parser_extrato = ParserExtrato(mapa_path)
    parser_planilha = ParserPlanilha(mapa_path)

    logger.info("--- ETAPA 1: Parse do extrato ---")
    lancamentos, resumos, nao_interpretadas_extrato = parser_extrato.parse(
        extrato_path, args.competencia
    )
    logger.info("Lancamentos: %d | Resumos: %d | Nao interpretadas: %d",
                 len(lancamentos), len(resumos), len(nao_interpretadas_extrato))

    logger.info("--- ETAPA 2: Parse da planilha-resumo ---")
    resumos_planilha = []
    nao_interpretadas_planilha = []
    if planilha_path and planilha_path.exists():
        resumos_planilha, nao_interpretadas_planilha = parser_planilha.parse_resumo(
            planilha_path
        )
        logger.info("Resumos planilha: %d | Nao mapeadas: %d",
                     len(resumos_planilha), len(nao_interpretadas_planilha))

    logger.info("--- ETAPA 3: Normalizacao ---")
    df_lancamentos_norm = gerar_lancamentos_normalizados(lancamentos)
    df_recebiveis_detalhe, df_resumo_unidade = separar_recebiveis_por_unidade(lancamentos)

    resumo_por_conta = gerar_resumo_mensal_por_conta(lancamentos)
    resumos_combinados = resumos + resumo_por_conta
    df_resumo_mensal_extrato = consolidar_resumo_mensal(resumos_combinados)
    df_resumo_mensal_planilha = consolidar_resumo_mensal(resumos_planilha)

    logger.info("--- ETAPA 4: Conciliacao ---")
    df_conciliacao = conciliar(df_resumo_mensal_extrato, df_resumo_mensal_planilha)

    logger.info("--- ETAPA 5: Exportacao ---")
    df_contas_nao_mapeadas = _gerar_nao_mapeadas(lancamentos)
    df_linhas_nao_interpretadas = _gerar_df_nao_interpretadas(
        nao_interpretadas_extrato + nao_interpretadas_planilha
    )

    arquivos = exportar_tudo(
        lancamentos_norm=df_lancamentos_norm,
        recebiveis_detalhe=df_recebiveis_detalhe,
        resumo_unidade=df_resumo_unidade,
        resumo_mensal=df_resumo_mensal_extrato,
        conciliacao_df=df_conciliacao,
        contas_nao_mapeadas=df_contas_nao_mapeadas,
        linhas_nao_interpretadas=df_linhas_nao_interpretadas,
        output_dir=output_dir,
    )

    return {
        "lancamentos": df_lancamentos_norm,
        "recebiveis_detalhe": df_recebiveis_detalhe,
        "resumo_unidade": df_resumo_unidade,
        "resumo_mensal": df_resumo_mensal_extrato,
        "resumo_planilha": df_resumo_mensal_planilha,
        "conciliacao": df_conciliacao,
        "contas_nao_mapeadas": df_contas_nao_mapeadas,
        "nao_interpretadas": df_linhas_nao_interpretadas,
    }


def processar_lote(args, mapa_path: str):
    """Pipeline em lote: renomeia, le todos extratos, classifica unidades, concilia."""
    input_dir = Path(args.input_dir)
    renamed_dir = Path("data/renamed")
    processed_dir = Path("data/processed")
    output_dir = Path(args.out)
    log_dir = Path("logs")

    for d in [renamed_dir, processed_dir, output_dir, log_dir]:
        d.mkdir(parents=True, exist_ok=True)

    # Etapa 1: Renomear arquivos
    logger.info("--- ETAPA 1: Renomeacao dos extratos ---")
    log_renomeacao = varrer_e_renomear(
        input_dir=input_dir,
        renamed_dir=renamed_dir,
        log_path=log_dir / "log_renomeacao.csv",
    )
    logger.info("Arquivos processados: %d", len(log_renomeacao))

    # Etapa 2: Carregar unidades esperadas
    unidades_esp = carregar_unidades_esperadas(UNIDADES_YAML)
    logger.info("Unidades esperadas carregadas: %d", len(unidades_esp))

    # Etapa 3: Ler todos os extratos renomeados
    logger.info("--- ETAPA 3: Leitura dos extratos ---")
    extratos = ler_todos_extratos(renamed_dir, mapa_path)
    logger.info("Extratos lidos: %d", len(extratos))

    # Etapa 4: Consolidar lancamentos de todos os meses
    logger.info("--- ETAPA 4: Consolidacao ---")
    todos_lancamentos = []
    todos_resumos = []
    todas_nao_int = []
    for comp, lancs, resums, nao_int in extratos:
        todos_lancamentos.extend(lancs)
        todos_resumos.extend(resums)
        todas_nao_int.extend(nao_int)

    logger.info("Total lancamentos: %d | Resumos: %d | Nao int: %d",
                 len(todos_lancamentos), len(todos_resumos), len(todas_nao_int))

    df_lancamentos_norm = gerar_lancamentos_normalizados(todos_lancamentos)
    df_recebiveis_detalhe, df_resumo_unidade = separar_recebiveis_por_unidade(todos_lancamentos)

    resumo_por_conta = gerar_resumo_mensal_por_conta(todos_lancamentos)
    resumos_combinados = todos_resumos + resumo_por_conta
    df_resumo_mensal_extrato = consolidar_resumo_mensal(resumos_combinados)

    # Planilha opcional
    planilha_path = Path(args.planilha) if args.planilha else None
    df_resumo_mensal_planilha = _ler_planilha(planilha_path, mapa_path) if planilha_path else None

    # Conciliacao
    logger.info("--- ETAPA 5: Conciliacao ---")
    df_conciliacao = conciliar(df_resumo_mensal_extrato, df_resumo_mensal_planilha)

    # Classificacao das unidades por competencia
    logger.info("--- ETAPA 6: Classificacao das unidades ---")
    comp_dict = gerar_inventario_unidades(extratos)
    df_classificacao = classificar_unidades(unidades_esp, comp_dict)
    df_cobertura = gerar_cobertura_72(unidades_esp, comp_dict)
    df_inadimplentes = df_classificacao[df_classificacao["status_final"] == "INADIMPLENTE_PRESUMIDO"].copy() if not df_classificacao.empty else pd.DataFrame()

    logger.info("Classificacao: %d linhas | Inadimplentes presumidos: %d",
                 len(df_classificacao), len(df_inadimplentes))

    # Caixa Local / PDF despesas locais
    logger.info("--- ETAPA 7: Analise de caixa local ---")
    pdf_data = processar_pdfs_despesas_locais(renamed_dir)
    df_caixa_local = conciliar_caixa_local(todos_lancamentos, pdf_data)
    df_resumo_caixa = gerar_resumo_caixa_local(df_caixa_local)
    if not df_caixa_local.empty:
        df_caixa_sem_match = df_caixa_local[df_caixa_local["status_match_pdf"] == "NAO_ENCONTRADO"].copy()
    else:
        df_caixa_sem_match = pd.DataFrame()
    logger.info("Caixa local: %d linhas conciliadas, %d sem match",
                 len(df_caixa_local), len(df_caixa_sem_match))

    # Exportacoes
    logger.info("--- ETAPA 8: Exportacao ---")
    df_contas_nao_mapeadas = _gerar_nao_mapeadas(todos_lancamentos)
    df_linhas_nao_interpretadas = _gerar_df_nao_interpretadas(todas_nao_int)
    df_log_renomeacao = _log_to_df(log_renomeacao)

    exportar_lote(
        output_dir=output_dir,
        lancamentos_norm=df_lancamentos_norm,
        recebiveis_detalhe=df_recebiveis_detalhe,
        resumo_unidade=df_resumo_unidade,
        resumo_mensal=df_resumo_mensal_extrato,
        conciliacao_df=df_conciliacao,
        contas_nao_mapeadas=df_contas_nao_mapeadas,
        linhas_nao_interpretadas=df_linhas_nao_interpretadas,
        classificacao=df_classificacao,
        cobertura=df_cobertura,
        inadimplentes=df_inadimplentes,
        log_renomeacao=df_log_renomeacao,
        caixa_local=df_caixa_local,
        resumo_caixa=df_resumo_caixa,
        caixa_sem_match=df_caixa_sem_match,
    )

    return {
        "lancamentos": df_lancamentos_norm,
        "recebiveis_detalhe": df_recebiveis_detalhe,
        "resumo_unidade": df_resumo_unidade,
        "resumo_mensal": df_resumo_mensal_extrato,
        "resumo_planilha": df_resumo_mensal_planilha,
        "conciliacao": df_conciliacao,
        "classificacao_unidades": df_classificacao,
        "cobertura_72": df_cobertura,
        "inadimplentes_presumidos": df_inadimplentes,
        "contas_nao_mapeadas": df_contas_nao_mapeadas,
        "nao_interpretadas": df_linhas_nao_interpretadas,
        "log_renomeacao": df_log_renomeacao,
        "caixa_local": df_caixa_local,
        "resumo_caixa": df_resumo_caixa,
        "caixa_sem_match": df_caixa_sem_match,
    }


def _ler_planilha(planilha_path: Path, mapa_path: str):
    from src.parser_planilha import ParserPlanilha
    parser = ParserPlanilha(mapa_path)
    resumos, _ = parser.parse_resumo(planilha_path)
    return consolidar_resumo_mensal(resumos)


def _log_to_df(registros: list[dict]):
    import pandas as pd
    if not registros:
        return pd.DataFrame()
    return pd.DataFrame(registros)


def exportar_lote(output_dir: Path, **dfs):
    """Exporta todos os DataFrames para CSV na pasta output."""
    import pandas as pd

    mapeamento = {
        "recebiveis_por_unidade.csv": "recebiveis_detalhe",
        "resumo_unidade.csv": "resumo_unidade",
        "resumo_mensal.csv": "resumo_mensal",
        "conciliacao.csv": "conciliacao_df",
        "contas_nao_mapeadas.csv": "contas_nao_mapeadas",
        "linhas_nao_interpretadas.csv": "linhas_nao_interpretadas",
        "cobertura_72_unidades.csv": "cobertura",
        "classificacao_unidades.csv": "classificacao",
        "unidades_inadimplentes_presumidas.csv": "inadimplentes",
        "log_renomeacao.csv": "log_renomeacao",
        "caixa_local_conciliado.csv": "caixa_local",
        "resumo_caixa_local.csv": "resumo_caixa",
        "caixa_local_sem_match.csv": "caixa_sem_match",
    }

    for nome_arquivo, chave in mapeamento.items():
        df = dfs.get(chave)
        if df is not None and isinstance(df, pd.DataFrame) and not df.empty:
            caminho = output_dir / nome_arquivo
            df.to_csv(caminho, index=False, encoding="utf-8-sig")
            logger.info("Exportado: %s", caminho)


def _gerar_nao_mapeadas(lancamentos):
    import pandas as pd
    nao = []
    for l in lancamentos:
        if not l.categoria:
            nao.append({
                "competencia": l.competencia,
                "codigo_conta": l.codigo_conta,
                "descricao_conta": l.descricao_conta,
                "arquivo_origem": l.arquivo_origem,
                "quantidade": 1,
            })
    if not nao:
        return pd.DataFrame()
    df = pd.DataFrame(nao)
    df = df.groupby(["competencia", "codigo_conta", "descricao_conta", "arquivo_origem"]).size().reset_index(name="quantidade")
    return df


def _gerar_df_nao_interpretadas(lista):
    import pandas as pd
    if not lista:
        return pd.DataFrame()
    return pd.DataFrame([vars(x) for x in lista])


def main():
    parser = build_parser()
    args = parser.parse_args()

    setup_logging(debug=args.debug)
    logger.info("=" * 60)
    logger.info("SISTEMA DE CONCILIACAO FINANCEIRA - CONDOMINIO")
    logger.info("=" * 60)

    mapa_path = args.mapa_contas
    if not Path(mapa_path).exists():
        logger.error("Mapa de contas nao encontrado: %s", mapa_path)
        sys.exit(1)

    try:
        if args.input_dir:
            processar_lote(args, mapa_path)
        elif args.extrato:
            processar_unico(args, mapa_path)
        else:
            logger.error("Forneca --extrato (modo unico) ou --input-dir (modo lote)")
            sys.exit(1)

        logger.info("=" * 60)
        logger.info("PROCESSAMENTO CONCLUIDO COM SUCESSO")
        logger.info("=" * 60)

    except Exception as e:
        logger.exception("Erro durante o processamento: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
