import logging
from collections import defaultdict
from pathlib import Path

import pandas as pd
import yaml

from src.parser_extrato import ParserExtrato

logger = logging.getLogger(__name__)

CAT_ATRASO = {"ATRASO_COTA_CONDOMINIAL", "ATRASO_BENFEITORIAS", "ATRASO_FUNDO_RESERVA", "ATRASO_SALAO_FESTAS"}
CAT_ACORDO = {"ACORDOS"}
CAT_COBRANCA = {"COTA_CONDOMINIAL", "AGUA", "GAS", "FUNDO_RESERVA", "BENFEITORIAS", "SALAO_FESTAS", "MANUT_PINTURA"}


def carregar_unidades_esperadas(caminho: str | Path) -> list[int]:
    with open(caminho, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return sorted(data["unidades"])


def ler_todos_extratos(renamed_dir: str | Path, mapa_contas: str | Path) -> list:
    """
    Lê todos os extratos renomeados em ordem cronológica de competência.
    Retorna lista de (competencia, lancamentos, resumos, nao_int).
    """
    dst = Path(renamed_dir)
    padroes = sorted(dst.glob("extrato_condominio_*.xls")) + sorted(dst.glob("extrato_condominio_*.xlsx"))
    padroes.sort()

    parser = ParserExtrato(str(mapa_contas))
    # Agrupar por competência; se houver múltiplos (ex: __v2), manter o último
    arquivos_por_comp: dict[str, Path] = {}
    for arquivo in padroes:
        comp = arquivo.stem.replace("extrato_condominio_", "").split("__")[0]
        arquivos_por_comp[comp] = arquivo  # overwrite: último vence

    resultados = []
    for comp in sorted(arquivos_por_comp):
        arquivo = arquivos_por_comp[comp]
        logger.info("Lendo extrato: %s (competência=%s)", arquivo.name, comp)
        lancamentos, resumos, nao_int = parser.parse(str(arquivo))
        resultados.append((comp, lancamentos, resumos, nao_int))

    resultados.sort(key=lambda x: x[0])
    return resultados


def gerar_inventario_unidades(lancamentos_por_comp: list[tuple]) -> dict:
    """
    Para cada competência, identifica unidades presentes, ausentes, etc.
    Retorna dict: {competencia: {unidade: {...}}}
    """
    comp_dict = {}
    for comp, lancs, _, _ in lancamentos_por_comp:
        unidades_presentes: set[int] = set()
        unidade_categorias: dict[int, set[str]] = defaultdict(set)
        unidade_tem_atraso: dict[int, bool] = defaultdict(bool)
        unidade_tem_acordo: dict[int, bool] = defaultdict(bool)
        unidade_total_pago: dict[int, float] = defaultdict(float)

        for l in lancs:
            if l.unidade:
                uid = int(l.unidade)
                unidades_presentes.add(uid)
                if l.categoria:
                    unidade_categorias[uid].add(l.categoria)
                if l.categoria in CAT_ATRASO:
                    unidade_tem_atraso[uid] = True
                if l.categoria in CAT_ACORDO:
                    unidade_tem_acordo[uid] = True
                if l.categoria in CAT_COBRANCA and l.tipo == "credito" and l.valor is not None:
                    unidade_total_pago[uid] += l.valor

        comp_dict[comp] = {
            "unidades_presentes": unidades_presentes,
            "unidade_categorias": dict(unidade_categorias),
            "unidade_tem_atraso": dict(unidade_tem_atraso),
            "unidade_tem_acordo": dict(unidade_tem_acordo),
            "unidade_total_pago": dict(unidade_total_pago),
        }

    return comp_dict


def classificar_unidades(
    unidades_esperadas: list[int],
    comp_dict: dict,
) -> pd.DataFrame:
    """
    Para cada competência e cada unidade esperada, classifica o status.
    """
    competencias = sorted(comp_dict.keys())
    linhas = []

    for i, comp in enumerate(competencias):
        info = comp_dict[comp]
        info_prev = comp_dict.get(competencias[i - 1]) if i > 0 else None

        for unid in unidades_esperadas:
            uid = unid
            apareceu_mes_atual = uid in info["unidades_presentes"]
            apareceu_mes_anterior = uid in info_prev["unidades_presentes"] if info_prev else False

            tem_atraso = info["unidade_tem_atraso"].get(uid, False)
            tem_acordo = info["unidade_tem_acordo"].get(uid, False)
            total_pago = info["unidade_total_pago"].get(uid, 0.0)

            if apareceu_mes_atual:
                if tem_atraso:
                    status = "EM_ATRASO"
                elif tem_acordo:
                    status = "EM_ACORDO"
                else:
                    status = "PAGO_NO_MES"
            else:
                if info_prev and uid in info_prev["unidades_presentes"]:
                    status = "PAGO_ANTECIPADO"
                else:
                    status = "INADIMPLENTE_PRESUMIDO"

            linhas.append({
                "competencia": comp,
                "unidade": str(uid),
                "unidade_int": uid,
                "apareceu_mes_atual": "sim" if apareceu_mes_atual else "nao",
                "apareceu_mes_anterior": "sim" if apareceu_mes_anterior else "nao",
                "em_atraso": "sim" if tem_atraso else "nao",
                "em_acordo": "sim" if tem_acordo else "nao",
                "total_pago_mes": round(total_pago, 2),
                "status_final": status,
            })

    df = pd.DataFrame(linhas)
    if df.empty:
        df = pd.DataFrame(columns=[
            "competencia", "unidade", "unidade_int",
            "apareceu_mes_atual", "apareceu_mes_anterior",
            "em_atraso", "em_acordo", "total_pago_mes", "status_final",
        ])
    return df


def gerar_cobertura_72(
    unidades_esperadas: list[int],
    comp_dict: dict,
    comp_dict_raw: dict | None = None,
) -> pd.DataFrame:
    """
    Gera tabela de cobertura: para cada competência, estatísticas consolidadas.
    """
    competencias = sorted(comp_dict.keys())
    linhas = []

    for comp in competencias:
        info = comp_dict[comp]
        presentes = info["unidades_presentes"]
        ausentes = [u for u in unidades_esperadas if u not in presentes]
        total_presentes = len(presentes)
        total_ausentes = len(ausentes)

        com_atraso = sum(1 for u in presentes if info["unidade_tem_atraso"].get(u, False))
        com_acordo = sum(1 for u in presentes if info["unidade_tem_acordo"].get(u, False))

        linhas.append({
            "competencia": comp,
            "unidades_esperadas": len(unidades_esperadas),
            "unidades_presentes": total_presentes,
            "unidades_ausentes": total_ausentes,
            "com_atraso": com_atraso,
            "com_acordo": com_acordo,
            "pct_cobertura": round(total_presentes / len(unidades_esperadas) * 100, 1) if unidades_esperadas else 0,
        })

    return pd.DataFrame(linhas)
