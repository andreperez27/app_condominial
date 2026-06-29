import logging
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

from src.models import LancamentoExtrato, LinhaNaoInterpretada, ResumoMensal
from src.utils import (
    load_mapa_contas,
    parse_competencia_from_periodo,
    parse_referencia_rb,
    parse_unidade,
    parse_valor_bruto,
    parse_vencimento,
)

logger = logging.getLogger(__name__)


class ParserExtrato:
    def __init__(self, mapa_contas_path: str | Path):
        self.mapa = load_mapa_contas(mapa_contas_path)
        self._build_categoria_map()

    def _build_categoria_map(self):
        self.categoria_por_codigo: dict[str, str] = {}
        for entry in self.mapa.get("mapeamento", []):
            cod = str(entry["codigo"])
            self.categoria_por_codigo[cod] = entry["categoria"]

    def get_categoria(self, codigo: str, descricao: str) -> str | None:
        cod = codigo.strip()
        if cod in self.categoria_por_codigo:
            return self.categoria_por_codigo[cod]
        for entry in self.mapa.get("mapeamento_texto", []):
            for palavra in entry["palavras"]:
                if palavra.upper() in descricao.upper():
                    return entry["categoria"]
        return None

    def parse(
        self, caminho: str | Path, competencia_filter: str | None = None
    ) -> tuple[list[LancamentoExtrato], list[ResumoMensal], list[LinhaNaoInterpretada]]:
        caminho = Path(caminho)
        logger.info("Lendo extrato: %s", caminho)

        try:
            df = pd.read_excel(caminho, header=None, engine="xlrd")
        except Exception as e:
            raise ValueError(f"Erro ao ler arquivo .xls: {e}") from e

        lancamentos: list[LancamentoExtrato] = []
        resumos: list[ResumoMensal] = []
        nao_interpretadas: list[LinhaNaoInterpretada] = []
        periodo: str | None = None
        em_resumo = False

        for idx, row in df.iterrows():
            linha_num = idx + 1
            raw = [str(v) if pd.notna(v) else "" for v in row]
            linha_texto = " ".join(raw).strip()

            if not linha_texto:
                continue

            col0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            col1 = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            col2 = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
            col3_raw = row.iloc[3] if len(row) > 3 else None
            col4_raw = row.iloc[4] if len(row) > 4 else None

            if col0 == "Conta Corrente":
                continue

            if "Período" in col0 or "Período" in col0:
                periodo = parse_competencia_from_periodo(linha_texto)
                logger.debug("Periodo detectado: %s", periodo)
                continue

            if col0 == "Conta" and col1 == "Data":
                continue

            if re.search(r"^RESUMO", col0, re.IGNORECASE):
                em_resumo = True
                logger.debug("Bloco RESUMO detectado")
                continue

            if em_resumo:
                resumo = self._parse_linha_resumo(
                    row, periodo or "desconhecido", caminho.name, linha_num
                )
                if resumo:
                    resumos.append(resumo)
                continue

            header_match = re.match(r"^\s*(\d{4,5})\s*-\s*(.+?)$", col0)
            if not header_match:
                continue

            conta_codigo = header_match.group(1).strip()
            conta_descricao = header_match.group(2).strip()

            if not col2:
                continue

            data_str = col1
            debito = parse_valor_bruto(col3_raw) or 0.0
            credito = parse_valor_bruto(col4_raw) or 0.0

            if not col2 or (debito == 0.0 and credito == 0.0):
                continue

            lanc = self._parse_lancamento(
                historico=col2,
                data_str=data_str,
                debito=debito,
                credito=credito,
                conta_codigo=conta_codigo,
                conta_descricao=conta_descricao,
                periodo=periodo or "desconhecido",
                arquivo=caminho.name,
                linha_num=linha_num,
            )

            if lanc:
                lancamentos.append(lanc)
            else:
                nao_interpretadas.append(
                    LinhaNaoInterpretada(
                        competencia=periodo or "desconhecido",
                        conteudo_bruto=col2,
                        codigo_conta=conta_codigo,
                        descricao_conta=conta_descricao,
                        motivo="historico_nao_interpretado",
                        arquivo_origem=caminho.name,
                        linha_origem=linha_num,
                        valor=credito or debito,
                    )
                )

        logger.info(
            "Extrato parseado: %d lancamentos, %d resumos, %d nao interpretadas",
            len(lancamentos),
            len(resumos),
            len(nao_interpretadas),
        )
        return lancamentos, resumos, nao_interpretadas

    def _parse_lancamento(
        self,
        historico: str,
        data_str: str,
        debito: float,
        credito: float,
        conta_codigo: str,
        conta_descricao: str,
        periodo: str,
        arquivo: str,
        linha_num: int,
    ) -> LancamentoExtrato | None:
        historico_limpo = historico.strip().strip('"').strip("'")
        unidade = parse_unidade(historico_limpo)
        referencia = parse_referencia_rb(historico_limpo)
        vencimento = parse_vencimento(historico_limpo)
        categoria = self.get_categoria(conta_codigo, conta_descricao)

        if not categoria:
            logger.warning(
                "Conta nao mapeada: %s - %s (linha %d)",
                conta_codigo,
                conta_descricao,
                linha_num,
            )

        data_lanc = data_str
        if data_lanc:
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(data_lanc, fmt)
                    data_lanc = dt.strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue

        valor = credito if credito > 0 else debito
        tipo = "credito" if credito > 0 else "debito"

        return LancamentoExtrato(
            competencia=periodo,
            data_lancamento=data_lanc,
            unidade=unidade,
            codigo_conta=conta_codigo,
            descricao_conta=conta_descricao,
            categoria=categoria,
            historico_bruto=historico_limpo,
            referencia_rb=referencia,
            vencimento=vencimento,
            valor=valor,
            tipo=tipo,
            arquivo_origem=arquivo,
            linha_origem=linha_num,
            observacao_parse="",
        )

    def _parse_linha_resumo(
        self, row, periodo: str, arquivo: str, linha_num: int
    ) -> ResumoMensal | None:
        col0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        col0_upper = col0.upper()

        val_raw = row.iloc[4] if len(row) > 4 else None
        val = parse_valor_bruto(val_raw)

        if "SALDO INICIAL" in col0_upper:
            return ResumoMensal(
                competencia=periodo,
                conta_origem="RESUMO",
                conta_normalizada="RESUMO",
                saldo_inicial=val,
                creditos=None, debitos=None, saldo_final=None,
                arquivo_origem=arquivo,
                observacao="saldo_inicial_do_resumo",
            )
        if "DEBITO TOTAL" in col0_upper.replace("É", "E"):
            return ResumoMensal(
                competencia=periodo,
                conta_origem="RESUMO",
                conta_normalizada="RESUMO",
                saldo_inicial=None, creditos=None, debitos=val, saldo_final=None,
                arquivo_origem=arquivo,
                observacao="debito_total_do_resumo",
            )
        if "CREDITO TOTAL" in col0_upper.replace("É", "E"):
            return ResumoMensal(
                competencia=periodo,
                conta_origem="RESUMO",
                conta_normalizada="RESUMO",
                saldo_inicial=None, creditos=val, debitos=None, saldo_final=None,
                arquivo_origem=arquivo,
                observacao="credito_total_do_resumo",
            )
        if "SALDO FINAL" in col0_upper:
            return ResumoMensal(
                competencia=periodo,
                conta_origem="RESUMO",
                conta_normalizada="RESUMO",
                saldo_inicial=None, creditos=None, debitos=None, saldo_final=val,
                arquivo_origem=arquivo,
                observacao="saldo_final_do_resumo",
            )
        return None
