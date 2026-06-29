import logging
import re
from pathlib import Path

import pandas as pd

from src.models import ResumoMensal, LinhaNaoInterpretada
from src.utils import load_mapa_contas, parse_valor_bruto

logger = logging.getLogger(__name__)


class ParserPlanilha:
    def __init__(self, mapa_contas_path: str | Path):
        self.mapa = load_mapa_contas(mapa_contas_path)
        self._build_map()

    def _build_map(self):
        self.planilha_map: dict[str, str] = {}
        for k, v in self.mapa.get("mapeamento_planilha", {}).items():
            self.planilha_map[k.upper().strip()] = v

    def normalizar_conta_planilha(self, nome: str) -> str | None:
        chave = nome.upper().strip()
        if chave in self.planilha_map:
            return self.planilha_map[chave]
        for k, v in self.planilha_map.items():
            if k in chave or chave in k:
                return v
        return None

    def parse_resumo(
        self, caminho: str | Path
    ) -> tuple[list[ResumoMensal], list[LinhaNaoInterpretada]]:
        caminho = Path(caminho)
        logger.info("Lendo planilha-resumo: %s", caminho)

        xl = pd.ExcelFile(caminho, engine="openpyxl")
        if "resumo do saldo" not in xl.sheet_names:
            logger.warning("Aba 'resumo do saldo' nao encontrada na planilha")
            return [], []

        df = pd.read_excel(caminho, sheet_name="resumo do saldo", header=None)
        resumos: list[ResumoMensal] = []
        nao_interpretadas: list[LinhaNaoInterpretada] = []

        competencia_atual: str | None = None
        extra_cols = max(1, df.shape[1])

        for idx, row in df.iterrows():
            linha_num = idx + 1
            col0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""

            if not col0:
                continue

            m_comp = re.match(
                r"Saldo\s+referente\s+ao\s+m[eê]s\s+(\d{2})/(\d{4})",
                col0,
                re.IGNORECASE,
            )
            if m_comp:
                competencia_atual = f"{m_comp.group(2)}-{m_comp.group(1)}"
                logger.debug("Competencia planilha: %s", competencia_atual)
                continue

            if "Histórico" in col0 or "Hist" in col0 or "DESCRIÇÃO" in col0 or col0 == "DESCRICAO":
                header_check = col0.upper()
                if "SALDO ANT" in header_check:
                    pass
                continue

            if not competencia_atual:
                continue

            linha_vals = []
            for c in range(extra_cols):
                v = row.iloc[c] if c < len(row) else None
                linha_vals.append(v)

            col4 = row.iloc[4] if len(row) > 4 else None
            if col0.upper() == "SALDO FINAL":
                resumos.append(
                    ResumoMensal(
                        competencia=competencia_atual,
                        conta_origem=col0,
                        conta_normalizada="SALDO_FINAL",
                        saldo_inicial=parse_valor_bruto(row.iloc[1] if len(row) > 1 else None),
                        creditos=parse_valor_bruto(row.iloc[2] if len(row) > 2 else None),
                        debitos=parse_valor_bruto(row.iloc[3] if len(row) > 3 else None),
                        saldo_final=parse_valor_bruto(col4),
                        arquivo_origem=caminho.name,
                        observacao="",
                    )
                )
                continue

            cat_norm = self.normalizar_conta_planilha(col0)
            if cat_norm:
                resumos.append(
                    ResumoMensal(
                        competencia=competencia_atual,
                        conta_origem=col0,
                        conta_normalizada=cat_norm,
                        saldo_inicial=parse_valor_bruto(row.iloc[1] if len(row) > 1 else None),
                        creditos=parse_valor_bruto(row.iloc[2] if len(row) > 2 else None),
                        debitos=parse_valor_bruto(row.iloc[3] if len(row) > 3 else None),
                        saldo_final=parse_valor_bruto(col4),
                        arquivo_origem=caminho.name,
                        observacao="",
                    )
                )
            else:
                nao_interpretadas.append(
                    LinhaNaoInterpretada(
                        competencia=competencia_atual,
                        conteudo_bruto=col0,
                        codigo_conta=None,
                        descricao_conta=col0,
                        motivo="conta_nao_mapeada_planilha",
                        arquivo_origem=caminho.name,
                        linha_origem=linha_num,
                    )
                )

        logger.info(
            "Planilha parseada: %d resumos, %d nao mapeadas",
            len(resumos),
            len(nao_interpretadas),
        )
        return resumos, nao_interpretadas

    def parse_extrato_sheet(
        self, caminho: str | Path, sheet_name: str
    ) -> tuple[list, list, list]:
        caminho = Path(caminho)
        logger.info("Lendo aba '%s' de: %s", sheet_name, caminho)
        from src.parser_extrato import ParserExtrato

        df = pd.read_excel(caminho, sheet_name=sheet_name, header=None, engine="openpyxl")
        temp_path = caminho.parent / f"__temp_{sheet_name}.xls"
        try:
            df.to_excel(temp_path, index=False, header=False, engine="xlrd")
            parser = ParserExtrato(str(self.mapa))
            return parser.parse(temp_path)
        finally:
            if temp_path.exists():
                temp_path.unlink()
