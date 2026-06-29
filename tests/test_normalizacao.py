import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.utils import load_mapa_contas
from src.parser_extrato import ParserExtrato


CONFIG_DIR = os.path.join(os.path.dirname(__file__), "..", "config", "mapa_contas.yaml")


class TestNormalizacao:
    def setup_method(self):
        self.parser = ParserExtrato(CONFIG_DIR)

    def test_get_categoria_1002(self):
        assert self.parser.get_categoria("1002", "COTAS CONDOMINIAIS") == "COTA_CONDOMINIAL"

    def test_get_categoria_1010(self):
        assert self.parser.get_categoria("1010", "CONSUMO DE GAS") == "GAS"

    def test_get_categoria_1052(self):
        assert self.parser.get_categoria("1052", "CONSUMO DE AGUA") == "AGUA"

    def test_get_categoria_3002(self):
        assert self.parser.get_categoria("3002", "COTAS DO MES FUNDO DE RESERVA") == "FUNDO_RESERVA"

    def test_get_categoria_2002(self):
        assert self.parser.get_categoria("2002", "COTAS DO MES BENFEITORIAS") == "BENFEITORIAS"

    def test_get_categoria_19002(self):
        assert self.parser.get_categoria("19002", "COTAS DO MES SALAO DE FESTAS") == "SALAO_FESTAS"

    def test_get_categoria_1098(self):
        assert self.parser.get_categoria("1098", "MULTAS") == "MULTAS"

    def test_get_categoria_1006(self):
        assert self.parser.get_categoria("1006", "RECEITA PARCELAS ACORDO") == "ACORDOS"

    def test_get_categoria_1023(self):
        assert self.parser.get_categoria("1023", "CORRECAO MONETARIA") == "CORRECAO_MONETARIA"

    def test_get_categoria_1026(self):
        assert self.parser.get_categoria("1026", "REND.APLIC.FINANC") == "RENDIMENTO_APLICACAO"

    def test_get_categoria_desconhecida(self):
        assert self.parser.get_categoria("9999", "CONTA DESCONHECIDA") is None

    def test_load_mapa_contas(self):
        mapa = load_mapa_contas(CONFIG_DIR)
        assert "mapeamento" in mapa
        assert "mapeamento_texto" in mapa
        assert "mapeamento_planilha" in mapa
        assert len(mapa["mapeamento"]) > 0


if __name__ == "__main__":
    t = TestNormalizacao()
    t.setup_method()
    for name in dir(t):
        if name.startswith("test_"):
            getattr(t, name)()
            print(f"  PASS: {name}")
