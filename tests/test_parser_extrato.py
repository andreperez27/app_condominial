import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.utils import parse_unidade, parse_referencia_rb, parse_vencimento, parse_valor_bruto


class TestUtils:
    def test_parse_unidade_rc(self):
        assert parse_unidade("RC.UN.000082 RB.51859790 VC.10/05/26") == "82"

    def test_parse_unidade_mlt(self):
        assert parse_unidade("RC.MLT.UN.000063 RC.51520926 V.10/04/26") == "63"

    def test_parse_unidade_decimal(self):
        assert parse_unidade("MUL.DECOB UN.000061 05/05/26 RB.3160157") == "61"

    def test_parse_unidade_decob(self):
        assert parse_unidade("RC.DECOB 000061 RB.3160157 VC 05/05/26") == "61"

    def test_parse_unidade_atr(self):
        assert parse_unidade("RC.ATR.UN.000153 CB.51520962 V.25/05/26") == "153"

    def test_parse_referencia_rb(self):
        assert parse_referencia_rb("RC.UN.000082 RB.51859790 VC.10/05/26") == "51859790"

    def test_parse_referencia_rc(self):
        assert parse_referencia_rb("RC.MLT.UN.000063 RC.51520926 V.10/04/26") == "51520926"

    def test_parse_vencimento_vc(self):
        assert parse_vencimento("RC.UN.000082 RB.51859790 VC.10/05/26") == "10/05/26"

    def test_parse_vencimento_v(self):
        assert parse_vencimento("RC.MLT.UN.000063 RC.51520926 V.10/04/26") == "10/04/26"

    def test_parse_valor_bruto_float(self):
        assert parse_valor_bruto(840.24) == 840.24

    def test_parse_valor_bruto_str(self):
        assert parse_valor_bruto("0") == 0.0

    def test_parse_valor_bruto_br(self):
        assert parse_valor_bruto("158.383,29") == 158383.29

    def test_parse_valor_bruto_none(self):
        assert parse_valor_bruto(None) is None


if __name__ == "__main__":
    t = TestUtils()
    for name in dir(t):
        if name.startswith("test_"):
            getattr(t, name)()
            print(f"  PASS: {name}")
