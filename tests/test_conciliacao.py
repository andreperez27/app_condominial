import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pandas as pd

from src.conciliacao import conciliar


class TestConciliacao:
    def test_conciliacao_vazia(self):
        df = conciliar(pd.DataFrame(), pd.DataFrame())
        assert df.empty

    def test_conciliacao_simples(self):
        extrato = pd.DataFrame([
            {"competencia": "2025-05", "conta_normalizada": "COTA_CONDOMINIAL",
             "saldo_inicial": 100.0, "creditos": 500.0, "debitos": 200.0,
             "saldo_final": 400.0, "arquivo_origem": "extrato.xls"}
        ])
        planilha = pd.DataFrame([
            {"competencia": "2025-05", "conta_normalizada": "COTA_CONDOMINIAL",
             "saldo_inicial": 100.0, "creditos": 500.0, "debitos": 200.0,
             "saldo_final": 400.0, "arquivo_origem": "planilha.xlsx"}
        ])
        df = conciliar(extrato, planilha)
        assert not df.empty
        row = df.iloc[0]
        assert row["status_conciliacao"] == "conciliado"
        assert abs(row["diferenca"]) < 0.01

    def test_conciliacao_divergente(self):
        extrato = pd.DataFrame([
            {"competencia": "2025-05", "conta_normalizada": "COTA_CONDOMINIAL",
             "saldo_final": 400.0}
        ])
        planilha = pd.DataFrame([
            {"competencia": "2025-05", "conta_normalizada": "COTA_CONDOMINIAL",
             "saldo_final": 350.0}
        ])
        df = conciliar(extrato, planilha)
        assert not df.empty
        row = df.iloc[0]
        assert row["status_conciliacao"] == "divergente"
        assert row["diferenca"] == 50.0

    def test_conciliacao_apenas_extrato(self):
        extrato = pd.DataFrame([
            {"competencia": "2025-05", "conta_normalizada": "COTA_CONDOMINIAL",
             "saldo_final": 400.0}
        ])
        df = conciliar(extrato, pd.DataFrame())
        assert not df.empty
        row = df.iloc[0]
        assert row["status_conciliacao"] == "apenas_no_extrato"


if __name__ == "__main__":
    t = TestConciliacao()
    for name in dir(t):
        if name.startswith("test_"):
            getattr(t, name)()
            print(f"  PASS: {name}")
