from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal
from typing import Optional


@dataclass
class LancamentoExtrato:
    competencia: str
    data_lancamento: str
    unidade: Optional[str]
    codigo_conta: str
    descricao_conta: str
    categoria: Optional[str]
    historico_bruto: str
    referencia_rb: Optional[str]
    vencimento: Optional[str]
    valor: float
    tipo: str  # 'credito' ou 'debito'
    arquivo_origem: str
    linha_origem: int
    observacao_parse: str = ""


@dataclass
class RecebivelUnidade:
    competencia: str
    unidade: str
    categoria: str
    valor: float
    data_lancamento: str
    codigo_conta: str
    descricao_conta: str
    referencia_rb: Optional[str]
    vencimento: Optional[str]
    historico_bruto: str
    arquivo_origem: str


@dataclass
class ResumoMensal:
    competencia: str
    conta_origem: str
    conta_normalizada: Optional[str]
    saldo_inicial: Optional[float]
    creditos: Optional[float]
    debitos: Optional[float]
    saldo_final: Optional[float]
    arquivo_origem: str
    observacao: str = ""


@dataclass
class LinhaConciliacao:
    competencia: str
    conta_normalizada: str
    saldo_extrato: Optional[float]
    saldo_planilha: Optional[float]
    diferenca: Optional[float]
    status_conciliacao: str
    observacao: str = ""


@dataclass
class ContaMapeada:
    codigo: str
    descricao: str
    categoria: str


@dataclass
class LinhaNaoInterpretada:
    competencia: str
    conteudo_bruto: str
    codigo_conta: Optional[str]
    descricao_conta: Optional[str]
    motivo: str
    arquivo_origem: str
    linha_origem: int
    valor: Optional[float] = None
