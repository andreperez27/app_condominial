# Sistema de Conciliação Financeira de Condomínio

Sistema em Python para ler extratos financeiros de condomínio, extrair saldos e recebíveis, separar valores por unidade, consolidar dados por competência e comparar com planilha-resumo de saldos.

## Estrutura de Pastas

```
.
├── src/
│   ├── __init__.py
│   ├── main.py                 # CLI principal
│   ├── parser_extrato.py       # Parse do extrato .xls da administradora
│   ├── parser_planilha.py      # Parse da planilha-resumo .xlsx
│   ├── normalizacao.py         # Consolidação e normalização de dados
│   ├── conciliacao.py          # Conciliação entre extrato e planilha
│   ├── exportacao.py           # Exportação CSV e XLSX
│   ├── models.py               # Dataclasses dos modelos de dados
│   ├── logging_config.py       # Configuração de logging
│   └── utils.py                # Funções utilitárias (regex, parse)
├── config/
│   └── mapa_contas.yaml        # Mapeamento código-conta -> categoria
├── tests/
│   ├── __init__.py
│   ├── test_parser_extrato.py  # Testes de parsing
│   ├── test_normalizacao.py    # Testes de normalização
│   └── test_conciliacao.py     # Testes de conciliação
├── output/                     # Relatórios gerados (gitignored)
├── requirements.txt
└── README.md
```

## Instalação

```bash
pip install -r requirements.txt
```

## Como Usar

```bash
python -m src.main --extrato caminho/para/extrato.xls --planilha caminho/para/planilha.xlsx --out output/
```

### Parâmetros

| Parâmetro       | Descrição                                      |
|-----------------|------------------------------------------------|
| `--extrato`     | Caminho do arquivo .xls do extrato (obrigatório) |
| `--planilha`    | Caminho do arquivo .xlsx da planilha-resumo     |
| `--out`         | Diretório de saída (default: output/)           |
| `--competencia` | Filtrar por competência (formato YYYY-MM)       |
| `--debug`       | Log detalhado                                   |
| `--mapa-contas` | Caminho alternativo do mapa_contas.yaml         |

### Saídas Geradas

Todos os relatórios são salvos em `output/`:

| Arquivo                          | Descrição                                      |
|----------------------------------|------------------------------------------------|
| lancamentos_normalizados.csv     | Todos os lançamentos do extrato normalizados    |
| recebiveis_por_unidade.csv       | Recebíveis detalhados por unidade e categoria   |
| resumo_unidade.csv               | Tabela pivotada por unidade                     |
| resumo_mensal.csv                | Resumo mensal por conta/categoria               |
| conciliacao.csv                  | Comparação extrato vs planilha                  |
| contas_nao_mapeadas.csv          | Contas sem mapeamento de categoria              |
| linhas_nao_interpretadas.csv     | Linhas que não puderam ser interpretadas        |
| conciliacao.xlsx                 | Todas as abas acima em um único arquivo Excel   |

## Formato Esperado dos Arquivos

### Extrato (.xls)

Arquivo exportado pela administradora com colunas: Conta, Data, Histórico, Débito, Crédito. Contém blocos de conta contábil (ex: `1002 - COTAS CONDOMINIAIS`) e linhas de recebimento por unidade (ex: `RC.UN.000082 RB.51859790 VC.10/05/26`). Ao final possui um bloco RESUMO com saldo inicial, débito total, crédito total e saldo final.

### Planilha-Resumo (.xlsx)

Planilha de controle com aba `resumo do saldo` contendo saldos mensais por categoria (CONTA CONDOMINIO, FUNDO DE RESERVA, SALAO DE FESTAS, etc.) e abas mensais com extratos avulsos.

## Lógica de Classificação

A categoria do recebimento é herdada do bloco contábil atual do extrato:

- Se a linha está dentro do bloco `3002 - COTAS DO MES FUNDO DE RESERVA` → classificado como FUNDO_RESERVA
- Se está em `1010 - CONSUMO DE GAS` → classificado como GAS
- Se está em `1052 - CONSUMO DE AGUA` → classificado como AGUA

## Como Ajustar o Mapa de Contas

Edite `config/mapa_contas.yaml`:

```yaml
mapeamento:
  - codigo: 1002
    descricao: "COTAS CONDOMINIAIS"
    categoria: COTA_CONDOMINIAL
```

Para mapeamento por texto (fallback):

```yaml
mapeamento_texto:
  - palavras: ["FUNDO DE RESERVA"]
    categoria: FUNDO_RESERVA
```

Para mapeamento da planilha-resumo:

```yaml
mapeamento_planilha:
  "CONTA CONDOMINIO": COTA_CONDOMINIAL
```

## Limitações Atuais

- Parser do extrato funciona com o formato .xls (HTML/CSV legado)
- A planilha-resumo precisa ter aba chamada `resumo do saldo`
- Valores em formato brasileiro (vírgula como separador decimal) são tratados pelo `parse_valor_bruto`
- Data de competência é extraída do cabeçalho "Período" do extrato
- Unidades são identificadas por padrões `RC.UN.XXXXXX` ou `UN.XXXXXX`

## Próximos Passos Recomendados

1. Adicionar interface web com Streamlit ou Gradio
2. Suporte a múltiplos extratos simultâneos
3. Dashboard interativo com gráficos de saldos ao longo do tempo
4. Histórico de execuções e comparação entre meses
5. Notificações de divergências encontradas
6.API REST para integração com sistemas externos
7. Automação agendada (cron/jobs)
