import io
import os
import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.analise_cobertura import analisar_cobertura
from src.conciliacao import conciliar
from src.services.espelho_boleto import gerar_espelho_boleto, gerar_espelho_consolidado_por_unidade
from src.utils import ensure_columns, ensure_schema, limpar_auditoria_colunas, relatorio_colunas_ausentes
from src.services.unidade_validator import (
    carregar_unidades_validas,
    validar_lancamentos_unidades,
    gerar_historico_unidades,
)
from src.logging_config import setup_logging
from src.normalizacao import (
    consolidar_resumo_mensal,
    gerar_lancamentos_normalizados,
    gerar_resumo_mensal_por_conta,
    separar_recebiveis_por_unidade,
)
from src.parser_planilha import ParserPlanilha
from src.services.caixa_local_matcher import (
    conciliar_caixa_local,
    gerar_resumo_caixa_local,
)
from src.services.file_renamer import varrer_e_renomear
from src.services.unidade_classifier import (
    carregar_unidades_esperadas,
    classificar_unidades,
    gerar_cobertura_72,
    gerar_inventario_unidades,
)

setup_logging()

ROOT = Path(__file__).resolve().parent
CONFIG_MAPA = ROOT / "config" / "mapa_contas.yaml"
CONFIG_UNIDADES = ROOT / "config" / "unidades.yaml"
OUTPUT_DIR = ROOT / "output"
CONSOLIDADO_PATH = OUTPUT_DIR / "extrato_consolidado_condominio.xlsx"

UNIDADES_VALIDAS = carregar_unidades_validas(CONFIG_UNIDADES)

st.set_page_config(
    page_title="Conciliação Financeira - Condomínio",
    page_icon="\U0001f3e0",
    layout="wide",
)

st.markdown("""
<style>
    .kpi-card { background:#f0f2f6; border-radius:10px; padding:20px; text-align:center; }
    .kpi-value { font-size:1.8rem; font-weight:700; }
    .kpi-label { font-size:0.85rem; color:#555; }
    .status-ok { color:#0a0; font-weight:600; }
    .status-warn { color:#e80; font-weight:600; }
    .status-err { color:#d00; font-weight:600; }
    .stDataFrame { font-size:0.8rem; }
</style>
""", unsafe_allow_html=True)

# ── Session state ──
for key in ["dados", "log_renomeacao", "classificacao", "cobertura_72", "base_indexada"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Helpers ──
def fmt_br(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "\u2014"
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_br_int(val):
    return fmt_br(val)

def kpi_card(label, value, delta=None, color=""):
    style = f"color:{color};" if color else ""
    html = f"""<div class="kpi-card">
        <div class="kpi-value" style="{style}">{value}</div>
        <div class="kpi-label">{label}</div>
    </div>"""
    return html

def status_html(s):
    if s == "conciliado":
        return f'<span class="status-ok">{s}</span>'
    elif s in ("divergente",):
        return f'<span class="status-err">{s}</span>'
    else:
        return f'<span class="status-warn">{s}</span>'

CAT_CORES = {
    "COTA_CONDOMINIAL": "#1f77b4", "AGUA": "#2ca02c", "GAS": "#ff7f0e",
    "FUNDO_RESERVA": "#9467bd", "BENFEITORIAS": "#8c564b", "SALAO_FESTAS": "#e377c2",
    "MULTAS": "#d62728", "ACORDOS": "#7f7f7f", "MANUT_PINTURA": "#bcbd22",
    "MANUTENCOES_DIVERSAS": "#17becf", "TAXA_ADMINISTRACAO": "#aec7e8",
    "CORRECAO_MONETARIA": "#98df8a", "RECEITAS_EVENTUAIS": "#ffbb78",
    "RECEITAS_FUNDO_RESERVA": "#c5b0d5", "TRANSFERENCIA_ENTRE_CONTAS": "#c49c94",
    "RENDIMENTO_APLICACAO": "#f7b6d2", "APLICACAO_BTG": "#c7c7c7",
    "ATRASO_COTA_CONDOMINIAL": "#dbdb8d", "ATRASO_BENFEITORIAS": "#9edae5",
    "ATRASO_FUNDO_RESERVA": "#f5b041", "ATRASO_SALAO_FESTAS": "#e74c3c",
    "CONDOMINOS_EM_ATRASO": "#e74c3c",
}

CAT_ORDER = [
    "COTA_CONDOMINIAL", "AGUA", "GAS", "FUNDO_RESERVA", "BENFEITORIAS",
    "SALAO_FESTAS", "MULTAS", "ACORDOS", "MANUT_PINTURA",
    "MANUTENCOES_DIVERSAS", "TAXA_ADMINISTRACAO",
]

def _salvar_upload(uploaded_file, destino_dir: Path) -> str | None:
    """Salva arquivo enviado em destino_dir, retorna caminho salvo."""
    destino_dir.mkdir(parents=True, exist_ok=True)
    caminho = destino_dir / uploaded_file.name
    with open(caminho, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return str(caminho)

def _detectar_competencia_rapido(caminho: str) -> str | None:
    """Tenta detectar competência de forma simples para exibição."""
    from src.services.competencia_detector import detectar_competencia
    comp, _, _ = detectar_competencia(caminho)
    return comp

def _extensoes_validas(tipo: str) -> list[str]:
    return {"extrato": [".xls"], "planilha": [".xlsx"], "pdf": [".pdf"], "lote": [".xls", ".xlsx", ".pdf"]}.get(tipo, [])

# ── Pasta oficial ──
DATA_RENAMED = ROOT / "data" / "renamed"

def _carregar_historico_renamed() -> tuple:
    """Lê TODOS os arquivos renomeados (histórico acumulado) de data/renamed/."""
    from src.services.unidade_classifier import ler_todos_extratos
    extratos = ler_todos_extratos(DATA_RENAMED, str(CONFIG_MAPA))
    todos_lancamentos = []
    todos_resumos = []
    todas_nao_int = []
    extratos_tuples = []
    todas_aud = []
    for comp, lancs, resums, nao_int in extratos:
        lancs_val, df_aud = validar_lancamentos_unidades(lancs, UNIDADES_VALIDAS)
        if not df_aud.empty:
            todas_aud.append(df_aud)
        extratos_tuples.append((comp, lancs_val, resums, nao_int))
        todos_lancamentos.extend(lancs_val)
        todos_resumos.extend(resums)
        todas_nao_int.extend(nao_int)
    df_aud_final = pd.concat(todas_aud).reset_index(drop=True) if todas_aud else pd.DataFrame()
    return todos_lancamentos, todos_resumos, todas_nao_int, extratos_tuples, df_aud_final

def _carregar_do_consolidado(caminho: Path) -> tuple:
    """Carrega dados do consolidado Excel (gerado por consolidar_extratos.py).
    Retorna (todos_lancamentos, todos_resumos, todas_nao_int, extratos_tuples, df_aud_final)."""
    from src.models import LancamentoExtrato

    df = pd.read_excel(caminho, sheet_name="Extrato Consolidado", dtype={"unidade": str})

    todos_lancamentos = []
    for _, row in df.iterrows():
        data_lanc = str(row.get("data_lancamento", "")) if pd.notna(row.get("data_lancamento")) else ""
        data_doc = str(row.get("data_documento", "")) if pd.notna(row.get("data_documento")) else ""
        tipo = str(row.get("tipo_movimento", ""))
        raw_valor = row.get("valor", 0)
        valor = abs(raw_valor) if pd.notna(raw_valor) else 0.0

        raw_unid = row.get("unidade")
        unidade = str(int(raw_unid)) if pd.notna(raw_unid) and str(raw_unid).strip() not in ("", "nan") else None

        lanc = LancamentoExtrato(
            competencia=str(row.get("competencia", "")),
            data_lancamento=data_lanc,
            unidade=unidade,
            codigo_conta=str(row.get("codigo_conta", "")),
            descricao_conta=str(row.get("descricao_conta", "")),
            categoria=str(row.get("categoria", "")) or None,
            historico_bruto=str(row.get("historico", "")),
            referencia_rb=str(row.get("documento", "")),
            vencimento=data_doc,
            valor=valor,
            tipo=tipo,
            arquivo_origem=str(row.get("arquivo_origem", "")),
            linha_origem=0,
            observacao_parse=str(row.get("observacao_parse", "")),
        )
        todos_lancamentos.append(lanc)

    lancs_val, df_aud = validar_lancamentos_unidades(todos_lancamentos, UNIDADES_VALIDAS)

    comps = sorted({l.competencia for l in lancs_val})
    extratos_tuples = [(comp, [l for l in lancs_val if l.competencia == comp], [], []) for comp in comps]

    return lancs_val, [], [], extratos_tuples, df_aud


def _contas_nao_mapeadas(lancamentos):
    from collections import Counter
    c = Counter()
    for l in lancamentos:
        if not l.categoria:
            c[(l.competencia, l.codigo_conta, l.descricao_conta)] += 1
    if not c:
        return pd.DataFrame()
    rows = []
    for (comp, cod, desc), qtd in c.items():
        rows.append({"competencia": comp, "codigo": cod, "descricao": desc, "ocorrencias": qtd})
    return pd.DataFrame(rows).sort_values("ocorrencias", ascending=False)

def _df_nao_interpretadas(lista):
    if not lista:
        return pd.DataFrame()
    return pd.DataFrame([vars(x) for x in lista])

def _processar_historico(todos_lancamentos, todos_resumos, todas_nao_int, extratos_tuples, df_unidades_auditoria, mostrar_status=True, aplicar_validacoes=False):
    """Pipeline comum: normaliza, analisa cobertura, concilia, classifica unidades, caixa local."""
    def _log(msg):
        if mostrar_status:
            st.write(msg)

    _log("Normalizando...")
    df_lanc = gerar_lancamentos_normalizados(todos_lancamentos)
    df_lanc = ensure_columns(
        df_lanc,
        ["competencia", "unidade", "categoria", "valor", "codigo_conta",
         "descricao_conta", "historico_bruto", "arquivo_origem",
         "data_lancamento", "referencia_rb", "vencimento"],
        stage="pipeline_lancamentos",
    )

    # ── Apply manual validations if requested ──
    if aplicar_validacoes:
        from src.services.manual_validation import aplicar_validacoes as _aplicar_val
        df_lanc = _aplicar_val(df_lanc)
        # Filter excluded items from the raw lancamento list
        ids_excluidos = set(df_lanc[df_lanc["incluir_no_calculo"] == "nao"].index)
        # Also record the original list for impact tracking
        _log(f"Validações aplicadas: {len(ids_excluidos)} registros excluídos")
    else:
        df_lanc["status_validacao"] = "pendente"
        df_lanc["origem_ajuste_manual"] = ""
        df_lanc["incluir_no_calculo"] = "sim"
        df_lanc["chave_deduplicacao"] = ""

    df_rec_det, df_rec_pivot = separar_recebiveis_por_unidade(todos_lancamentos)
    df_rec_det = ensure_columns(
        df_rec_det,
        ["competencia", "unidade", "categoria", "valor", "codigo_conta",
         "descricao_conta", "historico_bruto", "arquivo_origem",
         "data_lancamento", "referencia_rb", "vencimento"],
        stage="pipeline_recebiveis_detalhe",
    )
    resumo_por_conta = gerar_resumo_mensal_por_conta(todos_lancamentos)
    df_resumo_ext = consolidar_resumo_mensal(todos_resumos + resumo_por_conta)

    # ── Synthesize RESUMO rows if missing (consolidated Excel doesn't provide them) ──
    if not df_resumo_ext.empty and "RESUMO" not in df_resumo_ext["conta_normalizada"].values:
        _log("Sintetizando linhas RESUMO a partir dos lançamentos...")
        resumo_rows = []
        for comp in sorted(df_resumo_ext["competencia"].unique()):
            sub = df_resumo_ext[df_resumo_ext["competencia"] == comp]
            creditos = sub["creditos"].sum() if "creditos" in sub.columns else None
            debitos = sub["debitos"].sum() if "debitos" in sub.columns else None
            resumo_rows.append({
                "competencia": comp,
                "conta_origem": "RESUMO",
                "conta_normalizada": "RESUMO",
                "saldo_inicial": None,
                "creditos": creditos,
                "debitos": debitos,
                "saldo_final": None,
                "arquivo_origem": "sintetizado",
                "observacao": "sintetizado_da_base_consolidada",
            })
        if resumo_rows:
            df_resumo_sint = pd.DataFrame(resumo_rows)
            df_resumo_ext = pd.concat([df_resumo_ext, df_resumo_sint], ignore_index=True)
            df_resumo_ext = df_resumo_ext.sort_values(["competencia", "conta_normalizada"]).reset_index(drop=True)
            _log(f"  {len(resumo_rows)} linhas RESUMO adicionadas.")

    _log("Analisando cobertura água/gás...")
    resultado_cobertura = analisar_cobertura(df_lanc)
    df_cobertura = resultado_cobertura["tabela_mensal"]
    df_cobertura_unidade = resultado_cobertura["detalhe_unidade"]

    _log("Montando espelho do boleto por unidade...")
    df_espelho = gerar_espelho_boleto(df_rec_det)
    df_espelho_consolidado = gerar_espelho_consolidado_por_unidade(df_espelho)

    df_resumo_plan = None
    planilhas_rename = sorted(DATA_RENAMED.glob("planilha_resumo_*.xlsx"))
    if planilhas_rename:
        pp = ParserPlanilha(str(CONFIG_MAPA))
        res_p, _ = pp.parse_resumo(planilhas_rename[-1])
        df_resumo_plan = consolidar_resumo_mensal(res_p)
        _log(f"Planilha: {len(res_p)} registros ({planilhas_rename[-1].name})")

    _log("Conciliando...")
    df_conc = conciliar(df_resumo_ext, df_resumo_plan)

    _log("Classificando unidades...")
    unidades_esp = carregar_unidades_esperadas(CONFIG_UNIDADES)
    comp_dict = gerar_inventario_unidades(extratos_tuples)
    df_classif = classificar_unidades(unidades_esp, comp_dict)
    df_cob = gerar_cobertura_72(unidades_esp, comp_dict)
    df_inad = df_classif[df_classif["status_final"] == "INADIMPLENTE_PRESUMIDO"].copy() if not df_classif.empty else pd.DataFrame()

    _log("Analisando caixa local (PDFs)...")
    from src.services.caixa_local_matcher import processar_pdfs_despesas_locais
    pdf_data = processar_pdfs_despesas_locais(DATA_RENAMED)
    if pdf_data and mostrar_status:
        for comp, info in pdf_data.items():
            st.write(f"  PDF comp={comp}: {len(info['lancamentos'])} lançamentos")

    df_caixa_local = conciliar_caixa_local(todos_lancamentos, pdf_data)
    df_resumo_caixa = gerar_resumo_caixa_local(df_caixa_local)
    df_caixa_sem_match = df_caixa_local[df_caixa_local["status_match_pdf"] == "NAO_ENCONTRADO"].copy() if not df_caixa_local.empty else pd.DataFrame()

    contas_nao_map = _contas_nao_mapeadas(todos_lancamentos)
    df_nao_int = _df_nao_interpretadas(todas_nao_int)

    # ── Export audit CSV files ──
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        df_esp = df_espelho if isinstance(df_espelho, pd.DataFrame) and not df_espelho.empty else None
        if df_esp is not None:
            aud = df_esp.attrs.get("auditoria", [])
            if aud:
                pd.DataFrame(aud).to_csv(OUTPUT_DIR / "auditoria_espelho_unidade.csv", index=False, encoding="utf-8-sig")
            dup = df_esp.attrs.get("duplicatas", [])
            if dup:
                pd.DataFrame(dup).to_csv(OUTPUT_DIR / "itens_duplicados_unidade.csv", index=False, encoding="utf-8-sig")
            disc = df_esp.attrs.get("discrepancias", [])
            if disc:
                pd.DataFrame(disc).to_csv(OUTPUT_DIR / "valores_discrepantes_unidade.csv", index=False, encoding="utf-8-sig")
    except Exception as e:
        import logging
        logging.getLogger(__name__).warning("Erro ao exportar audit CSVs: %s", e)

    return {
        "lancamentos": df_lanc,
        "recebiveis_detalhe": df_rec_det,
        "resumo_unidade": df_rec_pivot,
        "resumo_mensal": df_resumo_ext,
        "resumo_planilha": df_resumo_plan,
        "conciliacao": df_conc,
        "classificacao_unidades": df_classif,
        "cobertura_72": df_cob,
        "cobertura_agua_gas": df_cobertura,
        "cobertura_agua_gas_detalhe": df_cobertura_unidade,
        "espelho_boleto": df_espelho,
        "espelho_boleto_consolidado": df_espelho_consolidado,
        "inadimplentes_presumidos": df_inad,
        "unidades_auditoria": df_unidades_auditoria,
        "caixa_local": df_caixa_local,
        "resumo_caixa": df_resumo_caixa,
        "caixa_sem_match": df_caixa_sem_match,
        "contas_nao_mapeadas": contas_nao_map,
        "nao_interpretadas": df_nao_int,
    }

# ── Auto-load na inicializa├º├úo ──
if st.session_state.dados is None and DATA_RENAMED.exists():
    with st.spinner("Carregando base histórica..."):
        try:
            from src.services.file_indexer import escanear_pasta
            base_idx = escanear_pasta(DATA_RENAMED)
            st.session_state.base_indexada = base_idx
            if base_idx.total_arquivos > 0:
                if CONSOLIDADO_PATH.exists():
                    lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist = _carregar_do_consolidado(CONSOLIDADO_PATH)
                else:
                    lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist = _carregar_historico_renamed()
                dados = _processar_historico(lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist, mostrar_status=False)
                if dados is not None:
                    st.session_state.dados = dados
        except Exception as e:
            st.warning(f"Aviso no auto-load: {e}")

# ── Sidebar ──
st.sidebar.title("\U0001f3e0 Condom\u00ednio")
st.sidebar.markdown("Dashboard de conciliação financeira")
st.sidebar.markdown(f"**Pasta oficial:** `data/renamed/`")
fonte = "📄 Consolidado Excel" if CONSOLIDADO_PATH.exists() else "📂 PDFs individuais"
st.sidebar.caption(f"Carregado via: {fonte}")

# ── Status da base ──
st.sidebar.markdown("---")
st.sidebar.subheader("Base Histórica")
if st.session_state.base_indexada is not None:
    bi = st.session_state.base_indexada
    st.sidebar.markdown(f"**{bi.total_arquivos}** arquivos indexados")
    st.sidebar.markdown(f"**{len(bi.competencias)}** competências")
    if bi.erros:
        st.sidebar.warning(f"{len(bi.erros)} erros")

# ── Novos arquivos (importação opcional) ──
st.sidebar.markdown("---")
with st.sidebar.expander("📥 Importar novos arquivos"):
    novos = st.file_uploader(
        "Selecionar arquivos",
        type=["xls", "xlsx", "pdf"],
        accept_multiple_files=True,
        key="import_uploader",
        help="Extratos (.xls), planilha (.xlsx) e PDFs de despesas locais (.pdf)",
    )
    if novos:
        for f in novos:
            ext = Path(f.name).suffix.lower()
            tipo = {"xls": "Extrato", "xlsx": "Planilha", "pdf": "PDF Desp. Locais"}.get(ext.lstrip("."), "?")
            size_kb = len(f.getvalue()) / 1024
            st.sidebar.markdown(f"- {f.name}  _({tipo}, {size_kb:.0f} KB)_")

        if st.sidebar.button("Importar e processar", type="primary", use_container_width=True):
            _preparar_e_renomear(novos)
            lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist = _carregar_historico_renamed()
            dados = _processar_historico(lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist, mostrar_status=False)
            if dados is not None:
                st.session_state.dados = dados
                from src.services.file_indexer import escanear_pasta
                st.session_state.base_indexada = escanear_pasta(DATA_RENAMED)
                st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Atualizar base", use_container_width=True):
    if DATA_RENAMED.exists():
        with st.spinner("Reescaneando pasta..."):
            from src.services.file_indexer import escanear_pasta
            st.session_state.base_indexada = escanear_pasta(DATA_RENAMED)
            if st.session_state.base_indexada.total_arquivos > 0:
                if CONSOLIDADO_PATH.exists():
                    lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist = _carregar_do_consolidado(CONSOLIDADO_PATH)
                else:
                    lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist = _carregar_historico_renamed()
                dados = _processar_historico(lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist, mostrar_status=False)
                if dados is not None:
                    st.session_state.dados = dados
            st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🔁 Recalcular base com validações", use_container_width=True):
    if st.session_state.dados is not None:
        with st.spinner("Reaplicando validações e recalculando..."):
            lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist = (
                _carregar_do_consolidado(CONSOLIDADO_PATH)
                if CONSOLIDADO_PATH.exists()
                else _carregar_historico_renamed()
            )
            dados = _processar_historico(
                lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist,
                mostrar_status=False, aplicar_validacoes=True,
            )
            if dados is not None:
                st.session_state.dados = dados
        st.rerun()

# ── Helper: preparar e renomear novos arquivos ──
def _preparar_e_renomear(arquivos):
    """Salva uploads em data/renamed/ via renomeação."""
    DATA_UPLOADS = ROOT / "data" / "uploads"
    DATA_UPLOADS.mkdir(parents=True, exist_ok=True)
    DATA_RENAMED.mkdir(parents=True, exist_ok=True)
    for f in arquivos:
        _salvar_upload(f, DATA_UPLOADS)
    log_ren = varrer_e_renomear(DATA_UPLOADS, DATA_RENAMED)
    st.session_state.log_renomeacao = pd.DataFrame(log_ren) if log_ren else pd.DataFrame()
    for f in DATA_UPLOADS.glob("*"):
        if f.is_file() and f.name != ".gitkeep":
            f.unlink()
    return log_ren





def _validar_totalizacao(resumo_df: pd.DataFrame) -> pd.DataFrame:
    """Verifica: saldo_inicial + créditos - débitos = saldo_final para cada competência."""
    if resumo_df.empty:
        return pd.DataFrame()
    resumo_rows = resumo_df[resumo_df["conta_normalizada"] == "RESUMO"].copy()
    if resumo_rows.empty:
        return pd.DataFrame()
    resumo_rows["saldo_calculado"] = (
        resumo_rows["saldo_inicial"].fillna(0)
        + resumo_rows["creditos"].fillna(0)
        - resumo_rows["debitos"].fillna(0)
    ).round(2)
    resumo_rows["divergencia"] = (
        (resumo_rows["saldo_calculado"] - resumo_rows["saldo_final"]).abs().round(2)
    )
    resumo_rows["status"] = resumo_rows["divergencia"].apply(
        lambda d: "✅ OK" if d < 0.01 else f"⚠️ DIVERGÊNCIA R$ {d:,.2f}"
    )
    cols = ["competencia", "saldo_inicial", "creditos", "debitos",
            "saldo_final", "saldo_calculado", "divergencia", "status"]
    return resumo_rows[[c for c in cols if c in resumo_rows.columns]]

# ── Se não houver dados, exibe tela de boas-vindas ──
if st.session_state.dados is None:
    st.title("🏠 Conciliação Financeira - Condomínio")
    st.markdown("---")

    col_w1, col_w2 = st.columns([2, 1])
    with col_w1:
        st.markdown("### Base Histórica Automática")
        st.markdown(
            "O sistema monitora automaticamente a pasta **`data/renamed/`**.\n\n"
            "Todos os extratos, planilhas e PDFs renomeados são carregados "
            "na inicialização sem necessidade de upload manual."
        )
    with col_w2:
        st.markdown(
            "### 📥 Importar novos arquivos\n"
            "Use a seção **Importar novos arquivos** no menu lateral\n"
            "para adicionar novos extratos ou PDFs ao sistema."
        )

    # Status da base no centro
    if st.session_state.base_indexada is not None:
        bi = st.session_state.base_indexada
        st.markdown("---")
        st.markdown("### 📊 Status da Base Histórica")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Arquivos Válidos", bi.total_arquivos)
        m2.metric("Competências", len(bi.competencias))
        m3.metric("Erros", len(bi.erros))
        m4.metric("Duplicatas Ignoradas", len(bi.duplicatas))

        if bi.arquivos:
            with st.expander("📄 Arquivos indexados"):
                st.dataframe(
                    [{
                        "Tipo": a.tipo,
                        "Competência": a.competencia,
                        "Arquivo": a.nome_renomeado,
                        "Tamanho (KB)": a.tamanho_kb,
                        "Status": a.status,
                    } for a in bi.arquivos],
                    use_container_width=True, hide_index=True,
                )

        if bi.erros:
            with st.expander("⚠️ Arquivos com erro"):
                st.dataframe(
                    [{"Arquivo": e.nome_renomeado, "Erro": e.mensagem} for e in bi.erros],
                    use_container_width=True, hide_index=True,
                )

        if bi.duplicatas:
            with st.expander("🔁 Arquivos duplicados ignorados"):
                st.dataframe(
                    [{
                        "Arquivo": d.nome_renomeado,
                        "Motivo": d.mensagem,
                    } for d in bi.duplicatas],
                    use_container_width=True, hide_index=True,
                )

        # Exportar índice
        if bi.arquivos:
            csv_idx = pd.DataFrame([
                {"Tipo": a.tipo, "Competência": a.competencia,
                 "Arquivo": a.nome_renomeado, "Hash": a.hash_md5,
                 "Tamanho_KB": a.tamanho_kb, "Data": a.data_inclusao}
                for a in bi.arquivos
            ]).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "📥 Exportar índice de arquivos (CSV)",
                data=csv_idx,
                file_name="file_index.csv",
                mime="text/csv",
                use_container_width=True,
            )
    else:
        st.info(
            "Nenhuma base encontrada em `data/renamed/`. "
            "Use a seção **Importar novos arquivos** no menu lateral "
            "para adicionar os primeiros extratos."
        )

    st.markdown("---")
    st.caption(
        f"Última atualização: "
        f"{st.session_state.base_indexada.ultima_atualizacao if st.session_state.base_indexada else '—'}"
    )
    st.stop()

d = st.session_state.dados

# ── Filtros globais ──
st.sidebar.markdown("---")
st.sidebar.subheader("Filtros")

todas_comps = sorted(d["resumo_mensal"]["competencia"].unique()) if not d["resumo_mensal"].empty else []
comp_sel = st.sidebar.selectbox("Competência", ["Todas"] + todas_comps)

unidades_disponiveis = sorted(
    u for u in d["resumo_unidade"]["unidade"].unique()
    if u.isdigit() and int(u) in UNIDADES_VALIDAS
) if not d["resumo_unidade"].empty else []
unidade_sel = st.sidebar.selectbox("Unidade", ["Todas"] + unidades_disponiveis)

cats = sorted(d["recebiveis_detalhe"]["categoria"].unique()) if not d["recebiveis_detalhe"].empty else []
cat_sel = st.sidebar.selectbox("Categoria", ["Todas"] + cats)

def filtrar(df, comp_col="competencia", unid_col="unidade", cat_col="categoria"):
    if df is None or df.empty:
        return df
    f = df.copy()
    if comp_sel != "Todas" and comp_col in f.columns:
        f = f[f[comp_col] == comp_sel]
    if unidade_sel != "Todas" and unid_col in f.columns:
        f = f[f[unid_col] == unidade_sel]
    if cat_sel != "Todas" and cat_col in f.columns:
        f = f[f[cat_col] == cat_sel]
    return f

# ── Tabs ──
tab_visao, tab_cobertura, tab_unidades, tab_categoria, tab_conciliacao, \
tab_auditoria, tab_caixa, tab_espelho, tab_extrato_cons, tab_validacao = st.tabs([
    "\U0001f4ca Visão Geral",
    "\U0001f3e0 Cobertura 72 Unidades",
    "\U0001f4cb Recebíveis por Unidade",
    "\U0001f4ca Análise por Categoria",
    "\U0001f504 Conciliação",
    "\U0001f50d Auditoria",
    "\U0001f4b0 Caixa Local",
    "\U0001f4c3 Espelho do Boleto",
    "\U0001f4c4 Extrato Consolidado",
    "\U0000270f Validação Manual",
])

# ═══════════════════════════════════════════════════════════════
# TAB 1 - VISÃO GERAL
# ═══════════════════════════════════════════════════════════════
with tab_visao:
    st.subheader("Visão Geral")

    # ── Metadata ──
    df_lanc_visao = d.get("lancamentos")
    df_resumo_visao = d["resumo_mensal"]

    num_competencias = df_lanc_visao["competencia"].nunique() if df_lanc_visao is not None and not df_lanc_visao.empty else 0
    num_lancamentos = len(df_lanc_visao) if df_lanc_visao is not None else 0
    num_unidades = df_lanc_visao["unidade"].nunique() if df_lanc_visao is not None and not df_lanc_visao.empty else 0
    fonte = "📄 Consolidado Excel" if CONSOLIDADO_PATH.exists() else "📂 PDFs individuais"

    col_info1, col_info2, col_info3, col_info4 = st.columns(4)
    with col_info1:
        st.markdown(kpi_card("Competência Ativa", comp_sel if comp_sel != "Todas" else "Todas"), unsafe_allow_html=True)
    with col_info2:
        st.markdown(kpi_card("Lançamentos", f"{num_lancamentos:,}"), unsafe_allow_html=True)
    with col_info3:
        st.markdown(kpi_card("Competências Disp.", str(num_competencias)), unsafe_allow_html=True)
    with col_info4:
        st.markdown(kpi_card("Fonte", fonte), unsafe_allow_html=True)

    st.markdown("---")

    # ── KPIs financeiros ──
    resumo_row = df_resumo_visao
    resumo_row = resumo_row[resumo_row["conta_normalizada"] == "RESUMO"] if not resumo_row.empty else pd.DataFrame()
    if comp_sel != "Todas" and not resumo_row.empty:
        resumo_row = resumo_row[resumo_row["competencia"] == comp_sel]

    has_resumo = not resumo_row.empty

    if has_resumo:
        saldo_ini = resumo_row["saldo_inicial"].iloc[0] if "saldo_inicial" in resumo_row.columns and pd.notna(resumo_row["saldo_inicial"].iloc[0]) else None
        creditos = resumo_row["creditos"].iloc[0] if "creditos" in resumo_row.columns and pd.notna(resumo_row["creditos"].iloc[0]) else None
        debitos = resumo_row["debitos"].iloc[0] if "debitos" in resumo_row.columns and pd.notna(resumo_row["debitos"].iloc[0]) else None
        saldo_fim = resumo_row["saldo_final"].iloc[0] if "saldo_final" in resumo_row.columns and pd.notna(resumo_row["saldo_final"].iloc[0]) else None
        kpi_origin = "resumo_mensal"
    elif df_lanc_visao is not None and not df_lanc_visao.empty:
        lanc_filt = df_lanc_visao.copy()
        if comp_sel != "Todas" and "competencia" in lanc_filt.columns:
            lanc_filt = lanc_filt[lanc_filt["competencia"] == comp_sel]
        creditos_totais = lanc_filt[lanc_filt["tipo"] == "credito"]["valor"].sum() if "tipo" in lanc_filt.columns and "valor" in lanc_filt.columns else None
        debitos_totais = lanc_filt[lanc_filt["tipo"] == "debito"]["valor"].sum() if "tipo" in lanc_filt.columns and "valor" in lanc_filt.columns else None
        saldo_ini = None
        creditos = round(creditos_totais, 2) if creditos_totais is not None else None
        debitos = round(debitos_totais, 2) if debitos_totais is not None else None
        saldo_fim = round(creditos - debitos, 2) if creditos is not None and debitos is not None else None
        kpi_origin = "lancamentos (fallback)"
    else:
        saldo_ini = creditos = debitos = saldo_fim = None
        kpi_origin = "sem dados"

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        val = fmt_br(saldo_ini) if saldo_ini is not None else "—"
        st.markdown(kpi_card("Saldo Inicial", val), unsafe_allow_html=True)
    with col2:
        val = fmt_br(creditos) if creditos is not None else "—"
        st.markdown(kpi_card("Créditos", val, color="green"), unsafe_allow_html=True)
    with col3:
        val = fmt_br(debitos) if debitos is not None else "—"
        st.markdown(kpi_card("Débitos", val, color="red"), unsafe_allow_html=True)
    with col4:
        val = fmt_br(saldo_fim) if saldo_fim is not None else "—"
        st.markdown(kpi_card("Saldo Final", val), unsafe_allow_html=True)

    if all(v is None for v in [saldo_ini, creditos, debitos, saldo_fim]):
        if df_lanc_visao is None or df_lanc_visao.empty:
            st.warning("Base consolidada não carregada. Importe ou carregue a base histórica primeiro.")
        elif comp_sel != "Todas" and "competencia" in df_lanc_visao.columns and comp_sel not in df_lanc_visao["competencia"].values:
            st.warning(f"Sem dados para a competência selecionada: {comp_sel}")
        else:
            st.warning("Colunas obrigatórias ausentes para calcular KPIs.")

    st.caption(f"Origem dos KPIs: {kpi_origin}")

    # ── Validação de totalização ──
    df_totalizacao = _validar_totalizacao(d["resumo_mensal"])
    if not df_totalizacao.empty:
        st.markdown("---")
        for _, r in df_totalizacao.iterrows():
            if "⚠️" in str(r["status"]):
                st.error(f"**Totalização {r['competencia']}:** {r['status']}")
            else:
                st.success(f"**Totalização {r['competencia']}:** {r['status']}")

    st.markdown("---")

    # ── Cobertura das unidades ──
    if "classificacao_unidades" in d and d["classificacao_unidades"] is not None and not d["classificacao_unidades"].empty:
        df_class_f = d["classificacao_unidades"]
        if comp_sel != "Todas":
            df_class_f = df_class_f[df_class_f["competencia"] == comp_sel]

        total_unids = df_class_f["unidade"].nunique() if not df_class_f.empty else 0
        localizadas = len(df_class_f[df_class_f["apareceu_mes_atual"] == "sim"]) if "apareceu_mes_atual" in df_class_f.columns else 0
        inadimplentes = len(df_class_f[df_class_f["status_final"] == "INADIMPLENTE_PRESUMIDO"]) if "status_final" in df_class_f.columns else 0
        antecipados = len(df_class_f[df_class_f["status_final"] == "PAGO_ANTECIPADO"]) if "status_final" in df_class_f.columns else 0

        st.subheader("Cobertura das Unidades")
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            st.markdown(kpi_card("Unidades", str(total_unids)), unsafe_allow_html=True)
        with col_s2:
            st.markdown(kpi_card("Localizadas", str(localizadas), color="green"), unsafe_allow_html=True)
        with col_s3:
            st.markdown(kpi_card("Adiantamentos", str(antecipados), color="orange"), unsafe_allow_html=True)
        with col_s4:
            st.markdown(kpi_card("Inadimplentes Presumidos", str(inadimplentes), color="red"), unsafe_allow_html=True)
    elif df_lanc_visao is not None and not df_lanc_visao.empty:
        uni_count = df_lanc_visao["unidade"].nunique() if "unidade" in df_lanc_visao.columns else 0
        col_s1, _ = st.columns(2)
        with col_s1:
            st.markdown(kpi_card("Unidades no extrato", str(uni_count)), unsafe_allow_html=True)
    else:
        # Modo único: mostra info de unidades do resumo
        uni_count = d["resumo_unidade"]["unidade"].nunique() if not d["resumo_unidade"].empty else 0
        col_s1, _ = st.columns(2)
        with col_s1:
            st.markdown(kpi_card("Unidades no extrato", str(uni_count)), unsafe_allow_html=True)

    st.markdown("---")

    # ── Resumo por categoria ──
    df_res_cat = d["resumo_mensal"]
    df_res_cat = df_res_cat[df_res_cat["conta_normalizada"] != "RESUMO"] if not df_res_cat.empty else df_res_cat
    df_res_cat_f = filtrar(df_res_cat)

    if not df_res_cat_f.empty:
        st.subheader("Resumo por Categoria")
        col_a, col_b = st.columns(2)
        with col_a:
            if "creditos" in df_res_cat_f.columns and df_res_cat_f["creditos"].notna().any():
                fig_cred = px.bar(
                    df_res_cat_f.sort_values("creditos", ascending=False),
                    x="conta_normalizada", y="creditos",
                    title="Créditos por Categoria",
                    color="conta_normalizada",
                    color_discrete_map=CAT_CORES,
                    text_auto=".2f",
                )
                fig_cred.update_layout(xaxis_title="", yaxis_title="R$", showlegend=False)
                st.plotly_chart(fig_cred, use_container_width=True)
        with col_b:
            if "debitos" in df_res_cat_f.columns and df_res_cat_f["debitos"].notna().any():
                fig_deb = px.bar(
                    df_res_cat_f.sort_values("debitos", ascending=False),
                    x="conta_normalizada", y="debitos",
                    title="Débitos por Categoria",
                    color="conta_normalizada",
                    color_discrete_map=CAT_CORES,
                    text_auto=".2f",
                )
                fig_deb.update_layout(xaxis_title="", yaxis_title="R$", showlegend=False)
                st.plotly_chart(fig_deb, use_container_width=True)

        st.markdown("---")
        st.subheader("Tabela de Totais por Categoria")
        st.dataframe(df_res_cat_f, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum dado de resumo por categoria disponível para os filtros atuais.")

    if not d["conciliacao"].empty:
        st.markdown("---")
        st.subheader("Status da Conciliação")
        df_conc_f = filtrar(d["conciliacao"], comp_col="competencia", unid_col="conta_normalizada", cat_col="conta_normalizada")
        conc_ok = len(df_conc_f[df_conc_f["status_conciliacao"] == "conciliado"])
        conc_div = len(df_conc_f[df_conc_f["status_conciliacao"] == "divergente"])
        conc_only_e = len(df_conc_f[df_conc_f["status_conciliacao"] == "apenas_no_extrato"])
        conc_only_p = len(df_conc_f[df_conc_f["status_conciliacao"] == "apenas_na_planilha"])
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            st.markdown(kpi_card("Conciliados", str(conc_ok), color="green"), unsafe_allow_html=True)
        with col_s2:
            st.markdown(kpi_card("Divergentes", str(conc_div), color="red"), unsafe_allow_html=True)
        with col_s3:
            st.markdown(kpi_card("Só no Extrato", str(conc_only_e), color="orange"), unsafe_allow_html=True)
        with col_s4:
            st.markdown(kpi_card("Só na Planilha", str(conc_only_p), color="orange"), unsafe_allow_html=True)

    # ── Debug ──
    with st.expander("🔧 Debug - Visão Geral", expanded=False):
        st.markdown("**DataFrames disponíveis:**")
        st.write(f"- `lancamentos`: shape={df_lanc_visao.shape if df_lanc_visao is not None else 'None'}, cols={list(df_lanc_visao.columns) if df_lanc_visao is not None else 'N/A'}")
        st.write(f"- `resumo_mensal`: shape={df_resumo_visao.shape if df_resumo_visao is not None else 'None'}")
        st.write(f"- `resumo_unidade`: shape={d['resumo_unidade'].shape if d['resumo_unidade'] is not None else 'None'}")
        st.write(f"- `classificacao_unidades`: shape={d.get('classificacao_unidades').shape if d.get('classificacao_unidades') is not None else 'None'}")
        st.write(f"**Filtro:** competência={comp_sel}, unidade={unidade_sel}, categoria={cat_sel}")
        st.write(f"**RESUMO rows encontrados:** {len(resumo_row) if not resumo_row.empty else 0}")
        st.write(f"**Origem KPI:** {kpi_origin}")
        if df_lanc_visao is not None and not df_lanc_visao.empty:
            st.write("**Primeiras linhas (lancamentos):**")
            st.dataframe(df_lanc_visao.head(5), use_container_width=True)
        from src.utils import COLUMN_AUDIT_LOG, relatorio_colunas_ausentes
        df_aud_col = relatorio_colunas_ausentes()
        if not df_aud_col.empty:
            st.write("**Auditoria de colunas ausentes:**")
            st.dataframe(df_aud_col, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 2 - COBERTURA 72 UNIDADES
# ═══════════════════════════════════════════════════════════════
with tab_cobertura:
    st.subheader("Cobertura das 72 Unidades por Competência")

    if "classificacao_unidades" not in d or d["classificacao_unidades"] is None or d["classificacao_unidades"].empty:
        st.warning("Disponível apenas no modo Lote. Carregue os dados com data/raw/ contendo múltiplos extratos.")
    else:
        df_class = d["classificacao_unidades"].copy()
        if comp_sel != "Todas" and "competencia" in df_class.columns:
            df_class = df_class[df_class["competencia"] == comp_sel]

        # Filtro por status
        status_opts = ["Todos"] + sorted(df_class["status_final"].unique())
        status_filtro = st.selectbox("Filtrar por status", status_opts)
        if status_filtro != "Todos":
            df_class = df_class[df_class["status_final"] == status_filtro]

        # KPIs
        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        with col_k1:
            st.markdown(kpi_card("Total", str(df_class["unidade"].nunique())), unsafe_allow_html=True)
        with col_k2:
            pago = len(df_class[df_class["status_final"] == "PAGO_NO_MES"])
            st.markdown(kpi_card("Pago no Mês", str(pago), color="green"), unsafe_allow_html=True)
        with col_k3:
            ant = len(df_class[df_class["status_final"] == "PAGO_ANTECIPADO"])
            st.markdown(kpi_card("Antecipado", str(ant), color="orange"), unsafe_allow_html=True)
        with col_k4:
            inad = len(df_class[df_class["status_final"] == "INADIMPLENTE_PRESUMIDO"])
            st.markdown(kpi_card("Inadimplente Presumido", str(inad), color="red"), unsafe_allow_html=True)

        st.markdown("---")

        # Tabela
        st.dataframe(
            df_class,
            use_container_width=True, hide_index=True,
            column_config={
                "total_pago_mes": st.column_config.NumberColumn("Total Pago", format="R$ %.2f"),
                "unidade_int": None,
            },
        )

        # Gráfico
        status_counts = df_class["status_final"].value_counts().reset_index()
        status_counts.columns = ["status", "quantidade"]
        fig_status = px.bar(
            status_counts, x="status", y="quantidade",
            title="Distribuição de Status",
            color="status",
            color_discrete_map={
                "PAGO_NO_MES": "green", "PAGO_ANTECIPADO": "orange",
                "EM_ACORDO": "blue", "EM_ATRASO": "red",
                "INADIMPLENTE_PRESUMIDO": "darkred", "NAO_LOCALIZADO": "gray",
            },
        )
        fig_status.update_layout(xaxis_title="", showlegend=False)
        st.plotly_chart(fig_status, use_container_width=True)

        # Export
        csv_class = d["classificacao_unidades"].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("\U0001f4e5 Exportar CSV (classificacao_unidades.csv)", data=csv_class,
                           file_name="classificacao_unidades.csv", mime="text/csv")

# ═══════════════════════════════════════════════════════════════
# TAB 3 - RECEBÍVEIS POR UNIDADE
# ═══════════════════════════════════════════════════════════════
with tab_unidades:
    st.subheader("Recebíveis por Unidade")

    df_uni = d["resumo_unidade"]
    if df_uni.empty:
        st.warning("Nenhum recebível por unidade encontrado.")
    else:
        df_uni_f = filtrar(df_uni, comp_col="competencia", unid_col="unidade")
        if unidade_sel != "Todas":
            df_uni_f = df_uni_f[df_uni_f["unidade"] == unidade_sel]

        valor_cols = [c for c in CAT_ORDER if c in df_uni_f.columns]

        st.dataframe(df_uni_f, use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format="R$ %.2f") for c in valor_cols})

        if not df_uni_f.empty and len(valor_cols) > 0:
            st.markdown("---")
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                top_n = min(15, len(df_uni_f))
                df_top = df_uni_f.head(top_n).copy()
                df_top["_label"] = "Un " + df_top["unidade"].astype(str)
                fig_uni = px.bar(
                    df_top.sort_values("total_pago"),
                    y="_label", x="total_pago",
                    orientation="h",
                    title=f"Top {top_n} Unidades por Total Pago",
                    color="total_pago",
                    color_continuous_scale="Blues",
                    text_auto=".2f",
                )
                fig_uni.update_layout(yaxis_title="", xaxis_title="R$", showlegend=False)
                st.plotly_chart(fig_uni, use_container_width=True)
            with col_g2:
                df_stack = df_uni_f.melt(
                    id_vars=["competencia", "unidade"],
                    value_vars=valor_cols,
                    var_name="categoria", value_name="valor",
                )
                df_stack = df_stack[df_stack["valor"] > 0]
                if not df_stack.empty:
                    fig_stack = px.bar(
                        df_stack, x="unidade", y="valor",
                        color="categoria", title="Distribuição por Unidade",
                        color_discrete_map=CAT_CORES,
                    )
                    fig_stack.update_layout(xaxis_title="Unidade", yaxis_title="R$")
                    st.plotly_chart(fig_stack, use_container_width=True)

        st.markdown("---")
        st.subheader("Lançamentos Detalhados")

        if not d["recebiveis_detalhe"].empty:
            df_rec_f = filtrar(d["recebiveis_detalhe"])
            if unidade_sel != "Todas":
                df_rec_f = df_rec_f[df_rec_f["unidade"] == unidade_sel]
            if comp_sel != "Todas":
                df_rec_f = df_rec_f[df_rec_f["competencia"] == comp_sel]

            st.dataframe(df_rec_f, use_container_width=True, hide_index=True)
            if not df_rec_f.empty:
                total_rec = df_rec_f["valor"].sum()
                st.markdown(f"**Total:** {fmt_br(total_rec)}")

        csv_data = df_uni.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("\U0001f4e5 Exportar CSV (resumo_unidade.csv)", data=csv_data,
                           file_name="resumo_unidade.csv", mime="text/csv")

# ═══════════════════════════════════════════════════════════════
# TAB 4 - ANÁLISE POR CATEGORIA
# ═══════════════════════════════════════════════════════════════
with tab_categoria:
    st.subheader("Análise por Categoria")

    df_rec = d["recebiveis_detalhe"]
    if df_rec.empty:
        st.warning("Nenhum recebível disponível.")
    else:
        df_rec_f = filtrar(df_rec)
        if unidade_sel != "Todas":
            df_rec_f = df_rec_f[df_rec_f["unidade"] == unidade_sel]
        if cat_sel != "Todas":
            df_rec_f = df_rec_f[df_rec_f["categoria"] == cat_sel]

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            cat_sum = df_rec_f.groupby("categoria")["valor"].sum().reset_index().sort_values("valor", ascending=False)
            fig_cat = px.bar(
                cat_sum, x="categoria", y="valor",
                title="Total por Categoria",
                color="categoria",
                color_discrete_map=CAT_CORES,
                text_auto=".2f",
            )
            fig_cat.update_layout(xaxis_title="", yaxis_title="R$", showlegend=False)
            st.plotly_chart(fig_cat, use_container_width=True)
        with col_g2:
            if unidade_sel == "Todas":
                uni_cat = df_rec_f.groupby(["unidade", "categoria"])["valor"].sum().reset_index()
                uni_cat = uni_cat[uni_cat["valor"] > 0]
                if not uni_cat.empty:
                    fig_heat = px.bar(
                        uni_cat, x="unidade", y="valor",
                        color="categoria", title="Distribuição Unidade × Categoria",
                        color_discrete_map=CAT_CORES,
                    )
                    fig_heat.update_layout(xaxis_title="Unidade", yaxis_title="R$")
                    st.plotly_chart(fig_heat, use_container_width=True)
            else:
                uni_detail = df_rec_f.groupby("categoria")["valor"].sum().reset_index().sort_values("valor", ascending=False)
                fig_pie = px.pie(
                    uni_detail, values="valor", names="categoria",
                    title=f"Distribuição - Unidade {unidade_sel}",
                    color_discrete_map=CAT_CORES,
                )
                st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("---")
        st.subheader("Lançamentos Detalhados")
        cols_detail = ["competencia", "data_lancamento", "unidade", "codigo_conta",
                       "descricao_conta", "categoria", "historico_bruto",
                       "referencia_rb", "vencimento", "valor"]
        cols_disp = [c for c in cols_detail if c in df_rec_f.columns]
        st.dataframe(df_rec_f[cols_disp], use_container_width=True, hide_index=True)

        csv_rec = df_rec.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("\U0001f4e5 Exportar CSV (recebiveis_por_unidade.csv)", data=csv_rec,
                           file_name="recebiveis_por_unidade.csv", mime="text/csv")

# ═══════════════════════════════════════════════════════════════
# TAB 5 - CONCILIAÇÃO
# ═══════════════════════════════════════════════════════════════
with tab_conciliacao:
    st.subheader("Conciliação Extrato × Planilha")

    df_conc = d["conciliacao"]
    if df_conc.empty:
        st.warning("Nenhum dado de conciliação disponível.")
    else:
        df_conc_f = filtrar(df_conc, comp_col="competencia", unid_col="conta_normalizada", cat_col="conta_normalizada")
        if unidade_sel != "Todas":
            df_conc_f = df_conc_f[df_conc_f["conta_normalizada"] == unidade_sel]

        df_display = df_conc_f.copy()
        df_display["status"] = df_conc_f["status_conciliacao"]

        st.dataframe(
            df_display,
            use_container_width=True, hide_index=True,
            column_config={
                "saldo_extrato": st.column_config.NumberColumn("Saldo Extrato", format="R$ %.2f"),
                "saldo_planilha": st.column_config.NumberColumn("Saldo Planilha", format="R$ %.2f"),
                "diferenca": st.column_config.NumberColumn("Diferença", format="R$ %.2f"),
            },
        )

        st.markdown("---")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            conc_status = df_conc_f["status_conciliacao"].value_counts().reset_index()
            conc_status.columns = ["status", "quantidade"]
            fig_status = px.bar(conc_status, x="status", y="quantidade",
                                title="Status da Conciliação",
                                color="status",
                                color_discrete_map={
                                    "conciliado": "green", "divergente": "red",
                                    "apenas_no_extrato": "orange", "apenas_na_planilha": "orange",
                                })
            fig_status.update_layout(xaxis_title="", showlegend=False)
            st.plotly_chart(fig_status, use_container_width=True)
        with col_s2:
            if "conta_normalizada" in df_conc_f.columns:
                div = df_conc_f[df_conc_f["status_conciliacao"] == "divergente"]
                if not div.empty:
                    div_sum = div.groupby("conta_normalizada")["diferenca"].sum().reset_index().sort_values("diferenca", ascending=False)
                    fig_div = px.bar(div_sum, x="conta_normalizada", y="diferenca",
                                     title="Divergências por Conta",
                                     color="diferenca", color_continuous_scale="Reds",
                                     text_auto=".2f")
                    fig_div.update_layout(xaxis_title="", showlegend=False)
                    st.plotly_chart(fig_div, use_container_width=True)

        csv_conc = df_conc.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("\U0001f4e5 Exportar CSV (conciliacao.csv)", data=csv_conc,
                           file_name="conciliacao.csv", mime="text/csv")

# ═══════════════════════════════════════════════════════════════
# TAB 6 - AUDITORIA
# ═══════════════════════════════════════════════════════════════
with tab_auditoria:
    st.subheader("Auditoria")

    tab_a0, tab_a1, tab_a2, tab_a3, tab_a4, tab_a5, tab_a6 = st.tabs([
        "\U0001f4a7 Água/Gás",
        "\U0001f4cb Contas Não Mapeadas",
        "\U0001f4dd Linhas Não Interpretadas",
        "\U0001f50d Busca Textual",
        "\U0001f4c4 Log de Renomeação",
        "\U000026a0 Divergências",
        "\U0001f6ab Unidades Inválidas",
    ])

    with tab_a0:
        st.subheader("Auditoria Ordinária — Água e Gás")

        df_cob = d.get("cobertura_agua_gas")
        if df_cob is not None and not df_cob.empty:
            # Semáforo
            for _, row in df_cob.iterrows():
                cat = row["categoria"]
                comp = row["competencia"]
                arrec = row["valor_arrecadado"]
                despesa = row["valor_despesa"]
                diff = arrec - despesa
                pct = row["cobertura_arrecadada_pct"]

                if despesa == 0:
                    cor = "green"
                    status_text = "Sem despesa"
                elif arrec >= despesa:
                    cor = "green"
                    status_text = "✅ Coberto"
                elif arrec >= despesa * 0.95:
                    cor = "orange"
                    status_text = "⚠️ Déficit < 5%"
                else:
                    cor = "red"
                    status_text = "🔴 Déficit > 5%"

                st.markdown(
                    f"""
                    <div style="border-left: 6px solid {cor}; padding: 12px; margin: 8px 0; background:#f9f9f9; border-radius:4px;">
                        <strong style="font-size:1.1rem;">{cat}</strong> — {comp}<br>
                        Arrecadado: <strong>R$ {arrec:,.2f}</strong> | 
                        Fatura: <strong>R$ {despesa:,.2f}</strong> | 
                        Diferença: <strong style="color:{cor};">R$ {diff:,.2f}</strong><br>
                        Cobertura: <strong>{pct:.1f}%</strong> — {status_text}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("---")
            st.subheader("Tabela Mensal de Cobertura")
            st.dataframe(
                df_cob,
                use_container_width=True, hide_index=True,
                column_config={
                    "valor_cobrado": st.column_config.NumberColumn("Valor Cobrado", format="R$ %.2f"),
                    "valor_arrecadado": st.column_config.NumberColumn("Arrecadado", format="R$ %.2f"),
                    "valor_despesa": st.column_config.NumberColumn("Fatura", format="R$ %.2f"),
                    "valor_em_aberto": st.column_config.NumberColumn("Em Aberto", format="R$ %.2f"),
                    "cobertura_arrecadada_pct": st.column_config.NumberColumn("Cobertura %", format="%.1f%%"),
                },
            )

            csv_cob = df_cob.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("\U0001f4e5 Exportar CSV (cobertura_agua_gas.csv)", data=csv_cob,
                               file_name="cobertura_agua_gas.csv", mime="text/csv")

            # Detalhe por unidade
            df_cob_unid = d.get("cobertura_agua_gas_detalhe")
            if df_cob_unid is not None and not df_cob_unid.empty:
                st.markdown("---")
                st.subheader("Detalhamento por Unidade")
                df_cob_unid_f = df_cob_unid.copy()
                if comp_sel != "Todas":
                    df_cob_unid_f = df_cob_unid_f[df_cob_unid_f["competencia"] == comp_sel]

                inad_unids = df_cob_unid_f[df_cob_unid_f["indicio_inadimplencia"] == "sim"]
                if not inad_unids.empty:
                    st.warning(f"{len(inad_unids)} unidades com indício de inadimplência em água/gás")

                st.dataframe(
                    df_cob_unid_f,
                    use_container_width=True, hide_index=True,
                    column_config={
                        "valor_cobrado": st.column_config.NumberColumn("Cobrado", format="R$ %.2f"),
                        "valor_pago": st.column_config.NumberColumn("Pago", format="R$ %.2f"),
                        "valor_em_atraso": st.column_config.NumberColumn("Em Atraso", format="R$ %.2f"),
                        "valor_em_acordo": st.column_config.NumberColumn("Em Acordo", format="R$ %.2f"),
                    },
                )
        else:
            st.info("Dados de cobertura água/gás não disponíveis. Processe os dados primeiro.")

    with tab_a1:
        df_nm = d["contas_nao_mapeadas"]
        if df_nm.empty:
            st.success("Todas as contas foram mapeadas com sucesso!")
        else:
            st.warning(f"{len(df_nm)} contas sem mapeamento de categoria")
            st.dataframe(df_nm, use_container_width=True, hide_index=True)

    with tab_a2:
        df_ni = d["nao_interpretadas"]
        if df_ni.empty:
            st.success("Nenhuma linha não interpretada.")
        else:
            st.warning(f"{len(df_ni)} linhas não interpretadas")
            st.dataframe(df_ni, use_container_width=True, hide_index=True)
            csv_ni = df_ni.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("\U0001f4e5 Exportar CSV (linhas_nao_interpretadas.csv)", data=csv_ni,
                               file_name="linhas_nao_interpretadas.csv", mime="text/csv")

    with tab_a3:
        st.markdown("Buscar no histórico bruto dos lançamentos")
        df_lanc = d["lancamentos"]
        if not df_lanc.empty:
            busca = st.text_input("Termo de busca", placeholder="Ex: RC.UN.000082 ou 51859790")
            if busca:
                mask = df_lanc["historico_bruto"].str.contains(busca, case=False, na=False)
                resultado = df_lanc[mask]
                if not resultado.empty:
                    st.success(f"{len(resultado)} ocorrências encontradas")
                    st.dataframe(resultado, use_container_width=True, hide_index=True)
                else:
                    st.info("Nenhum resultado encontrado.")
        else:
            st.info("Nenhum lançamento disponível.")

    with tab_a4:
        if st.session_state.log_renomeacao is not None and not st.session_state.log_renomeacao.empty:
            st.dataframe(st.session_state.log_renomeacao, use_container_width=True, hide_index=True)
            csv_log = st.session_state.log_renomeacao.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("\U0001f4e5 Exportar CSV (log_renomeacao.csv)", data=csv_log,
                               file_name="log_renomeacao.csv", mime="text/csv")
        else:
            st.info("Log de renomeação disponível apenas no modo Lote.")

    with tab_a5:
        st.subheader("Divergências Pendentes")

        # Totalização
        df_totalizacao = _validar_totalizacao(d["resumo_mensal"])
        if not df_totalizacao.empty:
            st.markdown("**Validação de Totalização**")
            for _, r in df_totalizacao.iterrows():
                if "⚠️" in str(r["status"]):
                    st.error(f"{r['competencia']}: {r['status']}")

        # Divergências da conciliação
        df_conc = d["conciliacao"]
        if not df_conc.empty:
            divs = df_conc[df_conc["status_conciliacao"] == "divergente"]
            if not divs.empty:
                st.markdown("---")
                st.error(f"**{len(divs)} divergências na conciliação**")
                st.dataframe(divs, use_container_width=True, hide_index=True)

        # Caixa local sem match
        df_caixa_sem = d.get("caixa_sem_match")
        if df_caixa_sem is not None and not df_caixa_sem.empty:
            st.markdown("---")
            st.warning(f"**{len(df_caixa_sem)} itens de caixa local sem match no extrato**")
            st.dataframe(df_caixa_sem, use_container_width=True, hide_index=True)

        # Contas não mapeadas
        df_nm = d["contas_nao_mapeadas"]
        if not df_nm.empty:
            st.markdown("---")
            st.warning(f"**{len(df_nm)} contas sem mapeamento**")
            st.dataframe(df_nm, use_container_width=True, hide_index=True)

        # Lançamentos não interpretados
        df_ni = d["nao_interpretadas"]
        if not df_ni.empty:
            st.markdown("---")
            st.warning(f"**{len(df_ni)} linhas não interpretadas**")
            st.dataframe(df_ni, use_container_width=True, hide_index=True)

        if (df_totalizacao.empty or all("✅" in str(s) for s in df_totalizacao["status"])) and \
           (df_conc.empty or divs.empty) and \
           (df_caixa_sem is None or df_caixa_sem.empty) and \
           df_nm.empty and df_ni.empty:
            st.success("Nenhuma divergência pendente encontrada.")

    with tab_a6:
        st.subheader("Auditoria de Unidades Inválidas")
        df_aud_unid = d.get("unidades_auditoria")
        if df_aud_unid is not None and not df_aud_unid.empty:
            st.warning(f"{len(df_aud_unid)} ocorrências de unidades inválidas detectadas")
            st.dataframe(df_aud_unid, use_container_width=True, hide_index=True)
            csv_aud_unid = df_aud_unid.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("\U0001f4e5 Exportar CSV (unidades_invalidas_auditoria.csv)", data=csv_aud_unid,
                               file_name="unidades_invalidas_auditoria.csv", mime="text/csv")

            # ── Quick validation from auditoria ──
            st.markdown("---")
            st.subheader("Validar unidade inválida")
            from src.services.manual_validation import (
                salvar_validacao, gerar_id_registro,
            )

            sel_idx = st.selectbox(
                "Selecionar registro para validar",
                range(len(df_aud_unid)),
                format_func=lambda i: (
                    f"{df_aud_unid.iloc[i].get('competencia','')} | "
                    f"Unid detectada: {df_aud_unid.iloc[i].get('numero_detectado','')} | "
                    f"Valor: R$ {df_aud_unid.iloc[i].get('valor',0):.2f}"
                ),
                key="aud_unid_sel",
            )
            row_aud = df_aud_unid.iloc[sel_idx]
            reg_id = gerar_id_registro(
                row_aud.get("competencia", ""),
                row_aud.get("numero_detectado", ""),
                row_aud.get("codigo_conta", ""),
                row_aud.get("data", ""),
            )
            st.text(f"ID: {reg_id}")

            col_au1, col_au2 = st.columns(2)
            with col_au1:
                unid_corrigida = st.text_input(
                    "Unidade correta",
                    value="",
                    placeholder="Ex: 162",
                    key="aud_unid_corr",
                )
            with col_au2:
                justif_aud = st.text_area("Justificativa", key="aud_justif", height=68)

            if st.button("💾 Salvar correção de unidade", use_container_width=True):
                if unid_corrigida:
                    registro = {
                        "id_registro": reg_id,
                        "competencia_original": str(row_aud.get("competencia", "")),
                        "numero_detectado": str(row_aud.get("numero_detectado", "")),
                        "unidade_corrigida": unid_corrigida,
                        "historico_bruto": str(row_aud.get("historico_bruto", "")),
                        "codigo_conta": str(row_aud.get("codigo_conta", "")),
                        "descricao_conta": str(row_aud.get("descricao_conta", "")),
                        "categoria_original": str(row_aud.get("categoria", "")),
                        "categoria_corrigida": "",
                        "competencia_corrigida": "",
                        "valor": str(row_aud.get("valor", 0)),
                        "status_suspeita": str(row_aud.get("status_suspeita", "pendente")),
                        "acao_manual": "corrigir_unidade",
                        "incluir_no_calculo": "sim",
                        "justificativa": justif_aud,
                    }
                    if salvar_validacao(registro):
                        st.success("Correção de unidade salva! Recalcule a base para aplicar.")
                        st.rerun()
                else:
                    st.error("Informe a unidade correta.")
        else:
            st.success("Nenhuma unidade inválida detectada.")

        # Mostrar lista de unidades oficiais
        st.markdown("---")
        st.subheader("Unidades Oficiais do Condomínio")
        unids_str = ", ".join(str(u) for u in sorted(UNIDADES_VALIDAS))
        st.markdown(f"**{len(UNIDADES_VALIDAS)} unidades cadastradas:** {unids_str}")

    st.markdown("---")
    st.subheader("Lançamentos Completos")
    df_lanc_f = filtrar(d["lancamentos"])
    if unidade_sel != "Todas":
        df_lanc_f = df_lanc_f[df_lanc_f["unidade"] == unidade_sel]
    if cat_sel != "Todas":
        df_lanc_f = df_lanc_f[df_lanc_f["categoria"] == cat_sel]
    if not df_lanc_f.empty:
        st.dataframe(df_lanc_f, use_container_width=True, hide_index=True)
        total_cred = df_lanc_f[df_lanc_f["tipo"] == "credito"]["valor"].sum()
        total_deb = df_lanc_f[df_lanc_f["tipo"] == "debito"]["valor"].sum()
        st.markdown(f"**Créditos:** {fmt_br(total_cred)}  |  **Débitos:** {fmt_br(total_deb)}")

# ═══════════════════════════════════════════════════════════════
# TAB 7 - CAIXA LOCAL (Reposição de Caixa Local)
# ═══════════════════════════════════════════════════════════════
with tab_caixa:
    st.subheader("Reposição de Caixa Local - Despesas Locais")

    df_caixa = d.get("caixa_local")
    df_resumo = d.get("resumo_caixa")
    df_sem_match = d.get("caixa_sem_match")

    if df_caixa is None or df_caixa.empty:
        st.warning("Nenhum dado de caixa local disponível. Use o modo Lote com PDFs de despesas locais em data/renamed/.")
    else:
        tab_caixa_resumo, tab_caixa_detalhe, tab_caixa_audit, tab_caixa_timeline = st.tabs([
            "\U0001f4ca Resumo", "\U0001f4cb Despesas Locais",
            "\U0001f50d Auditoria", "\U0001f4c8 Evolução",
        ])

        with tab_caixa_resumo:
            if df_resumo is not None and not df_resumo.empty:
                st.subheader("Resumo por Competência")
                st.dataframe(
                    df_resumo,
                    use_container_width=True, hide_index=True,
                    column_config={
                        "total_gastos_pdf": st.column_config.NumberColumn("Gastos PDF", format="R$ %.2f"),
                        "total_extrato": st.column_config.NumberColumn("Total Extrato", format="R$ %.2f"),
                        "diferenca": st.column_config.NumberColumn("Diferença", format="R$ %.2f"),
                    },
                )

                st.markdown("---")
                col_r1, col_r2, col_r3, col_r4 = st.columns(4)
                with col_r1:
                    st.markdown(kpi_card("Total Gastos PDF", fmt_br(df_resumo["total_gastos_pdf"].sum())), unsafe_allow_html=True)
                with col_r2:
                    st.markdown(kpi_card("Total Extrato", fmt_br(df_resumo["total_extrato"].sum()), color="green"), unsafe_allow_html=True)
                with col_r3:
                    st.markdown(kpi_card("Itens Conciliados", str(df_resumo["itens_conciliados"].sum())), unsafe_allow_html=True)
                with col_r4:
                    nao_conc = df_resumo["itens_nao_conciliados"].sum()
                    st.markdown(kpi_card("Não Conciliados", str(nao_conc), color="red" if nao_conc > 0 else ""), unsafe_allow_html=True)

                fig_resumo = px.bar(
                    df_resumo, x="competencia",
                    y=["total_gastos_pdf", "total_extrato"],
                    barmode="group",
                    title="Caixa Local: Gastos PDF vs Extrato",
                    labels={"value": "R$", "variable": "", "competencia": "Competência"},
                    color_discrete_map={"total_gastos_pdf": "#d62728", "total_extrato": "#2ca02c"},
                )
                st.plotly_chart(fig_resumo, use_container_width=True)

        with tab_caixa_detalhe:
            st.subheader("Despesas Locais - Detalhamento")
            comps_caixa = ["Todas"] + sorted(df_caixa["competencia"].unique())
            comp_caixa_sel = st.selectbox("Competência", comps_caixa, key="comp_caixa")
            df_caixa_f = df_caixa.copy()
            if comp_caixa_sel != "Todas":
                df_caixa_f = df_caixa_f[df_caixa_f["competencia"] == comp_caixa_sel]

            st.dataframe(
                df_caixa_f,
                use_container_width=True, hide_index=True,
                column_config={
                    "valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                    "descricao_extrato_original": st.column_config.Column("Descrição Original"),
                    "descricao_final_conciliada": st.column_config.Column("Descrição Final"),
                    "status_match_pdf": st.column_config.Column("Status Match"),
                },
            )

            csv_caixa = df_caixa.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button("\U0001f4e5 Exportar CSV (caixa_local_conciliado.csv)", data=csv_caixa,
                               file_name="caixa_local_conciliado.csv", mime="text/csv")

        with tab_caixa_audit:
            st.subheader("Auditoria de Caixa Local")

            if df_sem_match is not None and not df_sem_match.empty:
                st.warning(f"{len(df_sem_match)} itens sem match no extrato")
                st.dataframe(df_sem_match, use_container_width=True, hide_index=True)

                csv_sem = df_sem_match.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button("\U0001f4e5 Exportar CSV (caixa_local_sem_match.csv)", data=csv_sem,
                                   file_name="caixa_local_sem_match.csv", mime="text/csv")
            else:
                st.success("Todos os itens do PDF foram conciliados com o extrato!")

            st.markdown("---")
            st.subheader("Itens do Extrato (1723) sem detalhe no PDF")
            if not df_caixa.empty:
                sem_pdf = df_caixa[df_caixa["status_match_pdf"] == "NAO_ENCONTRADO"]
                sem_pdf_extrato = sem_pdf[sem_pdf["descricao_extrato_original"] != ""]
                if not sem_pdf_extrato.empty:
                    st.warning(f"{len(sem_pdf_extrato)} itens no extrato sem correspondência no PDF")
                    st.dataframe(sem_pdf_extrato, use_container_width=True, hide_index=True)

        with tab_caixa_timeline:
            st.subheader("Evolução Mensal dos Gastos de Caixa Local")

            if df_caixa is not None and not df_caixa.empty:
                fig_evol = px.bar(
                    df_caixa,
                    x="competencia", y="valor",
                    color="subcategoria",
                    title="Gastos por Competência e Subcategoria",
                    labels={"valor": "R$", "competencia": "Competência"},
                    color_discrete_map={
                        "portaria": "#1f77b4", "piscina": "#2ca02c", "exaustao": "#ff7f0e",
                        "copa": "#9467bd", "limpeza": "#8c564b", "jardinagem": "#e377c2",
                        "lazer": "#7f7f7f", "manutencao": "#bcbd22", "materiais": "#17becf",
                        "despensa": "#aec7e8", "outros": "#d62728",
                    },
                )
                fig_evol.update_layout(xaxis_title="Competência", yaxis_title="R$")
                st.plotly_chart(fig_evol, use_container_width=True)

                st.markdown("---")
                subcat_count = df_caixa["subcategoria"].value_counts().reset_index()
                subcat_count.columns = ["subcategoria", "quantidade"]
                fig_sub = px.pie(subcat_count, values="quantidade", names="subcategoria",
                                 title="Distribuição por Subcategoria",
                                 color="subcategoria",
                                 color_discrete_map={
                                     "portaria": "#1f77b4", "piscina": "#2ca02c", "exaustao": "#ff7f0e",
                                     "copa": "#9467bd", "limpeza": "#8c564b", "jardinagem": "#e377c2",
                                     "lazer": "#7f7f7f", "manutencao": "#bcbd22", "materiais": "#17becf",
                                     "despensa": "#aec7e8", "outros": "#d62728",
                                 })
                st.plotly_chart(fig_sub, use_container_width=True)

                if "fornecedor_sugerido" in df_caixa.columns:
                    fornecs = df_caixa[df_caixa["fornecedor_sugerido"].notna() & (df_caixa["fornecedor_sugerido"] != "")]
                    if not fornecs.empty:
                        st.markdown("---")
                        st.subheader("Fornecedores Identificados")
                        top_forn = fornecs["fornecedor_sugerido"].value_counts().head(10).reset_index()
                        top_forn.columns = ["fornecedor", "ocorrencias"]
                        fig_forn = px.bar(top_forn, x="fornecedor", y="ocorrencias",
                                          title="Fornecedores mais Recorrentes",
                                          color="ocorrencias", color_continuous_scale="Viridis")
                        fig_forn.update_layout(xaxis_title="", showlegend=False)
                        st.plotly_chart(fig_forn, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 8 - ESPELHO DO BOLETO POR UNIDADE
# ═══════════════════════════════════════════════════════════════
with tab_espelho:
    st.subheader("Espelho do Boleto por Unidade")

    df_espelho = d.get("espelho_boleto")
    df_espelho_cons = d.get("espelho_boleto_consolidado")

    if df_espelho is None or df_espelho.empty:
        st.warning("Dados de espelho do boleto não disponíveis. Processe os dados primeiro.")
    else:
        # ── Audit attrs ──
        audit_data = df_espelho.attrs.get("auditoria", [])
        excluidos = df_espelho.attrs.get("itens_excluidos", [])
        duplicatas = df_espelho.attrs.get("duplicatas", [])
        discrepancias = df_espelho.attrs.get("discrepancias", [])
        confiabilidade = df_espelho.attrs.get("confiabilidade", "alta")

        # ── Reliability / Warning banners ──
        conf_colors = {"alta": "#0a0", "media": "#e80", "baixa": "#d00"}
        conf_icon = {"alta": "✅", "media": "⚠️", "baixa": "🔴"}
        st.markdown(
            f"""<div style="border:1px solid {conf_colors[confiabilidade]}; border-radius:6px;
                        padding:8px 16px; margin-bottom:12px; background:#f9f9f9;">
                <strong>{conf_icon[confiabilidade]} Confiabilidade: {confiabilidade.upper()}</strong>
            </div>""",
            unsafe_allow_html=True,
        )

        if discrepancias:
            st.warning(f"Valores atípicos detectados: {len(discrepancias)} item(ns) fogem do padrão histórico")

        # ── Selectors ──
        unids_data = sorted(u for u in df_espelho["unidade"].unique() if u.isdigit() and int(u) in UNIDADES_VALIDAS)
        if not unids_data:
            unids_data = sorted(UNIDADES_VALIDAS)
        col_sel1, col_sel2, _ = st.columns([2, 2, 6])
        with col_sel1:
            unid_sel = st.selectbox("Unidade", [str(u) for u in unids_data], key="espelho_unidade")
        with col_sel2:
            comps_ref = ["Todas"] + sorted(df_espelho["competencia_referencia"].unique(), reverse=True)
            comp_ref_sel = st.selectbox("Competência Referência", comps_ref, key="espelho_comp_ref")

        # ── Sub-tabs ──
        sub_bol, sub_aud, sub_exc, sub_dup, sub_dis = st.tabs([
            "📄 Itens do Boleto", "🔍 Auditoria", "🚫 Itens Excluídos",
            "🔁 Duplicatas", "📊 Discrepâncias",
        ])

        with sub_bol:
            df_esp_f = df_espelho[df_espelho["unidade"] == unid_sel].copy()
            if comp_ref_sel != "Todas":
                df_esp_f = df_esp_f[df_esp_f["competencia_referencia"] == comp_ref_sel]

            if df_esp_f.empty:
                st.info("Nenhum lançamento encontrado para esta unidade no período selecionado.")
            else:
                comps_disponiveis = sorted(df_esp_f["competencia_referencia"].unique(), reverse=True)
                for comp_ref in comps_disponiveis:
                    df_comp = df_esp_f[df_esp_f["competencia_referencia"] == comp_ref]
                    df_comp = ensure_columns(
                        df_comp,
                        ["codigo", "descricao_da_verba", "ref", "valor_esperado",
                         "valor_pago", "categoria", "observacoes", "status_item"],
                        stage="espelho_itens_boleto",
                    )

                    total_esp = df_comp["valor_esperado"].sum()
                    total_pago = df_comp["valor_pago"].sum()
                    diferenca = round(total_pago - total_esp, 2)
                    status_comp = df_comp["status_competencia"].iloc[0] if not df_comp.empty else ""

                    data_pag = df_comp["data_pagamento"].iloc[0] if not df_comp.empty and df_comp["data_pagamento"].iloc[0] else "—"
                    mes_label = {
                        "01": "JANEIRO", "02": "FEVEREIRO", "03": "MARÇO", "04": "ABRIL",
                        "05": "MAIO", "06": "JUNHO", "07": "JULHO", "08": "AGOSTO",
                        "09": "SETEMBRO", "10": "OUTUBRO", "11": "NOVEMBRO", "12": "DEZEMBRO",
                    }
                    try:
                        mes_num = comp_ref.split("-")[1]
                        mes_nome = mes_label.get(mes_num, comp_ref)
                        ano_num = comp_ref.split("-")[0]
                        titulo = f"BOLETO DE {mes_nome} / {ano_num}"
                    except (IndexError, AttributeError):
                        titulo = f"BOLETO {comp_ref}"

                    status_cor = {
                        "pago_no_mes": "#0a0", "pago_antecipado": "#e80",
                        "pago_com_divergencia": "#e80", "parcial": "#d00",
                        "em_acordo": "#0a0", "inadimplente_presumido": "#d00",
                        "nao_localizado": "#d00", "sem_lancamentos": "#888",
                    }.get(status_comp, "#888")

                    st.markdown(
                        f"""
                        <div style="border:2px solid #ddd; border-radius:8px; padding:16px; margin:16px 0; background:#fafafa;">
                            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:2px solid #333; padding-bottom:8px; margin-bottom:12px;">
                                <span style="font-size:1.3rem; font-weight:700;">{titulo}</span>
                                <span style="font-size:0.9rem; color:#555;">Unidade {unid_sel}</span>
                            </div>
                            <div style="display:flex; gap:24px; margin-bottom:12px; font-size:0.85rem; color:#555;">
                                <span>Pagamento: <strong>{data_pag}</strong></span>
                                <span>Status: <strong style="color:{status_cor};">{status_comp.replace('_', ' ').upper()}</strong></span>
                            </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    st.dataframe(
                        df_comp[["codigo", "descricao_da_verba", "ref", "valor_esperado",
                                 "valor_pago", "categoria", "observacoes", "status_item"]],
                        use_container_width=True, hide_index=True,
                        column_config={
                            "valor_esperado": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f"),
                            "valor_pago": st.column_config.NumberColumn("Pago (R$)", format="R$ %.2f"),
                            "codigo": st.column_config.Column("Código", width="small"),
                            "descricao_da_verba": st.column_config.Column("Descrição"),
                            "ref": st.column_config.Column("Ref", width="small"),
                            "categoria": st.column_config.Column("Categoria", width="small"),
                            "observacoes": st.column_config.Column("Observações"),
                            "status_item": st.column_config.Column("Status"),
                        },
                    )

                    col_t1, col_t2, col_t3 = st.columns(3)
                    with col_t1:
                        st.markdown(f"**Total do Boleto:** R$ {total_esp:,.2f}")
                    with col_t2:
                        st.markdown(f"**Total Pago:** R$ {total_pago:,.2f}")
                    with col_t3:
                        cor_diff = "green" if abs(diferenca) < 0.01 else ("red" if diferenca < 0 else "orange")
                        st.markdown(f"**Diferença:** <span style='color:{cor_diff};'>R$ {diferenca:,.2f}</span>", unsafe_allow_html=True)

                    st.markdown("</div>", unsafe_allow_html=True)

                if len(comps_disponiveis) > 1:
                    st.markdown("---")
                    st.subheader("Comparativo entre Meses")
                    df_comp_meses = df_esp_f.groupby("competencia_referencia").agg(
                        total_esperado=("valor_esperado", "sum"),
                        total_pago=("valor_pago", "sum"),
                        status_competencia=("status_competencia", "first"),
                    ).reset_index()
                    df_comp_meses["diferenca"] = (df_comp_meses["total_pago"] - df_comp_meses["total_esperado"]).round(2)
                    st.dataframe(
                        df_comp_meses.sort_values("competencia_referencia"),
                        use_container_width=True, hide_index=True,
                        column_config={
                            "total_esperado": st.column_config.NumberColumn("Esperado", format="R$ %.2f"),
                            "total_pago": st.column_config.NumberColumn("Pago", format="R$ %.2f"),
                            "diferenca": st.column_config.NumberColumn("Diferença", format="R$ %.2f"),
                        },
                    )

            # ── Consolidated view ──
            st.markdown("---")
            st.subheader("Resumo por Unidade e Competência")
            if df_espelho_cons is not None and not df_espelho_cons.empty:
                df_cons_f = df_espelho_cons.copy()
                if unid_sel and unid_sel != "Todas":
                    df_cons_f = df_cons_f[df_cons_f["unidade"] == unid_sel]

                st.dataframe(
                    df_cons_f,
                    use_container_width=True, hide_index=True,
                    column_config={
                        "total_esperado": st.column_config.NumberColumn("Esperado", format="R$ %.2f"),
                        "total_pago": st.column_config.NumberColumn("Pago", format="R$ %.2f"),
                        "diferenca": st.column_config.NumberColumn("Diferença", format="R$ %.2f"),
                        "data_pagamento": st.column_config.Column("Data Pagamento"),
                        "status_competencia": st.column_config.Column("Status"),
                    },
                )

        with sub_aud:
            st.subheader("Auditoria do Espelho por Unidade")
            if not audit_data:
                st.success("Nenhum item suspeito na auditoria.")
            else:
                df_aud = pd.DataFrame(audit_data)
                # Filter by unit/comp
                if unid_sel and unid_sel != "Todas":
                    df_aud = df_aud[df_aud["unidade"] == unid_sel]
                if comp_ref_sel != "Todas":
                    df_aud = df_aud[df_aud["competencia_referencia"] == comp_ref_sel]

                st.warning(f"{len(df_aud)} registros de auditoria")
                st.dataframe(df_aud, use_container_width=True, hide_index=True)

                csv_aud = df_aud.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "\U0001f4e5 Exportar Auditoria (csv)",
                    data=csv_aud, file_name="auditoria_espelho_unidade.csv",
                    mime="text/csv", use_container_width=True,
                )

        with sub_exc:
            st.subheader("Itens Excluídos do Espelho")
            st.caption("Categorias que não fazem parte do boleto da unidade")
            if not excluidos:
                st.success("Nenhum item foi excluído.")
            else:
                df_exc = pd.DataFrame(excluidos)
                df_exc = ensure_columns(
                    df_exc,
                    ["competencia", "unidade", "categoria", "valor", "codigo_conta",
                     "descricao_conta", "historico_bruto"],
                    stage="espelho_excluidos",
                )
                # Filter by unit
                if unid_sel and unid_sel != "Todas":
                    df_exc = df_exc[df_exc["unidade"] == unid_sel]
                if comp_ref_sel != "Todas" and "competencia" in df_exc.columns:
                    df_exc = df_exc[df_exc["competencia"] == comp_ref_sel]

                st.info(f"{len(df_exc)} itens excluídos por não serem verbas de boleto")
                st.dataframe(
                    df_exc.reindex(columns=["competencia", "unidade", "categoria", "valor",
                                            "codigo_conta", "descricao_conta", "historico_bruto"]),
                    use_container_width=True, hide_index=True,
                    column_config={
                        "valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                    },
                )

                csv_exc = df_exc.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "\U0001f4e5 Exportar Itens Excluídos (csv)",
                    data=csv_exc, file_name="itens_excluidos_espelho.csv",
                    mime="text/csv", use_container_width=True,
                )

        with sub_dup:
            st.subheader("Itens Duplicados Detectados")
            if not duplicatas:
                st.success("Nenhuma duplicata encontrada.")
            else:
                df_dup = pd.DataFrame(duplicatas)
                df_dup = ensure_columns(
                    df_dup,
                    ["unidade", "competencia_referencia", "codigo", "valor",
                     "data_pagamento", "arquivo_origem"],
                    stage="espelho_duplicatas",
                )
                if unid_sel and unid_sel != "Todas":
                    df_dup = df_dup[df_dup["unidade"] == unid_sel]
                if comp_ref_sel != "Todas":
                    df_dup = df_dup[df_dup["competencia_referencia"] == comp_ref_sel]

                st.warning(f"{len(df_dup)} duplicatas detectadas (foram excluídas do cálculo)")
                st.dataframe(
                    df_dup.reindex(columns=["unidade", "competencia_referencia", "codigo", "valor",
                                            "data_pagamento", "arquivo_origem"]),
                    use_container_width=True, hide_index=True,
                    column_config={
                        "valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                    },
                )

                csv_dup = df_dup.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "\U0001f4e5 Exportar Duplicatas (csv)",
                    data=csv_dup, file_name="itens_duplicados_unidade.csv",
                    mime="text/csv", use_container_width=True,
                )

        with sub_dis:
            st.subheader("Valores Discrepantes")
            st.caption("Itens cujo valor foge mais de 50% da mediana histórica da mesma unidade/verba")
            if not discrepancias:
                st.success("Nenhum valor discrepante detectado.")
            else:
                df_dis = pd.DataFrame(discrepancias)
                if unid_sel and unid_sel != "Todas":
                    df_dis = df_dis[df_dis["unidade"] == unid_sel]

                for _, r in df_dis.iterrows():
                    cor = "#d00" if r["classificacao"] == "VALOR_DISCREPANTE" else "#e80"
                    st.markdown(
                        f"""<div style="border-left:6px solid {cor}; padding:8px 12px; margin:4px 0; background:#f9f9f9;">
                            <strong>Un {r['unidade']}</strong> — {r['competencia_referencia']}
                            — {r['descricao']}<br>
                            Atual: <strong>R$ {r['valor_atual']:,.2f}</strong> |
                            Mediana: R$ {r['mediana_historica']:,.2f} |
                            Razão: {r['razao']:.2f}x |
                            <span style="color:{cor};">{r['classificacao']}</span>
                        </div>""",
                        unsafe_allow_html=True,
                    )

                if not df_dis.empty:
                    csv_dis = df_dis.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.download_button(
                        "\U0001f4e5 Exportar Discrepâncias (csv)",
                        data=csv_dis, file_name="valores_discrepantes_unidade.csv",
                        mime="text/csv", use_container_width=True,
                    )

        # ── Main exports ──
        st.markdown("---")
        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            csv_esp = df_espelho.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "\U0001f4e5 Exportar Espelho Detalhado (csv)",
                data=csv_esp,
                file_name="espelho_boleto_unidade_detalhado.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with col_exp2:
            if df_espelho_cons is not None and not df_espelho_cons.empty:
                csv_cons = df_espelho_cons.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "\U0001f4e5 Exportar Resumo Consolidado (csv)",
                    data=csv_cons,
                    file_name="espelho_boleto_unidade.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

    # ── Debug ──
    with st.expander("🔧 Debug - Espelho do Boleto", expanded=False):
        st.markdown("**DataFrames:**")
        st.write(f"- `df_espelho`: shape={df_espelho.shape if df_espelho is not None else 'None'}, cols={list(df_espelho.columns) if df_espelho is not None else 'N/A'}")
        st.write(f"- `df_espelho_cons`: shape={df_espelho_cons.shape if df_espelho_cons is not None else 'None'}")
        st.write(f"- audit_data: {len(audit_data)} itens")
        st.write(f"- excluidos: {len(excluidos)} itens")
        st.write(f"- duplicatas: {len(duplicatas)} itens")
        st.write(f"- discrepancias: {len(discrepancias)} itens")
        st.write(f"**Seleção:** unidade={unid_sel}, competência={comp_ref_sel}")
        if df_espelho is not None and not df_espelho.empty:
            st.write("**Primeiras linhas:**")
            st.dataframe(df_espelho.head(5), use_container_width=True)
        if excluidos:
            df_exc_debug = pd.DataFrame(excluidos)
            st.write(f"**Colunas df_excluidos:** {list(df_exc_debug.columns)}")
            st.dataframe(df_exc_debug.head(3), use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 9 - EXTRATO CONSOLIDADO
# ═══════════════════════════════════════════════════════════════
with tab_extrato_cons:
    st.subheader("Extrato Consolidado - Lançamentos Normalizados")
    st.caption(
        "Todos os lançamentos extraídos dos extratos bancários, "
        "com colunas de validação manual aplicáveis."
    )

    if st.session_state.dados is None:
        st.info("Nenhum dado carregado. Importe ou carregue a base histórica primeiro.")
    else:
        df_ext = st.session_state.dados["lancamentos"].copy()
        if "data_lancamento" in df_ext.columns:
            df_ext["data_lancamento"] = df_ext["data_lancamento"].astype(str)

        # ── Ensure validation columns exist ──
        for col in ["status_validacao", "origem_ajuste_manual", "incluir_no_calculo", "chave_deduplicacao"]:
            if col not in df_ext.columns:
                df_ext[col] = "" if col in ("status_validacao", "origem_ajuste_manual", "chave_deduplicacao") else "sim"

        total = len(df_ext)
        incluidas = (df_ext["incluir_no_calculo"] == "sim").sum()
        excluidas = total - incluidas

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Total Lançamentos", total)
        col_m2.metric("Incluídos no Cálculo", incluidas)
        col_m3.metric("Excluídos", excluidas)
        if "status_validacao" in df_ext.columns:
            pendentes = (df_ext["status_validacao"] == "pendente").sum()
            col_m4.metric("Pendentes Validação", pendentes)

        # ── Filters ──
        col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
        comps = ["Todas"] + sorted(df_ext["competencia"].dropna().unique()) if "competencia" in df_ext.columns else ["Todas"]
        unids = ["Todas"] + sorted(df_ext["unidade"].dropna().unique()) if "unidade" in df_ext.columns else ["Todas"]
        cats = ["Todas"] + sorted(df_ext["categoria"].dropna().unique()) if "categoria" in df_ext.columns else ["Todas"]
        status_opts = ["Todas", "pendente", "corrigido", "excluido", "mantido"]

        with col_f1:
            f_comp = st.selectbox("Competência", comps, key="ext_comp")
        with col_f2:
            f_unid = st.selectbox("Unidade", unids, key="ext_unid")
        with col_f3:
            f_cat = st.selectbox("Categoria", cats, key="ext_cat")
        with col_f4:
            f_status = st.selectbox("Status Validação", status_opts, key="ext_status")
        with col_f5:
            f_inc = st.selectbox("Incluir no cálculo", ["Todos", "sim", "nao"], key="ext_inc")

        # ── Apply filters ──
        if f_comp != "Todas":
            df_ext = df_ext[df_ext["competencia"] == f_comp]
        if f_unid != "Todas":
            df_ext = df_ext[df_ext["unidade"] == f_unid]
        if f_cat != "Todas":
            df_ext = df_ext[df_ext["categoria"] == f_cat]
        if f_status != "Todas":
            df_ext = df_ext[df_ext["status_validacao"] == f_status]
        if f_inc != "Todos":
            df_ext = df_ext[df_ext["incluir_no_calculo"] == f_inc]

        # ── Display ──
        display_cols = [c for c in [
            "unidade", "competencia", "codigo_conta", "descricao_conta",
            "categoria", "valor", "data_lancamento",
            "status_validacao", "incluir_no_calculo", "origem_ajuste_manual",
        ] if c in df_ext.columns]

        st.dataframe(
            df_ext[display_cols].sort_values(["unidade", "competencia"]),
            use_container_width=True, hide_index=True,
        )

        csv_ext = df_ext.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "\U0001f4e5 Exportar Extrato Consolidado (csv)",
            data=csv_ext,
            file_name="extrato_consolidado_ajustado.csv",
            mime="text/csv",
            use_container_width=True,
        )

# ═══════════════════════════════════════════════════════════════
# TAB 10 - VALIDAÇÃO MANUAL
# ═══════════════════════════════════════════════════════════════
with tab_validacao:
    st.subheader("Validação Manual de Lançamentos")
    st.caption(
        "Registre correções manuais para lançamentos suspeitos ou "
        "incorretamente identificados. As alterações serão aplicadas "
        "ao recalcular a base."
    )

    from src.services.manual_validation import (
        carregar_validacoes, salvar_validacao, remover_validacao,
        gerar_id_registro, ACOES_DISPONIVEIS, COLUNAS_VALIDACAO,
    )

    validacoes_existentes = carregar_validacoes()

    col_v1, col_v2, col_v3 = st.columns(3)
    col_v1.metric("Total Validações", len(validacoes_existentes))
    if not validacoes_existentes.empty and "acao_manual" in validacoes_existentes.columns:
        col_v2.metric("Exclusões", (validacoes_existentes["acao_manual"].str.contains("excluir")).sum())
        col_v3.metric("Correções", (validacoes_existentes["acao_manual"].str.contains("corrigir")).sum())

    # ── Section: Registrar nova validação ──
    with st.expander("➕ Registrar nova validação", expanded=False):
        if st.session_state.dados is not None:
            df_base = st.session_state.dados["lancamentos"].copy()
            if "data_lancamento" in df_base.columns:
                df_base["data_lancamento"] = df_base["data_lancamento"].astype(str)

            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                bus_comp = st.text_input("Competência (ex: 2026-04)", key="val_bus_comp")
            with col_s2:
                bus_unid = st.text_input("Unidade (ex: 162)", key="val_bus_unid")
            with col_s3:
                bus_cod = st.text_input("Código conta (ex: 1.01.001)", key="val_bus_cod")

            if st.button("🔍 Buscar registro", use_container_width=True):
                df_match = df_base.copy()
                if bus_comp:
                    df_match = df_match[df_match["competencia"].astype(str).str.contains(bus_comp)]
                if bus_unid:
                    df_match = df_match[df_match["unidade"].astype(str).str.contains(bus_unid)]
                if bus_cod:
                    df_match = df_match[df_match.get("codigo_conta", "").astype(str).str.contains(bus_cod)]

                if df_match.empty:
                    st.warning("Nenhum registro encontrado com esses critérios.")
                else:
                    st.dataframe(
                        df_match[[c for c in [
                            "unidade", "competencia", "codigo_conta", "descricao_conta",
                            "categoria", "valor", "data_lancamento",
                        ] if c in df_match.columns]].head(20),
                        use_container_width=True, hide_index=True,
                    )

                    # ── Selection and validation form ──
                    sel_opts = df_match.apply(
                        lambda r: f"{r.get('unidade','')} | {r.get('competencia','')} | "
                                  f"{r.get('codigo_conta','')} | R$ {r.get('valor',0):.2f}",
                        axis=1,
                    ).tolist()
                    sel_idx = st.selectbox(
                        "Selecionar registro para validar",
                        range(len(sel_opts)),
                        format_func=lambda i: sel_opts[i],
                        key="val_sel_reg",
                    )

                    row = df_match.iloc[sel_idx]
                    reg_id = gerar_id_registro(
                        row.get("competencia", ""),
                        row.get("unidade", ""),
                        row.get("codigo_conta", ""),
                        row.get("data_lancamento", ""),
                    )
                    st.text(f"ID do registro: {reg_id}")

                    col_a1, col_a2 = st.columns(2)
                    with col_a1:
                        acao = st.selectbox("Ação", ACOES_DISPONIVEIS, key="val_acao")
                    with col_a2:
                        incluir = st.selectbox("Incluir no cálculo", ["sim", "nao"], key="val_incluir")

                    col_c1, col_c2, col_c3 = st.columns(3)
                    with col_c1:
                        unid_corr = st.text_input(
                            "Unidade corrigida", value=str(row.get("unidade", "")),
                            key="val_unid_corr",
                        )
                    with col_c2:
                        comp_corr = st.text_input(
                            "Competência corrigida", value=str(row.get("competencia", "")),
                            key="val_comp_corr",
                        )
                    with col_c3:
                        cat_corr = st.text_input(
                            "Categoria corrigida", value=str(row.get("categoria", "")),
                            key="val_cat_corr",
                        )

                    justif = st.text_area("Justificativa", key="val_justif")

                    if st.button("💾 Salvar Validação", type="primary", use_container_width=True):
                        registro = {
                            "id_registro": reg_id,
                            "competencia_original": str(row.get("competencia", "")),
                            "numero_detectado": str(row.get("unidade", "")),
                            "unidade_corrigida": unid_corr if unid_corr != row.get("unidade", "") else "",
                            "historico_bruto": str(row.get("historico_bruto", row.get("descricao_conta", ""))),
                            "codigo_conta": str(row.get("codigo_conta", "")),
                            "descricao_conta": str(row.get("descricao_conta", "")),
                            "categoria_original": str(row.get("categoria", "")),
                            "categoria_corrigida": cat_corr if cat_corr != row.get("categoria", "") else "",
                            "competencia_corrigida": comp_corr if comp_corr != row.get("competencia", "") else "",
                            "valor": str(row.get("valor", 0)),
                            "status_suspeita": str(row.get("status_validacao", "pendente")),
                            "acao_manual": acao,
                            "incluir_no_calculo": incluir,
                            "justificativa": justif,
                        }
                        if salvar_validacao(registro):
                            st.success("Validação salva! Recalcule a base para aplicar as alterações.")
                            st.rerun()
        else:
            st.info("Carregue a base histórica primeiro para registrar validações.")

    # ── Section: Validações existentes ──
    with st.expander("📋 Validações registradas", expanded=True):
        if validacoes_existentes.empty:
            st.info("Nenhuma validação manual registrada ainda.")
        else:
            st.dataframe(
                validacoes_existentes[[c for c in COLUNAS_VALIDACAO if c in validacoes_existentes.columns]],
                use_container_width=True, hide_index=True,
            )

            csv_val = validacoes_existentes.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    "\U0001f4e5 Exportar Validações (csv)",
                    data=csv_val, file_name="validacoes_manuais.csv",
                    mime="text/csv", use_container_width=True,
                )
            with col_d2:
                ids_para_remover = st.multiselect(
                    "Selecionar IDs para remover",
                    validacoes_existentes["id_registro"].tolist(),
                    key="val_rem_ids",
                )
                if ids_para_remover and st.button("🗑️ Remover selecionadas", use_container_width=True):
                    for rid in ids_para_remover:
                        remover_validacao(rid)
                    st.success(f"Removidas {len(ids_para_remover)} validações.")
                    st.rerun()

    # ── Section: Recalcular ──
    st.markdown("---")
    st.subheader("🔄 Aplicar Validações e Recalcular")
    st.warning(
        "Ao recalcular, todas as validações manuais serão aplicadas "
        "como uma camada sobre os dados brutos. Lançamentos marcados "
        "como 'excluir_do_calculo' ou 'marcar_como_duplicado' serão "
        "removidos dos cálculos."
    )

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        if st.button("🔁 Recalcular base com validações", type="primary", use_container_width=True):
            if st.session_state.dados is not None:
                with st.spinner("Reaplicando validações e recalculando..."):
                    lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist = (
                        _carregar_do_consolidado(CONSOLIDADO_PATH)
                        if CONSOLIDADO_PATH.exists()
                        else _carregar_historico_renamed()
                    )
                    dados = _processar_historico(
                        lanc_hist, res_hist, nao_hist, ext_hist, df_aud_hist,
                        mostrar_status=False, aplicar_validacoes=True,
                    )
                    if dados is not None:
                        st.session_state.dados = dados
                st.rerun()

    # ── Impact Log ──
    with st.expander("📊 Log de Impacto das Validações", expanded=False):
        from src.services.manual_validation import gerar_impacto
        impact_path = Path("output") / "impacto_validacoes.csv"
        if impact_path.exists():
            df_impacto = pd.read_csv(impact_path, encoding="utf-8-sig")
            if not df_impacto.empty:
                st.dataframe(df_impacto, use_container_width=True, hide_index=True)
                csv_imp = df_impacto.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    "\U0001f4e5 Exportar Impacto (csv)",
                    data=csv_imp, file_name="impacto_validacoes.csv",
                    mime="text/csv", use_container_width=True,
                )
            else:
                st.info("Nenhum impacto registrado ainda.")
        else:
            st.info("Recalcule a base para gerar o log de impacto.")
