import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import math
import altair as alt
from datetime import datetime, timedelta
from pathlib import Path

# =============================================================================
# CONFIGURAÇÃO DA PÁGINA
# =============================================================================
st.set_page_config(page_title="Gestor HUUFMA PRO", layout="wide", page_icon="🏥")

st.markdown("""
    <style>
    .main .block-container { padding-top: 2rem; }
    h1 { color: #1E3A8A; font-weight: 700; margin-bottom: 0.5rem; }
    h2, h3 { color: #2C3E50; font-weight: 600; }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 1rem 1.25rem;
        border-radius: 0.75rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
    }
    div.stButton > button:first-child {
        background-color: #10B981 !important;
        color: white !important;
        font-weight: 600 !important;
        border: none !important;
        padding: 0.5rem 2rem !important;
        border-radius: 0.5rem !important;
        transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        background-color: #059669 !important;
        transform: translateY(-1px);
    }
    div[data-testid="stDataEditor"] { border-radius: 0.5rem; }
    </style>
""", unsafe_allow_html=True)


# =============================================================================
# FONTE ÚNICA GLOBAL DE STATUS E CORES
# Tupla por status: (cor_grafico, bg_excel, fg_excel)
# Alterar aqui reflete automaticamente em gráficos e nos dois Excels gerados.
# =============================================================================
MAPA_STATUS = {
    "Estoque Suficiente":       ("#A2E8A2", "#E2EFDA", "#375623"),
    "Solicitar":                ("#A6C8FF", "#DDEBF7", "#1F4E78"),
    "Estoque Crítico CAF":      ("#FFC499", "#FCE4D6", "#C65911"),
    "Remanejar":                ("#FFEAA6", "#FFF2CC", "#7F6000"),
    "Desabastecimento Crítico": ("#FFA6A6", "#F8CBAD", "#C00000"),
    "Estoque Excessivo":        ("#B2EBF2", "#E5F1F4", "#006666"),
    "Estoque Parado":           ("#CFD8DC", "#F2F2F2", "#595959"),
    "Estoque em Alerta":        ("#FFEB3B", "#FFF9C4", "#7F6000"),
    "Sem Consumo":              ("#ECEFF1", "#F5F5F5", "#9E9E9E"),
}
STATUS_CORES = {k: v[0] for k, v in MAPA_STATUS.items()}
EXCEL_CORES  = {k: (v[1], v[2]) for k, v in MAPA_STATUS.items()}

# Categorias disponíveis no sistema — edite aqui para adicionar novas
CATEGORIAS_DISPONIVEIS = sorted([
    "MEDICAMENTO", "MMH", "SORO", "NUTRIÇÃO", "GASES MEDICINAIS",
    "MATERIAL DIAGNÓSTICO", "OUTROS",
])

# =============================================================================
# CAMINHO DO ARQUIVO BASE DE CATEGORIAS
# Fica na mesma pasta do app.py — criado automaticamente se não existir.
# =============================================================================
ARQUIVO_CATEGORIAS = Path(__file__).parent / "categorias_base.xlsx"


# =============================================================================
# FUNÇÕES UTILITÁRIAS
# =============================================================================

def clean(t: str) -> str:
    if pd.isna(t):
        return ""
    nfkd = unicodedata.normalize('NFKD', str(t))
    t = "".join([c for c in nfkd if not unicodedata.combining(c)])
    t = re.sub(r'[^a-zA-Z0-9\s]', ' ', t)
    return " ".join(t.lower().split())


def clean_key(v) -> str:
    return re.sub(r'[^0-9]', '', str(v)).lstrip('0')


def p_num(v) -> float:
    try:
        if pd.isna(v) or str(v).strip() == "":
            return 0.0
        l = re.sub(r'[^0-9,.]', '', str(v))
        if "," in l and "." in l:
            l = l.replace(".", "").replace(",", ".")
        elif "," in l:
            l = l.replace(",", ".")
        return float(l) if l else 0.0
    except Exception:
        return 0.0


def find_col(df: pd.DataFrame, terms: list, forbidden: list = []):
    for col in df.columns:
        col_clean = clean(col)
        if any(t in col_clean for t in terms) and not any(f in col_clean for f in forbidden):
            return col
    return None


def validar_colunas(colunas: dict, contexto: str) -> bool:
    faltando = [nome for nome, col in colunas.items() if col is None]
    if faltando:
        st.error(
            f"❌ **Erro no arquivo '{contexto}':** Colunas não identificadas: "
            f"`{'`, `'.join(faltando)}`. Verifique os cabeçalhos."
        )
        return False
    return True


# =============================================================================
# SISTEMA DE CATEGORIAS — 3 CAMADAS
# Camada 1: arquivo categorias_base.xlsx (persistência em disco)
# Camada 2: st.session_state['df_categorias'] (memória de sessão editável)
# Camada 3: st.data_editor (edição direta na interface)
# =============================================================================

def _df_categorias_vazio() -> pd.DataFrame:
    """Retorna estrutura vazia padrão do DataFrame de categorias."""
    return pd.DataFrame(columns=["Código", "Material", "Categoria"])


def carregar_categorias_do_disco() -> pd.DataFrame:
    """
    Lê o arquivo categorias_base.xlsx.
    Se não existir, retorna DataFrame vazio e cria o arquivo.
    """
    if ARQUIVO_CATEGORIAS.exists():
        try:
            df = pd.read_excel(ARQUIVO_CATEGORIAS, dtype=str)
            col_cod = find_col(df, ['cod', 'ca3', 'material'], forbidden=['descri', 'prod'])
            col_cat = find_col(df, ['categ'])
            col_mat = find_col(df, ['descri', 'material', 'produto', 'nome'])

            if col_cod is None or col_cat is None:
                st.warning("⚠️ categorias_base.xlsx encontrado mas sem colunas 'Código' e 'Categoria' reconhecíveis. Iniciando vazio.")
                return _df_categorias_vazio()

            df_clean = pd.DataFrame()
            df_clean["Código"]    = df[col_cod].apply(clean_key)
            df_clean["Material"]  = df[col_mat].fillna("").astype(str) if col_mat else ""
            df_clean["Categoria"] = df[col_cat].str.upper().str.strip()
            df_clean = df_clean[df_clean["Código"] != ""].drop_duplicates("Código")
            return df_clean.reset_index(drop=True)
        except Exception as e:
            st.warning(f"⚠️ Erro ao ler categorias_base.xlsx: {e}. Iniciando vazio.")
            return _df_categorias_vazio()
    return _df_categorias_vazio()


def salvar_categorias_no_disco(df: pd.DataFrame):
    """Salva o DataFrame atual de categorias no arquivo base."""
    try:
        df.to_excel(ARQUIVO_CATEGORIAS, index=False)
        return True
    except Exception as e:
        st.error(f"❌ Não foi possível salvar o arquivo: {e}")
        return False


def inicializar_categorias_session():
    """Carrega categorias do disco para o session_state se ainda não carregado."""
    if "df_categorias" not in st.session_state:
        st.session_state["df_categorias"] = carregar_categorias_do_disco()


def obter_mapa_categorias() -> dict:
    """Retorna {código: categoria} a partir do session_state atual."""
    df = st.session_state.get("df_categorias", _df_categorias_vazio())
    if df.empty:
        return {}
    return dict(zip(df["Código"].astype(str), df["Categoria"].astype(str)))


def enriquecer_categorias_com_estoque(df_est: pd.DataFrame, c_est_cod: str, c_est_prod: str):
    """
    Ao processar pedido, adiciona itens novos encontrados no estoque
    que ainda não estão no mapa de categorias (com categoria 'OUTROS').
    Permite que o gestor os classifique depois na aba de categorias.
    """
    df_cat = st.session_state.get("df_categorias", _df_categorias_vazio()).copy()
    codigos_existentes = set(df_cat["Código"].astype(str).tolist())

    # Vetorizado — sem iterrows
    df_unicos      = df_est.drop_duplicates(subset=["key"]).copy()
    df_unicos["_ks"] = df_unicos["key"].astype(str)
    df_novos_raw   = df_unicos[~df_unicos["_ks"].isin(codigos_existentes)]

    if not df_novos_raw.empty:
        df_novos = pd.DataFrame({
            "Código":    df_novos_raw["_ks"].values,
            "Material":  df_novos_raw[c_est_prod].fillna("").astype(str).values,
            "Categoria": "OUTROS",
        })
        df_cat = pd.concat([df_cat, df_novos], ignore_index=True).drop_duplicates("Código")
        st.session_state["df_categorias"] = df_cat


# =============================================================================
# LEITURA DE ARQUIVOS COM CACHE
# =============================================================================

@st.cache_data(show_spinner=False)
def ler_csv_cached(file_bytes: bytes, nome: str) -> pd.DataFrame:
    return pd.read_csv(
        io.BytesIO(file_bytes), sep=None, engine='python',
        encoding='latin1', index_col=False
    )


@st.cache_data(show_spinner=False)
def ler_xlsx_cached(file_bytes: bytes, nome: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), dtype=str)


# =============================================================================
# LÓGICA DE NEGÓCIO
# =============================================================================

def calcular_cmd(qtd_total: float, dias: int) -> float:
    cmd_bruto = qtd_total / max(dias, 1)
    return float(math.ceil(cmd_bruto)) if cmd_bruto > 0 else 0.0


def calcular_sugestao(row: pd.Series, dias_pedido: int) -> int:
    """
    Corrigido: cmd=0 tratado separadamente para evitar pedido desnecessário
    em itens sem giro mas com estoque mínimo parametrizado.
    """
    cmd        = row['Consumo Médio Diário']
    est_atual  = row['Saldo Atual Satélite']
    est_minimo = row['Estoque Mínimo']

    if cmd == 0 and est_minimo > 0:
        return max(0, round(est_minimo - est_atual))
    if cmd == 0:
        return 0
    meta_final = max(cmd * dias_pedido, est_minimo)
    return max(0, round(meta_final - est_atual))


def definir_alerta_e_acao(
    row: pd.Series,
    dict_saldos_centrais: dict,
    dict_saldos_parceiras: dict,
    consumo_outras_total: dict,
) -> tuple:
    """
    Motor de inteligência logística.
    Avalia centrais 1, 6, 9, 41 e 43.
    """
    cod        = row['Código MV']
    sug        = row['Necessidade de Ressuprimento']
    cmd        = row['Consumo Médio Diário']
    est_un     = row['Saldo Atual Satélite']
    est_minimo = row['Estoque Mínimo']

    if cmd == 0 and est_minimo <= 0 and est_un == 0:
        return "Sem Consumo", "Avaliar se é necessário inativar o item na farmácia."
    if cmd > 0 and est_un > (cmd * 60):
        return "Estoque Excessivo", "Estoque cobre mais de 60 dias. Avaliar devolução."
    if cmd == 0 and est_un > 0:
        if est_un <= est_minimo:
            return "Estoque Parado", "Item sem consumo, mas dentro do estoque mínimo parametrizado."
        excedente = int(est_un - est_minimo)
        return "Estoque Parado", f"{excedente} un. acima do mínimo. Considerar devolução ou remanejamento."
    if sug <= 0:
        if est_un < est_minimo:
            return "Estoque em Alerta", "Estoque abaixo do mínimo de segurança. Monitorar giro."
        return "Estoque Suficiente", "Estoque dentro da cobertura ideal."

    saldos_nas_centrais = dict_saldos_centrais.get(cod, {})
    saldo_total_central = sum(saldos_nas_centrais.values())

    if saldo_total_central > 0:
        central_principal = max(saldos_nas_centrais, key=saldos_nas_centrais.get)
        if saldo_total_central >= sug:
            return "Solicitar", f"Solicitar {int(sug)} un. ao Almoxarifado Central {central_principal}."
        return "Estoque Crítico CAF", (
            f"Pegar {int(saldo_total_central)} un. no Almox Central {central_principal} "
            f"e remanejar o restante."
        )

    saldos_parceiras   = dict_saldos_parceiras.get(cod, {})
    consumos_parceiras = consumo_outras_total.get(cod, {})
    farmacias_paradas  = [
        f"Cód {fid} ({int(sf)} un.)"
        for fid, sf in saldos_parceiras.items()
        if sf > 0 and (
            consumos_parceiras.get(fid, 0) == 0 or
            sf > consumos_parceiras.get(fid, 0) * 3
        )
    ]
    if farmacias_paradas:
        return "Remanejar", f"Central Zerada! Transferir de: {', '.join(farmacias_paradas)}."

    return "Desabastecimento Crítico", "Sem saldo na central e sem estoque parado em parceiras."


# =============================================================================
# EXPORTAÇÃO EXCEL — função auxiliar reutilizada em todos os relatórios
# Formatos criados UMA VEZ por workbook (fora do loop de abas).
# =============================================================================

def _criar_formatos(workbook):
    """Cria e retorna os 3 formatos base do Excel (base, texto, header)."""
    fmt_base = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'border': 1, 'border_color': '#D0D0D0', 'font_size': 10,
    })
    fmt_texto = workbook.add_format({
        'align': 'left', 'valign': 'vcenter', 'text_wrap': True,
        'border': 1, 'border_color': '#D0D0D0', 'font_size': 10,
    })
    fmt_header = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'bg_color': '#1E3A8A', 'font_color': '#FFFFFF', 'border': 1, 'font_size': 11,
    })
    return fmt_base, fmt_texto, fmt_header


def _aplicar_aba(workbook, worksheet, df, larguras: dict,
                 fmt_base, fmt_texto, fmt_header, col_alerta: str):
    """
    Aplica formatação completa em uma aba Excel.
    Recebe formatos já criados (não recria por aba — evita limite xlsxwriter).
    """
    # Cabeçalho
    for ci, cn in enumerate(df.columns):
        worksheet.write(0, ci, cn, fmt_header)
    worksheet.set_row(0, 40)

    # Larguras e formatos por coluna
    for ci, cn in enumerate(df.columns):
        larg    = larguras.get(cn, 20)
        fmt_col = fmt_texto if cn in ('Material', 'Ação Logística Sugerida') else fmt_base
        worksheet.set_column(ci, ci, larg, fmt_col)

    # Altura das linhas de dados
    for ri in range(1, len(df) + 1):
        worksheet.set_row(ri, 30)

    worksheet.freeze_panes(1, 0)

    # Formatação condicional — linha inteira colorida pelo status (letra calculada dinamicamente)
    if col_alerta in df.columns:
        idx_al  = df.columns.get_loc(col_alerta)
        letra   = chr(ord('A') + idx_al)
        n_cols  = len(df.columns) - 1
        for status, (bg, fg) in EXCEL_CORES.items():
            fmt_cond = workbook.add_format({
                'bg_color': bg, 'font_color': fg, 'align': 'center', 'valign': 'vcenter',
                'text_wrap': True, 'bold': True, 'border': 1, 'border_color': '#D0D0D0',
                'font_size': 10,
            })
            worksheet.conditional_format(
                1, 0, len(df), n_cols,
                {'type': 'formula', 'criteria': f'=${letra}2="{status}"', 'format': fmt_cond}
            )


def exportar_excel_aba_unica(df: pd.DataFrame, nome_aba: str, larguras: dict,
                              col_alerta: str = "Parecer Logístico / Alerta") -> bytes:
    """Gera Excel de aba única com formatação completa."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
        df.to_excel(wr, sheet_name=nome_aba, index=False)
        fmt_base, fmt_texto, fmt_header = _criar_formatos(wr.book)
        _aplicar_aba(wr.book, wr.sheets[nome_aba], df, larguras,
                     fmt_base, fmt_texto, fmt_header, col_alerta)
    return buf.getvalue()


def exportar_excel_multi_aba(df_total: pd.DataFrame, ordem_cols: list,
                              col_categoria: str, col_alerta: str,
                              larguras: dict, excluir_alertas: list = None) -> bytes:
    """
    Gera Excel multi-aba separado por categoria.
    Formatos criados UMA VEZ fora do loop — corrige risco de limite xlsxwriter.
    """
    buf = io.BytesIO()
    excluir_alertas = excluir_alertas or []

    with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
        # Formatos criados UMA VEZ por workbook
        fmt_base, fmt_texto, fmt_header = _criar_formatos(wr.book)

        for cat, df_cat in df_total.groupby(col_categoria):
            df_exp = df_cat.copy()
            if excluir_alertas:
                df_exp = df_exp[~df_exp[col_alerta].isin(excluir_alertas)]

            # Filtra apenas colunas que existem no df
            cols_existentes = [c for c in ordem_cols if c in df_exp.columns]
            df_exp = df_exp[cols_existentes].sort_values('Material').copy()

            if df_exp.empty:
                continue

            nome_aba = str(cat)[:31]
            df_exp.to_excel(wr, sheet_name=nome_aba, index=False)
            _aplicar_aba(wr.book, wr.sheets[nome_aba], df_exp, larguras,
                         fmt_base, fmt_texto, fmt_header, col_alerta)

    return buf.getvalue()


# =============================================================================
# SIDEBAR
# =============================================================================
with st.sidebar:
    st.markdown("### 🏥 Unidade & Parâmetros")
    farmacias_opcoes = {
        "Farmácia UMI (Cód. 13)":              "13",
        "Farmácia Dutra (Cód. 31)":            "31",
        "Farmácia Centro Cirúrgico (Cód. 7)":  "7",
        "Farmácia Oftalmologia (Cód. 39)":     "39",
        "Farmácia UTI (Cód. 34)":              "34",
    }
    farmacia_selecionada = st.selectbox("Defina a Farmácia:", list(farmacias_opcoes.keys()))
    cod_farmacia_alvo    = farmacias_opcoes[farmacia_selecionada]

    st.write("---")
    st.markdown("### ⚙️ Parâmetros do Pedido")
    dias_pedido = st.number_input("Dias de ressuprimento:", value=15, min_value=1)

    hoje        = datetime.now()
    data_inicio = st.date_input("Início do Histórico:", value=hoje - timedelta(days=5), format="DD/MM/YYYY")
    data_fim    = st.date_input("Fim do Período:",      value=hoje,                        format="DD/MM/YYYY")

# Inicializa categorias do disco na sessão
inicializar_categorias_session()

# =============================================================================
# ESTRUTURA VISUAL PRINCIPAL
# =============================================================================
st.title("🏥 Gestor de Pedidos Logístico Avançado")
st.markdown(
    f"**Farmácia Ativa:** `{farmacia_selecionada}` | "
    f"**Janela Histórica:** `{data_inicio.strftime('%d/%m/%Y')}` até `{data_fim.strftime('%d/%m/%Y')}`"
)
st.write("")

tab1, tab2 = st.tabs(["⚡ Processar Pedido com IA Logística", "🗂️ Gestão de Categorias de Insumos"])


# =============================================================================
# TAB 2 — GESTÃO DE CATEGORIAS (NOVA IMPLEMENTAÇÃO)
# =============================================================================
with tab2:
    st.subheader("🗂️ Mapeamento Global de Categorias de Insumos")
    st.info(
        "As categorias aqui definidas são **globais** — independem da farmácia. "
        "Um item classificado como **MEDICAMENTO** será tratado assim em todas as unidades. "
        "Itens são excluídos da análise apenas quando não há **consumo registrado no período** "
        "e **nenhum estoque mínimo parametrizado** — ou seja, itens completamente inativos naquela unidade. "
        "Itens com estoque mínimo definido, mesmo sem giro recente, continuam visíveis e recebem parecer de alerta ou desabastecimento."
    )

    df_cat_session = st.session_state["df_categorias"].copy()

    # --- Métricas rápidas ---
    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("📦 Total de Itens Cadastrados", len(df_cat_session))
    mc2.metric("🏷️ Categorias Ativas", df_cat_session["Categoria"].nunique() if not df_cat_session.empty else 0)
    mc3.metric("❓ Itens sem Classificação (OUTROS)",
               len(df_cat_session[df_cat_session["Categoria"] == "OUTROS"]) if not df_cat_session.empty else 0)

    st.write("---")

    # --- Importação em lote ---
    with st.expander("📥 Importar Categorias em Lote (Excel/CSV)", expanded=False):
        st.markdown(
            "Suba um arquivo com as colunas **Código** e **Categoria**. "
            "Itens já existentes terão sua categoria atualizada. Novos itens serão adicionados."
        )
        arq_import = st.file_uploader("Selecionar arquivo:", type=["xlsx", "csv"], key="import_cat")
        if arq_import:
            with st.status("⏳ Processando arquivo de categorias...", expanded=True) as status_import:
                try:
                    st.write("📂 Lendo o arquivo...")
                    if arq_import.name.endswith(".csv"):
                        df_import = pd.read_csv(io.BytesIO(arq_import.read()), sep=None,
                                                engine='python', dtype=str)
                    else:
                        df_import = pd.read_excel(io.BytesIO(arq_import.read()), dtype=str)

                    st.write(f"🔍 Identificando colunas em `{arq_import.name}`...")
                    col_cod_i = find_col(df_import, ['cod', 'ca3'])
                    col_cat_i = find_col(df_import, ['categ'])
                    col_mat_i = find_col(df_import, ['descri', 'material', 'produto', 'nome'])

                    if not col_cod_i or not col_cat_i:
                        status_import.update(label="❌ Colunas não identificadas", state="error")
                        st.error("O arquivo precisa ter colunas reconhecíveis de **Código** e **Categoria**.")
                    else:
                        st.write("🔄 Normalizando e mesclando com a base atual...")
                        df_imp_clean = pd.DataFrame({
                            "Código":    df_import[col_cod_i].apply(clean_key),
                            "Material":  df_import[col_mat_i].fillna("").astype(str) if col_mat_i else "",
                            "Categoria": df_import[col_cat_i].str.upper().str.strip(),
                        })
                        df_imp_clean = df_imp_clean[df_imp_clean["Código"] != ""]

                        novos     = df_imp_clean[~df_imp_clean["Código"].isin(
                                        st.session_state["df_categorias"]["Código"])]["Código"].count()
                        atualizados = len(df_imp_clean) - novos

                        # Merge: atualiza existentes e adiciona novos
                        df_base   = st.session_state["df_categorias"].copy()
                        df_merged = pd.concat([df_base, df_imp_clean], ignore_index=True)
                        df_merged = df_merged.drop_duplicates("Código", keep="last")
                        st.session_state["df_categorias"] = df_merged.reset_index(drop=True)

                        st.write(f"✅ **{novos} novos itens** adicionados · **{atualizados} itens** atualizados.")
                        status_import.update(
                            label=f"✅ Importação concluída — {len(df_imp_clean)} itens processados",
                            state="complete", expanded=False
                        )
                        st.rerun()
                except Exception as e:
                    status_import.update(label="❌ Erro durante a importação", state="error")
                    st.error(f"Detalhe técnico: `{e}`")

    st.write("")

    # --- Filtros da tabela de categorias ---
    fc1, fc2 = st.columns([3, 2])
    filtro_nome_cat = fc1.text_input("🔍 Buscar por código ou descrição:", key="filtro_cat_nome")
    filtro_cat_sel  = fc2.selectbox("🏷️ Filtrar por categoria:", ["TODAS"] + CATEGORIAS_DISPONIVEIS, key="filtro_cat_sel")

    df_exibir = df_cat_session.copy()
    if filtro_nome_cat:
        mask = (
            df_exibir["Código"].astype(str).str.contains(filtro_nome_cat, case=False, na=False) |
            df_exibir["Material"].astype(str).str.contains(filtro_nome_cat, case=False, na=False)
        )
        df_exibir = df_exibir[mask]
    if filtro_cat_sel != "TODAS":
        df_exibir = df_exibir[df_exibir["Categoria"] == filtro_cat_sel]

    st.markdown(f"##### 📋 Tabela de Categorias ({len(df_exibir)} itens exibidos)")

    # --- Tabela editável (st.data_editor) ---
    df_editado = st.data_editor(
        df_exibir.reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Código":    st.column_config.TextColumn("Código MV", width="small"),
            "Material":  st.column_config.TextColumn("Descrição do Insumo", width="large"),
            "Categoria": st.column_config.SelectboxColumn(
                "Categoria", width="medium",
                options=CATEGORIAS_DISPONIVEIS, required=True
            ),
        },
        key="editor_categorias",
    )

    st.write("")
    col_salvar1, col_salvar2, col_salvar3 = st.columns([2, 2, 1])

    # Botão: aplicar edições na sessão e salvar no disco
    if col_salvar1.button("💾 SALVAR ALTERAÇÕES", use_container_width=True):
        # Atualiza apenas os itens visíveis (respeitando filtros ativos)
        df_base_atual = st.session_state["df_categorias"].copy()
        codigos_exibidos = set(df_exibir["Código"].astype(str).tolist())

        # Remove linhas editadas/excluídas do base e substitui pelas editadas
        df_base_sem_editados = df_base_atual[
            ~df_base_atual["Código"].astype(str).isin(codigos_exibidos)
        ]
        df_final_cat = pd.concat(
            [df_base_sem_editados, df_editado], ignore_index=True
        ).drop_duplicates("Código", keep="last")

        df_final_cat = df_final_cat[df_final_cat["Código"].astype(str).str.strip() != ""]
        st.session_state["df_categorias"] = df_final_cat.reset_index(drop=True)

        if salvar_categorias_no_disco(st.session_state["df_categorias"]):
            st.success("✅ Categorias salvas com sucesso no arquivo base!")
            st.rerun()

    # Botão: exportar mapa atual
    if not df_cat_session.empty:
        excel_cat = exportar_excel_aba_unica(
            df_cat_session,
            "Categorias",
            {"Código": 12, "Material": 45, "Categoria": 20},
            col_alerta="",
        )
        col_salvar2.download_button(
            "📤 Exportar Mapa Atual (.xlsx)",
            data=excel_cat,
            file_name=f"Categorias_HUUFMA_{datetime.now().strftime('%d%m%y')}.xlsx",
            use_container_width=True,
        )

    # Botão: recarregar do disco (desfaz edições não salvas)
    if col_salvar3.button("🔄 Recarregar", use_container_width=True, help="Descarta edições não salvas e recarrega o arquivo base"):
        st.session_state.pop("df_categorias", None)
        st.rerun()


# =============================================================================
# TAB 1 — PROCESSAR PEDIDO
# =============================================================================
with tab1:
    with st.container(border=True):
        st.markdown("##### 📥 Upload das Fontes de Dados Obrigatórias")
        col1, col2 = st.columns(2)
        file_mov_alvo  = col1.file_uploader("1. Movimento da Farmácia Alvo (.csv)", type=["csv"])
        file_est_geral = col2.file_uploader("2. Estoque Geral de todos os Almoxarifados (.csv)", type=["csv"])
        st.write("")
        files_mov_parceiras = st.file_uploader(
            "3. Movimentos das Outras Farmácias (Opcional — Múltiplos .csv)",
            type=["csv"], accept_multiple_files=True,
        )

    st.write("")

    if file_mov_alvo and file_est_geral:
        if st.button("🚀 GERAR PEDIDO COM INTELIGÊNCIA LOGÍSTICA", use_container_width=True):
            st.session_state["disparar_processamento"] = True

        if st.session_state.get("disparar_processamento", False):
            progress = st.progress(0, text="📂 Lendo arquivos...")

            try:
                # ------------------------------------------------------------------
                # ETAPA 1 — Leitura com cache
                # ------------------------------------------------------------------
                mov       = ler_csv_cached(file_mov_alvo.read(), file_mov_alvo.name)
                est_geral = ler_csv_cached(file_est_geral.read(), file_est_geral.name)

                progress.progress(10, text="🔍 Identificando colunas...")

                # ------------------------------------------------------------------
                # ETAPA 2 — Validação de colunas
                # ------------------------------------------------------------------
                cols_mov = {
                    'código':       find_col(mov, ['material', 'cod', 'ca3']),
                    'quantidade':   find_col(mov, ['quant']),
                    'tipo':         find_col(mov, ['tipo']),
                    'data':         find_col(mov, ['data', 'ger']),
                    'almoxarifado': find_col(mov, ['almox']),
                }
                if not validar_colunas(cols_mov, file_mov_alvo.name):
                    st.session_state["disparar_processamento"] = False
                    st.stop()

                cols_est = {
                    'código':       find_col(est_geral, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod']),
                    'quantidade':   find_col(est_geral, ['qtde disp', 'disponivel']),
                    'produto':      find_col(est_geral, ['material', 'produto', 'descri']),
                    'almoxarifado': find_col(est_geral, ['almox']),
                    'mínimo':       find_col(est_geral, ['qtde estq min', 'estoque minimo', 'minimo']),
                }
                if not validar_colunas({k: v for k, v in cols_est.items() if k != 'mínimo'}, file_est_geral.name):
                    st.session_state["disparar_processamento"] = False
                    st.stop()

                c_mov_cod   = cols_mov['código']
                c_mov_qtd   = cols_mov['quantidade']
                c_mov_tipo  = cols_mov['tipo']
                c_mov_data  = cols_mov['data']
                c_est_cod   = cols_est['código']
                c_est_qtd   = cols_est['quantidade']
                c_est_prod  = cols_est['produto']
                c_est_almox = cols_est['almoxarifado']
                c_est_min   = cols_est['mínimo']

                progress.progress(20, text="🏗️ Processando estoque geral...")

                # ------------------------------------------------------------------
                # ETAPA 3 — Preparar estoque geral
                # ------------------------------------------------------------------
                est_geral = est_geral.copy()
                est_geral['key']         = est_geral[c_est_cod].apply(clean_key)
                est_geral['almox_limpo'] = est_geral[c_est_almox].apply(clean_key)
                est_geral['saldo_num']   = est_geral[c_est_qtd].apply(p_num)
                est_geral['min_num']     = est_geral[c_est_min].apply(p_num) if c_est_min else 0.0

                est_farmacia_alvo = (
                    est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo]
                    .groupby('key')['saldo_num'].sum().to_dict()
                )
                est_min_alvo = (
                    est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo]
                    .groupby('key')['min_num'].sum().to_dict()
                )

                # Centrais vetorizadas — CORREÇÃO include_groups (pandas 2.2+)
                centrais_alvo         = ['1', '6', '9', '41', '43']
                est_centrais_filtrado = est_geral[est_geral['almox_limpo'].isin(centrais_alvo)]
                if not est_centrais_filtrado.empty:
                    _c_piv = (
                        est_centrais_filtrado.groupby(['key', 'almox_limpo'])['saldo_num']
                        .sum().reset_index()
                    )
                    dict_saldos_centrais = (
                        _c_piv.groupby('key')[['almox_limpo', 'saldo_num']]
                        .apply(lambda g: dict(zip(g['almox_limpo'], g['saldo_num'])),
                               include_groups=False)
                        .to_dict()
                    )
                else:
                    dict_saldos_centrais = {}

                # Parceiras vetorizadas — CORREÇÃO include_groups
                cod_parceiras = [c for c in ['7', '13', '31', '34', '39'] if c != cod_farmacia_alvo]
                est_outras    = est_geral[est_geral['almox_limpo'].isin(cod_parceiras)]
                if not est_outras.empty:
                    _p_piv = (
                        est_outras.groupby(['key', 'almox_limpo'])['saldo_num']
                        .sum().reset_index()
                    )
                    dict_saldos_parceiras = (
                        _p_piv.groupby('key')[['almox_limpo', 'saldo_num']]
                        .apply(lambda g: dict(zip(g['almox_limpo'], g['saldo_num'])),
                               include_groups=False)
                        .to_dict()
                    )
                else:
                    dict_saldos_parceiras = {}

                progress.progress(40, text="📊 Calculando consumo...")

                # ------------------------------------------------------------------
                # ETAPA 4 — Consumo da farmácia alvo
                # ------------------------------------------------------------------
                mov = mov.copy()
                mov['dt_formatada'] = pd.to_datetime(mov[c_mov_data], dayfirst=True, errors='coerce')
                mov_filtrado = mov[
                    (mov['dt_formatada'].dt.date >= data_inicio) &
                    (mov['dt_formatada'].dt.date <= data_fim) &
                    (mov[c_mov_tipo].astype(str).str.upper() == 'RM')
                ].copy()

                dias_considerados = max((data_fim - data_inicio).days + 1, 1)

                consumo = (
                    mov_filtrado
                    .assign(qtd_num=lambda df: df[c_mov_qtd].apply(p_num))
                    .assign(key=lambda df: df[c_mov_cod].apply(clean_key))
                    .groupby('key')['qtd_num'].sum()
                    .reset_index()
                    .rename(columns={'qtd_num': 'total_consumo'})
                )
                consumo['cmd'] = consumo['total_consumo'].apply(
                    lambda x: calcular_cmd(x, dias_considerados)
                )

                progress.progress(55, text="🔄 Cruzando dados entre parceiras...")

                # ------------------------------------------------------------------
                # ETAPA 5 — Consumo parceiras
                # CORREÇÃO: groupby duplo antes do zip garante soma correta
                # CORREÇÃO: include_groups=False para pandas 2.2+
                # ------------------------------------------------------------------
                consumo_outras_total = {}
                if files_mov_parceiras:
                    for f_parc in files_mov_parceiras:
                        try:
                            df_p = ler_csv_cached(f_parc.read(), f_parc.name)

                            c_p_cod   = find_col(df_p, ['material', 'cod', 'ca3'])
                            c_p_qtd   = find_col(df_p, ['quant'])
                            c_p_tipo  = find_col(df_p, ['tipo'])
                            c_p_almox = find_col(df_p, ['almox'])
                            c_p_data  = find_col(df_p, ['data', 'ger'])

                            if not all([c_p_cod, c_p_qtd, c_p_tipo, c_p_almox, c_p_data]):
                                st.warning(f"⚠️ '{f_parc.name}' ignorado: colunas não identificadas.")
                                continue

                            df_p = df_p.copy()
                            df_p['dt_formatada'] = pd.to_datetime(df_p[c_p_data], dayfirst=True, errors='coerce')
                            df_p_filt = df_p[
                                (df_p['dt_formatada'].dt.date >= data_inicio) &
                                (df_p['dt_formatada'].dt.date <= data_fim) &
                                (df_p[c_p_tipo].astype(str).str.upper() == 'RM')
                            ].assign(
                                key=lambda d: d[c_p_cod].apply(clean_key),
                                almox_limpo=lambda d: d[c_p_almox].apply(clean_key),
                                qtd_num=lambda d: d[c_p_qtd].apply(p_num),
                            )

                            if df_p_filt.empty:
                                continue

                            # Agrupa key+almox ANTES do zip — garante soma, não sobrescreve
                            resumo_parc = (
                                df_p_filt.groupby(['key', 'almox_limpo'])['qtd_num']
                                .sum().reset_index()
                            )
                            parc_dict = (
                                resumo_parc.groupby('key')[['almox_limpo', 'qtd_num']]
                                .apply(lambda g: dict(zip(g['almox_limpo'], g['qtd_num'])),
                                       include_groups=False)
                                .to_dict()
                            )
                            for k, v in parc_dict.items():
                                for alm, qtd in v.items():
                                    consumo_outras_total.setdefault(k, {})
                                    consumo_outras_total[k][alm] = (
                                        consumo_outras_total[k].get(alm, 0) + qtd
                                    )
                        except Exception as e:
                            st.warning(f"⚠️ '{f_parc.name}' ignorado por erro: {e}")
                            continue

                progress.progress(70, text="🧠 Gerando análise final...")

                # ------------------------------------------------------------------
                # ETAPA 6 — Montar DataFrame final
                # ------------------------------------------------------------------
                mapa_produtos = (
                    est_geral.drop_duplicates(subset=['key'])
                    .set_index('key')[c_est_prod].to_dict()
                )
                consumo_map = consumo.set_index('key')['cmd'].to_dict()

                todos_codigos = sorted(list(
                    set(est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo]['key'].unique()) |
                    set(consumo['key'])
                ))

                final = pd.DataFrame({'Código MV': todos_codigos})
                final['Material']              = final['Código MV'].map(mapa_produtos).fillna('PRODUTO SEM DESCRIÇÃO')
                final['Saldo Atual Satélite']  = final['Código MV'].map(est_farmacia_alvo).fillna(0)
                final['Consumo Médio Diário']  = final['Código MV'].map(consumo_map).fillna(0)
                final['Estoque Mínimo']        = final['Código MV'].map(est_min_alvo).fillna(0)
                final['Necessidade de Ressuprimento'] = final.apply(
                    lambda row: calcular_sugestao(row, dias_pedido), axis=1
                )

                # Enriquece o mapa de categorias com itens novos do estoque
                enriquecer_categorias_com_estoque(est_geral, c_est_cod, c_est_prod)

                # Aplica categoria global (sem separação por farmácia)
                mapa_cat = obter_mapa_categorias()
                final['Categoria'] = final['Código MV'].map(mapa_cat).fillna('OUTROS')

                final['Saldo Almox. Centrais Unificado'] = final['Código MV'].apply(
                    lambda c: sum(dict_saldos_centrais.get(c, {}).values())
                )

                alertas_acoes = final.apply(
                    lambda row: definir_alerta_e_acao(
                        row, dict_saldos_centrais, dict_saldos_parceiras, consumo_outras_total
                    ), axis=1
                )
                final['Parecer Logístico / Alerta'] = [r[0] for r in alertas_acoes]
                final['Ação Logística Sugerida']    = [r[1] for r in alertas_acoes]

                st.session_state['df_final_huufma'] = final
                st.session_state['disparar_processamento'] = False
                progress.progress(100, text="✅ Processamento concluído!")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Erro crítico: {e}")
                st.session_state['disparar_processamento'] = False

        # ------------------------------------------------------------------
        # PAINEL DE RESULTADOS — renderizado do session_state (estável)
        # ------------------------------------------------------------------
        if 'df_final_huufma' in st.session_state:
            df_view  = st.session_state['df_final_huufma'].copy()
            COL_PAR  = 'Parecer Logístico / Alerta'
            COL_ACAO = 'Ação Logística Sugerida'
            COL_RES  = f'Necessidade de Ressuprimento ({dias_pedido} dias)'

            # Larguras usadas em todos os relatórios Excel
            LARGURAS_REL = {
                'Código MV': 12, 'Material': 45, 'Categoria': 16,
                'Saldo Atual Satélite': 20, 'Consumo Médio Diário': 20, 'Estoque Mínimo': 16,
                COL_RES: 26, 'Saldo Almox. Centrais Unificado': 28,
                COL_PAR: 26, COL_ACAO: 50,
            }

            # --- MÉTRICAS ---
            st.write("---")
            df_desabast = df_view[df_view[COL_PAR] == "Desabastecimento Crítico"].sort_values('Material')
            df_remanej  = df_view[df_view[COL_PAR] == "Remanejar"].sort_values('Material')
            df_caf      = df_view[df_view[COL_PAR].isin(["Solicitar", "Estoque Crítico CAF"])].sort_values('Material')
            df_excesso  = df_view[
                df_view[COL_PAR].isin(["Estoque Excessivo", "Estoque Parado", "Estoque em Alerta"])
            ].sort_values('Material')

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("🚨 Desabastecimento Crítico", f"{len(df_desabast)} itens")
                st.download_button("📥 Extrair", key="ex_c1", use_container_width=True,
                    data=exportar_excel_aba_unica(df_desabast, "Rupturas", LARGURAS_REL, COL_PAR),
                    file_name=f"Rupturas_{cod_farmacia_alvo}.xlsx")
            with c2:
                st.metric("🔄 Remanejamento Potencial", f"{len(df_remanej)} itens")
                st.download_button("📥 Extrair", key="ex_c2", use_container_width=True,
                    data=exportar_excel_aba_unica(df_remanej, "Remanejamento", LARGURAS_REL, COL_PAR),
                    file_name=f"Remanejamento_{cod_farmacia_alvo}.xlsx")
            with c3:
                st.metric("📦 Disponível no Almoxarifado", f"{len(df_caf)} itens")
                st.download_button("📥 Extrair", key="ex_c3", use_container_width=True,
                    data=exportar_excel_aba_unica(df_caf, "Disponiveis_CAF", LARGURAS_REL, COL_PAR),
                    file_name=f"Disponiveis_{cod_farmacia_alvo}.xlsx")
            with c4:
                st.metric("⚠️ Overstock / Sem Giro", f"{len(df_excesso)} itens")
                st.download_button("📥 Extrair", key="ex_c4", use_container_width=True,
                    data=exportar_excel_aba_unica(df_excesso, "Overstock", LARGURAS_REL, COL_PAR),
                    file_name=f"Overstock_{cod_farmacia_alvo}.xlsx")

            # --- GRÁFICOS ---
            st.write("")
            g1, g2 = st.columns(2)

            with g1:
                st.markdown("**Saúde Geral do Estoque**")
                df_g1 = (
                    df_view[df_view[COL_PAR] != 'Sem Consumo']
                    .groupby(COL_PAR)['Código MV'].count()
                    .reset_index().rename(columns={'Código MV': 'Quantidade', COL_PAR: 'Status'})
                )
                if not df_g1.empty:
                    grafico_donut = alt.Chart(df_g1).mark_arc(innerRadius=65, stroke='#fff').encode(
                        theta=alt.Theta('Quantidade:Q'),
                        color=alt.Color('Status:N', scale=alt.Scale(
                            domain=list(STATUS_CORES.keys()),
                            range=list(STATUS_CORES.values())
                        ), legend=alt.Legend(title="Parecer")),
                        tooltip=['Status:N', 'Quantidade:Q']
                    ).properties(height=280)
                    st.altair_chart(grafico_donut, use_container_width=True)

            with g2:
                st.markdown("**Urgência por Categoria**")
                df_g2 = (
                    df_view[~df_view[COL_PAR].isin(['Sem Consumo', 'Estoque Suficiente'])]
                    .groupby(['Categoria', COL_PAR])['Código MV'].count()
                    .reset_index().rename(columns={'Código MV': 'Itens', COL_PAR: 'Parecer'})
                )
                if not df_g2.empty:
                    grafico_stacked = alt.Chart(df_g2).mark_bar().encode(
                        x=alt.X('Itens:Q', title="Quantidade"),
                        y=alt.Y('Categoria:N', title=None, sort='-x'),
                        color=alt.Color('Parecer:N', scale=alt.Scale(
                            domain=list(STATUS_CORES.keys()),
                            range=list(STATUS_CORES.values())
                        ), legend=None),
                        tooltip=['Categoria:N', 'Parecer:N', 'Itens:Q']
                    ).properties(height=280)
                    st.altair_chart(grafico_stacked, use_container_width=True)

            # --- PAINEL INTERATIVO ---
            st.write("---")
            st.markdown("#### 📋 Painel de Análise Inteligente")

            with st.container(border=True):
                st.markdown("##### 🔍 Filtros Dinâmicos")
                f1, f2, f3 = st.columns([2, 2, 1])
                busca_nome  = f1.text_input("Filtrar por nome ou código:", key="busca_nome")
                busca_par   = f2.multiselect("Filtrar por Parecer:", options=list(STATUS_CORES.keys()), key="busca_par")
                busca_cat   = f3.selectbox("Categoria:", ["TODAS"] + sorted(df_view['Categoria'].unique().tolist()), key="busca_cat")

            df_filtrado = df_view.copy()
            if busca_nome:
                df_filtrado = df_filtrado[
                    df_filtrado['Material'].astype(str).str.contains(busca_nome, case=False, na=False) |
                    df_filtrado['Código MV'].astype(str).str.contains(busca_nome, case=False, na=False)
                ]
            if busca_par:
                df_filtrado = df_filtrado[df_filtrado[COL_PAR].isin(busca_par)]
            if busca_cat != "TODAS":
                df_filtrado = df_filtrado[df_filtrado['Categoria'] == busca_cat]

            st.dataframe(
                df_filtrado.sort_values('Material'),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Código MV":                      st.column_config.TextColumn("Código MV", width="small"),
                    "Material":                       st.column_config.TextColumn("Descrição", width="large"),
                    "Saldo Atual Satélite":           st.column_config.NumberColumn("Saldo Atual", format="%d"),
                    "Consumo Médio Diário":           st.column_config.NumberColumn("CMD", format="%d"),
                    "Estoque Mínimo":                 st.column_config.NumberColumn("Mínimo", format="%d"),
                    "Necessidade de Ressuprimento":   st.column_config.NumberColumn("Necessidade", format="%d"),
                    "Saldo Almox. Centrais Unificado":st.column_config.NumberColumn("Saldo Central", format="%d"),
                    COL_PAR: st.column_config.SelectboxColumn("Parecer", width="medium", options=list(STATUS_CORES.keys())),
                }
            )

            # --- DOWNLOADS FINAIS ---
            st.write("")

            # Prepara df renomeado para os Excels completos
            df_export = st.session_state['df_final_huufma'].rename(columns={
                'Necessidade de Ressuprimento': COL_RES
            })
            ordem_cols = [
                'Código MV', 'Material', 'Categoria',
                'Saldo Atual Satélite', 'Consumo Médio Diário', 'Estoque Mínimo',
                COL_RES, 'Saldo Almox. Centrais Unificado',
                COL_PAR, COL_ACAO,
            ]

            b1, b2 = st.columns(2)
            with b1:
                st.download_button(
                    label="📥 BAIXAR PAINEL COMPLETO — ABA ÚNICA (.XLSX)",
                    data=exportar_excel_aba_unica(df_export, "Painel Geral", LARGURAS_REL, COL_PAR),
                    file_name=f"Painel_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )
            with b2:
                st.download_button(
                    label="📥 BAIXAR MAPA COMPLETO — POR ABAS DE CATEGORIA (.XLSX)",
                    data=exportar_excel_multi_aba(
                        df_export, ordem_cols,
                        col_categoria='Categoria',
                        col_alerta=COL_PAR,
                        larguras=LARGURAS_REL,
                        excluir_alertas=["Sem Consumo"],
                    ),
                    file_name=f"Pedido_Abas_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )

    else:
        st.warning("⚠️ Aguardando upload dos arquivos obrigatórios para iniciar o processamento.")
