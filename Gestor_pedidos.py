import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import os
import math
import altair as alt
from datetime import datetime, timedelta
from pathlib import Path

# =============================================================================
# CONFIGURAÇÃO DA PÁGINA E ESTILOS CUSTOMIZADOS
# =============================================================================
st.set_page_config(page_title="Gestor HUUFMA PRO", layout="wide", page_icon="🏥")

st.markdown("""
    <style>
    /* Ajustes Gerais de Layout */
    .main .block-container { padding-top: 2rem; }
    h1 { color: #1E3A8A; font-weight: 700; margin-bottom: 0.5rem; }
    h2, h3 { color: #2C3E50; font-weight: 600; }
    
    /* Quadrantes de Métricas Premium */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 1rem 1.25rem;
        border-radius: 0.75rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
    }
    
    /* Botão Principal de Processamento */
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

    /* Customização dos componentes de Upload */
    div[data-testid="stFileUploaderFileData"] {
        padding: 4px 8px !important;
        margin-top: 4px !important;
        background-color: #F8FAFC !important;
        border-radius: 6px !important;
        border: 1px dashed #E2E8F0 !important;
    }
    div[data-testid="stFileUploaderFileName"] {
        font-size: 0.85rem !important;
        color: #334155 !important;
        font-weight: 500 !important;
    }
    div[data-testid="stFileUploaderFileData"] section {
        gap: 6px !important;
    }
    div[data-testid="stFileUploaderDropzone"] {
        padding: 1rem !important;
        border-radius: 8px !important;
    }
    </style>
""", unsafe_allow_html=True)

# Mapeamento Global de Nomes Práticos das Unidades do Hospital
DIC_NOMES_FARMACIAS = {
    "7":  "Farmácia Centro Cirúrgico",
    "13": "Farmácia UMI",
    "31": "Farmácia Dutra",
    "34": "Farmácia UTI",
    "39": "Farmácia Oftalmologia",
}

# Fonte única global de dados para status e cores (Interface, Fundo Excel, Fonte Excel)
MAPA_STATUS = {
    "Estoque Suficiente":       ("#A2E8A2", "#E2EFDA", "#375623"),
    "Solicitar":                ("#A6C8FF", "#DDEBF7", "#1F4E78"),
    "Remanejar":                ("#FFEAA6", "#FFF2CC", "#7F6000"),
    "Desabastecimento Crítico": ("#FFA6A6", "#F8CBAD", "#C00000"),
    "Estoque Excessivo":        ("#B2EBF2", "#E5F1F4", "#006666"),
    "Estoque Parado":           ("#F2F2F2", "#F2F2F2", "#595959"),
    "Estoque em Alerta":        ("#FFEB3B", "#FFF2CC", "#7F6000"),
    "Sem Consumo":              ("#CFD8DC", "#F2F2F2", "#595959"),
}

STATUS_CORES = {k: v[0] for k, v in MAPA_STATUS.items()}
EXCEL_CORES  = {k: (v[1], v[2]) for k, v in MAPA_STATUS.items()}

CATEGORIAS_PADRAO = sorted([
    "MEDICAMENTO", "MMH", "SORO", "NUTRIÇÃO", "GASES MEDICINAIS",
    "MATERIAL DIAGNÓSTICO", "OUTROS",
])

# Caminho do arquivo físico permanente na mesma pasta do app.py
ARQUIVO_CATEGORIAS = Path(__file__).parent / "Categorias_base.xlsx"

# =============================================================================
# FUNÇÕES UTILITÁRIAS E LEITORES COM CACHE
# =============================================================================

@st.cache_data(show_spinner=False)
def ler_csv_cached(file_bytes: bytes, nome: str) -> pd.DataFrame:
    return pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python', encoding='latin1', index_col=False)


@st.cache_data(show_spinner=False)
def ler_xlsx_cached(file_bytes: bytes, nome: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), dtype=str)


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
            f"❌ **Erro no arquivo '{contexto}':** Colunas não identificadas automaticamente: "
            f"`{'`, `'.join(faltando)}`. Verifique os cabeçalhos."
        )
        return False
    return True


def exportar_excel_padronizado(df_dados: pd.DataFrame, nome_aba: str = "Dados") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
        df_dados.to_excel(wr, sheet_name=nome_aba, index=False)
        wb = wr.book
        ws = wr.sheets[nome_aba]

        fmt_base = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_texto = wb.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#1E3A8A', 'font_color': '#FFFFFF', 'border': 1, 'font_size': 11})

        for ci, cn in enumerate(df_dados.columns):
            ws.write(0, ci, cn, fmt_header)
        ws.set_row(0, 40)

        col_mv = find_col(df_dados, ['codigo', 'mv', 'id']) or df_dados.columns[0]
        col_mat = find_col(df_dados, ['material', 'produto', 'descri']) or df_dados.columns[1]
        col_cat = find_col(df_dados, ['categoria', 'grupo']) or df_dados.columns[2]
        col_alerta = find_col(df_dados, ['parecer', 'alerta', 'status'])

        LARGURAS = {
            col_mv: 12, col_mat: 45, col_cat: 16,
            'Saldo Atual Satélite': 20, 'Consumo Médio Diário': 20, 'Estoque Mínimo': 16,
            'Saldo Almox. Centrais Unificado': 28, 'Ação Logística Sugerida': 50
        }
        for ci, cn in enumerate(df_dados.columns):
            larg = LARGURAS.get(cn, 24)
            fmt_col = fmt_texto if cn in (col_mat, 'Ação Logística Sugerida') else fmt_base
            ws.set_column(ci, ci, larg, fmt_col)

        total_linhas = len(df_dados)
        for ri in range(1, total_linhas + 1):
            ws.set_row(ri, 30)
        ws.freeze_panes(1, 0)

        if col_alerta:
            idx_parecer = df_dados.columns.get_loc(col_alerta)
            letra_col = chr(ord('A') + idx_parecer)
            n_cols = len(df_dados.columns) - 1
            
            for status, (bg, fg) in EXCEL_CORES.items():
                fmt_cond = wb.add_format({'bg_color': bg, 'font_color': fg, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
                ws.conditional_format(1, 0, total_linhas, n_cols, {'type': 'formula', 'criteria': f'=${letra_col}2="{status}"', 'format': fmt_cond})
    return buf.getvalue()

# =============================================================================
# PERSISTÊNCIA EM DISCO COMPATÍVEL COM CLOUD (PLANILHA2)
# =============================================================================

def carregar_categorias_do_disco() -> pd.DataFrame:
    if ARQUIVO_CATEGORIAS.exists():
        try:
            excel_file = pd.ExcelFile(ARQUIVO_CATEGORIAS)
            abas_ordenadas = sorted(excel_file.sheet_names, key=lambda x: '2' in x or 'plan' in x.lower(), reverse=True)
            
            for aba in abas_ordenadas:
                df = excel_file.parse(aba, dtype=str)
                if df.empty or len(df) < 1: 
                    continue
                
                c_cod = find_col(df, ['codigo', 'cod', 'ca3', 'id'])
                c_mat = find_col(df, ['material', 'produto', 'insumo', 'descri', 'nome'])
                c_cat = find_col(df, ['categoria', 'grupo', 'classe', 'tipo'])
                
                if c_cod and c_cat:
                    df_clean = pd.DataFrame()
                    df_clean["Código"] = df[c_cod].apply(clean_key)
                    df_clean["Material"] = df[c_mat].fillna("").astype(str).str.strip() if c_mat else ""
                    df_clean["Categoria"] = df[c_cat].str.upper().str.strip().fillna("OUTROS")
                    
                    df_clean = df_clean[df_clean["Código"] != ""].drop_duplicates("Código")
                    if not df_clean.empty:
                        return df_clean.reset_index(drop=True)
        except Exception:
            pass
    
    df_vazio = pd.DataFrame(columns=["Código", "Material", "Categoria"])
    try:
        df_vazio.to_excel(ARQUIVO_CATEGORIAS, index=False)
    except Exception:
        pass
    return df_vazio


def salvar_categorias_no_disco(df: pd.DataFrame) -> bool:
    try:
        df.to_excel(ARQUIVO_CATEGORIAS, sheet_name="Planilha1", index=False)
        return True
    except Exception as e:
        st.error(f"❌ Erro ao salvar Categorias_base.xlsx no disco: {e}")
        return False


def inicializar_categorias_session():
    if "df_categorias" not in st.session_state:
        st.session_state["df_categorias"] = carregar_categorias_do_disco()


def obter_mapa_categorias() -> dict:
    df = st.session_state.get("df_categorias", pd.DataFrame())
    if df.empty:
        return {}
    return dict(zip(df["Código"].astype(str), df["Categoria"].astype(str)))


def enriquecer_e_auto_preencher_categorias(df_est: pd.DataFrame, c_est_cod: str, c_est_prod: str):
    df_cat = st.session_state.get("df_categorias", pd.DataFrame()).copy()
    
    df_est_unicos = df_est.drop_duplicates(subset=["key"]).copy()
    mapa_descricoes_reais = dict(zip(df_est_unicos["key"].astype(str), df_est_unicos[c_est_prod].fillna("").astype(str).str.strip()))
    
    def reparar_descricao(row):
        cod_str = str(row["Código"])
        desc_atual = str(row["Material"]).strip()
        if desc_atual == "" or desc_atual.upper() == "PRODUTO SEM DESCRIÇÃO":
            return mapa_descricoes_reais.get(cod_str, desc_atual)
        return desc_atual

    if not df_cat.empty:
        df_cat["Material"] = df_cat.apply(reparar_descricao, axis=1)

    codigos_existentes = set(df_cat["Código"].astype(str).tolist())
    df_novos_raw = df_est_unicos[~df_est_unicos["key"].astype(str).isin(codigos_existentes)]

    if not df_novos_raw.empty:
        df_novos = pd.DataFrame({
            "Código":    df_novos_raw["key"].astype(str).values,
            "Material":  df_novos_raw[c_est_prod].fillna("").astype(str).str.strip().values,
            "Categoria": "OUTROS"
        })
        df_cat = pd.concat([df_cat, df_novos], ignore_index=True)
    
    df_cat = df_cat.drop_duplicates("Código", keep="first").reset_index(drop=True)
    st.session_state["df_categorias"] = df_cat
    salvar_categorias_no_disco(df_cat)

# =============================================================================
# LÓGICA DE NEGÓCIO E REGRAS ATUALIZADAS
# =============================================================================

def calcular_cmd(qtd_total: float, dias: int) -> float:
    cmd_bruto = qtd_total / max(dias, 1)
    return float(math.ceil(cmd_bruto)) if cmd_bruto > 0 else 0.0


def calcular_sugestao(row: pd.Series, dias_pedido: int) -> int:
    cmd = row['Consumo Médio Diário']
    est_atual = row['Saldo Atual Satélite']
    est_minimo = row['Estoque Mínimo']
    
    if cmd == 0 and est_minimo > 0:
        return max(0, round(est_minimo - est_atual))
    if cmd == 0:
        return 0
        
    meta_consumo = cmd * dias_pedido
    meta_final = max(meta_consumo, est_minimo)
    return max(0, round(meta_final - est_atual))


def definir_alerta_e_acao(row: pd.Series, dict_saldos_centrais: dict, dict_saldos_parceiras: dict, consumo_outras_total: dict) -> tuple:
    cod = row['Código MV']
    sug = row['Necessidade de Ressuprimento']
    cmd = row['Consumo Médio Diário']
    est_un = row['Saldo Atual Satélite']
    est_minimo = row['Estoque Mínimo']

    if cmd == 0 and est_minimo <= 0 and est_un == 0:
        return "Sem Consumo", "Avaliar se é necessário inativar o item na farmácia."
    if cmd > 0 and est_un > (cmd * 60):
        return "Estoque Excessivo", "Estoque acima da necessidade de 60 dias. Devolver!."
    if cmd == 0 and est_un > 0:
        if est_un <= est_minimo:
            return "Estoque Parado", "Item sem consumo, mas dentro do estoque mínimo parametrizado."
        else:
            excedente = int(est_un - est_minimo)
            return "Estoque Parado", f"{excedente} unidades acima do estoque mínimo. Considerar devolver ou remanejar para outra farmácia."
            
    # Variável de prefixo: adiciona contexto visual de urgência quando aplicável
    prefixo_alerta = ""
    if sug > 0 and est_un < est_minimo:
        prefixo_alerta = "[ALERTA - ABAIXO DO MÍNIMO] "
        
    if sug <= 0:
        return "Estoque Suficiente", "Estoque dentro da cobertura ideal."

    # Processa disponibilidade nas parceiras primeiro, para ter a informação pronta
    saldos_parceiras = dict_saldos_parceiras.get(cod, {})
    consumos_parceiras = consumo_outras_total.get(cod, {})

    farmacias_com_estoque_parado = []
    for farm_id, saldo_f in saldos_parceiras.items():
        if saldo_f > 0:
            c_parc = consumos_parceiras.get(farm_id, 0)
            if c_parc == 0 or saldo_f > (c_parc * 3):
                nome_pratico = DIC_NOMES_FARMACIAS.get(str(farm_id), "Farmácia Satélite")
                farmacias_com_estoque_parado.append(f"Cód {farm_id} ({nome_pratico} - {int(saldo_f)} un.)")
    
    locais_remanejo = " | ".join(farmacias_com_estoque_parado)

    # Processa disponibilidade nas Centrais
    saldos_nas_centrais = dict_saldos_centrais.get(cod, {})
    saldo_total_central = sum(saldos_nas_centrais.values())

    # Se a central possui algo, avalia se é total ou parcial
    if saldo_total_central > 0:
        central_principal = max(saldos_nas_centrais, key=saldos_nas_centrais.get)
        if saldo_total_central >= sug:
            return "Solicitar", f"{prefixo_alerta}Solicitar {int(sug)} un. ao Almoxarifado Central {central_principal}."
        else:
            faltante = int(sug - saldo_total_central)
            if farmacias_com_estoque_parado:
                # O cenário corrigido: tem na central parcialmente, e TEM parceira para cobrir o resto
                return f"Estoque Crítico no Almoxarifado {central_principal}", f"{prefixo_alerta}Pegar {int(saldo_total_central)} un. no Almox Central {central_principal} e remanejar o restante ({faltante} un.) de: {locais_remanejo}."
            else:
                # O cenário corrigido: tem na central parcialmente, mas NÃO TEM parceira para cobrir o resto
                return f"Estoque Crítico no Almoxarifado {central_principal}", f"{prefixo_alerta}Pegar {int(saldo_total_central)} un. no Almox Central {central_principal}. ALERTA: Sem saldo nas farmácias parceiras para cobrir as {faltante} un. restantes."

    # Se a central está zerada, verifica se tem como remanejar
    if farmacias_com_estoque_parado:
        return "Remanejar", f"{prefixo_alerta}Central Zerada! Transferir de: {locais_remanejo}."

    # Se não houver nem na central, nem nas parceiras, é Ruptura.
    return "Desabastecimento Crítico", f"{prefixo_alerta}Sem saldo nos almoxarifados e sem estoque parado nas farmácias."

# Inicializa banco de dados local unificado
inicializar_categorias_session()

# =============================================================================
# SIDEBAR COM MANUAL E CRÉDITOS EXCLUSIVOS
# =============================================================================
with st.sidebar:
    st.markdown("### 🏥 Parâmetros")
    farmacias_opcoes = {
        "Farmácia UMI (Cód. 13)": "13",
        "Farmácia Dutra (Cód. 31)": "31",
        "Farmácia Centro Cirúrgico (Cód. 7)": "7",
        "Farmácia Oftalmologia (Cód. 39)": "39",
        "Farmácia UTI (Cód. 34)": "34",
    }
    farmacia_selecionada = st.selectbox("Defina a Farmácia:", list(farmacias_opcoes.keys()))
    cod_farmacia_alvo = farmacias_opcoes[farmacia_selecionada]

    st.write("---")
    st.markdown("### ⚙️ Parâmetros do Pedido")
    dias_pedido = st.number_input("Defina quantos dias de ressuprimento será solicitado:", value=15, min_value=1)

    ontem = datetime.now() - timedelta(days=1)
    data_inicio = st.date_input("Início do Histórico de Consumo:", value=ontem - timedelta(days=7), format="DD/MM/YYYY")
    data_fim = st.date_input("Fim do Período de Consumo:", value=ontem, format="DD/MM/YYYY")

    st.write("---")
    with st.expander("📄 Manual de Regras Logísticas", expanded=False):
        st.markdown("""
        ### Matriz de Alertas e Ações Sugeridas
        
        #### 1. ⚪ Sem Consumo
        * **Condição:** Consumo Médio Diário ($CMD$) = 0, Estoque Mínimo = 0 e Saldo Satélite = 0.
        * **Ação Logística:** *\"Avaliar se é necessário inativar o item na farmácia.\"*
        
        #### 2. 🔵 Estoque Excessivo
        * **Condição:** Item ativo ($CMD > 0$) com saldo cobrindo mais de 60 dias de consumo.
        * **Ação Logística:** *\"Estoque acima da necessidade de 60 dias. Devolver!.\"*
        
        #### 3. ⚫ Estoque Parado
        * **Condição:** Giro zerado ($CMD = 0$), mas possui saldo físico na farmácia satélite.
        * **Ação Logística:** Avisa o volume em excesso acima do estoque mínimo e sugere remanejamento externo.
        
        #### 4. 🟡 Estoque em Alerta (Integrado ao Fluxo)
        * **Condição:** A farmácia precisa de material ($Sugestão > 0$) e o saldo atual caiu abaixo do estoque mínimo de segurança.
        * **Ação Logística:** O item é encaminhado normalmente para a lista de **Solicitação**, **Remanejamento** ou **Ruptura** (dependendo da disponibilidade da rede), mas recebe a marcação urgente **[ALERTA - ABAIXO DO MÍNIMO]** na descrição da sua ação sugerida.
        
        #### 5. 🟢 Estoque Suficiente
        * **Condição:** O saldo cobre a janela histórica e está acima da margem de segurança.
        
        #### 6. 🔵 Solicitar Abastecimento Central
        * **Condição:** Pedido $> 0$ e o estoque das **Centrais Unificadas (1, 6, 9, 41, 43)** atende a demanda.
        * **Ação Logística:** Aponta a central ideal detentora do maior saldo para separação.
        
        #### 7. 🟠 Estoque Crítico no Almoxarifado X
        * **Condição:** Pedido $> 0$, mas a central específica de número X possui saldo parcial e insuficiente.
        * **Ação Logística:** Orienta raspar a central X e informa exatamente de quais farmácias satélites capturar o saldo restante via remanejamento.
        
        #### 8. 🟡 Remanejar entre Farmácias
        * **Condição:** Centrais zeradas, mas existem outras farmácias satélites com saldo parado ou excedente.
        * **Ação Logística:** Traduz o ID e aponta o nome do setor (ex: *Farmácia UMI*).
        
        #### 9. 🔴 Desabastecimento Crítico
        * **Condição:** Sem saldo nos almoxarifados e sem estoque parado nas farmácias. Ruptura total.
        """)

    with st.expander("🎖️ Créditos do Sistema", expanded=False):
        st.markdown("""
        **Idealização e Desenvolvimento:**
        * Elton Jonh Freitas Santos
        * Farmacêutico - Chefe da UDIS/HUUFMA
        *HUUFMA — Gestão e Inteligência Logística Avançada © 2026*
        """)

# =============================================================================
# INTERFACE WEB PRINCIPAL
# =============================================================================
st.title("🏥 Gestor do Estoque - Unidade de Dispensação Farmacêutica")
st.markdown(
    f"**Farmácia Ativa:** `{farmacia_selecionada}` | "
    f"**Janela Histórica:** `{data_inicio.strftime('%d/%m/%Y')}` até `{data_fim.strftime('%d/%m/%Y')}`"
)
st.write("")

tab1, tab2 = st.tabs(["⚡ Processar Pedido com IA Logística", "🗂️ Gestão de Categorias de Insumos"])

# =============================================================================
# TAB 2 — GERENCIADOR INTEGRADO DE CATEGORIAS
# =============================================================================
with tab2:
    st.subheader("🗂️ Mapeamento Global de Categorias")
    st.info(
        "As alterações feitas nesta tabela são salvas de forma **permanente** no computador. "
        "Para evitar lentidão e poluição visual, a tabela permanece oculta e **abre automaticamente assim que você realizar uma pesquisa por nome/código ou escolher um grupo específico** nos filtros abaixo."
    )

    df_cat_atual = st.session_state["df_categorias"].copy()

    lista_grupos_reais = sorted(list(df_cat_atual["Categoria"].unique())) if not df_cat_atual.empty else CATEGORIAS_PADRAO
    if "OUTROS" not in lista_grupos_reais:
        lista_grupos_reais.append("OUTROS")

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("📦 Total de Insumos Cadastrados", len(df_cat_atual))
    mc2.metric("🏷️ Categorias Mapeadas", df_cat_atual["Categoria"].nunique() if not df_cat_atual.empty else 0)
    mc3.metric("❓ Sem Classificação (OUTROS)", len(df_cat_atual[df_cat_atual["Categoria"] == "OUTROS"]) if not df_cat_atual.empty else 0)

    st.write("---")
    
    fc1, fc2 = st.columns([3, 1])
    filtro_termo_cat = fc1.text_input("🔍 Pesquisar Insumo por nome ou código para alteração:", value="")
    filtro_sel_cat = fc2.selectbox("Filtrar por Grupo Correspondente:", ["TODOS"] + lista_grupos_reais)

    alguem_pesquisou = (filtro_termo_cat.strip() != "") or (filtro_sel_cat != "TODOS")

    if alguem_pesquisou:
        df_filtrado_cat = df_cat_atual.copy()
        if filtro_termo_cat:
            df_filtrado_cat = df_filtrado_cat[
                (df_filtrado_cat["Código"].astype(str).str.contains(filtro_termo_cat, case=False, na=False)) |
                (df_filtrado_cat["Material"].astype(str).str.contains(filtro_termo_cat, case=False, na=False))
            ]
        if filtro_sel_cat != "TODOS":
            df_filtrado_cat = df_filtrado_cat[df_filtrado_cat["Categoria"] == filtro_sel_cat]

        st.markdown(f"##### 📋 Itens Encontrados ({len(df_filtrado_cat)} registros)")

        df_editor_output = st.data_editor(
            df_filtrado_cat.reset_index(drop=True),
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Código": st.column_config.TextColumn("Código MV", required=True, width="small"),
                "Material": st.column_config.TextColumn("Descrição Completa do Insumo", required=True, width="large"),
                "Categoria": st.column_config.SelectboxColumn("Categoria Logística", options=lista_grupos_reais, required=True, width="medium")
            },
            key="editor_categorias"
        )

        col_btn_salvar, col_btn_reset = st.columns([4, 1])
        if col_btn_salvar.button("💾 SALVAR ALTERAÇÕES DE FORMA PERMANENTE NO DISCO", use_container_width=True):
            codigos_visíveis = set(df_filtrado_cat["Código"].astype(str).tolist())
            df_base_original = st.session_state["df_categorias"].copy()
            
            df_base_limpo = df_base_original[~df_base_original["Código"].astype(str).isin(codigos_visíveis)]
            df_novo_consolidado = pd.concat([df_base_limpo, df_editor_output], ignore_index=True).drop_duplicates("Código", keep="last")
            df_novo_consolidado = df_novo_consolidado[df_novo_consolidado["Código"].astype(str).str.strip() != ""]
            
            st.session_state["df_categorias"] = df_novo_consolidado.reset_index(drop=True)
            if salvar_categorias_no_disco(st.session_state["df_categorias"]):
                st.success("✅ Sucesso! Todas as alterações foram salvas fisicamente em 'Categorias_base.xlsx'.")
                st.rerun()

        if col_btn_reset.button("🔄 Cancelar Edições", use_container_width=True):
            st.session_state.pop("df_categorias", None)
            st.rerun()
    else:
        st.warning("🔍 Para visualizar, incluir ou editar os registros das categorias, utilize os campos de pesquisa ou escolha um grupo acima.")

# =============================================================================
# TAB 1 — PROCESSAR PEDIDO DE COMPRAS / SOLICITAÇÃO
# =============================================================================
with tab1:
    with st.container(border=True):
        st.markdown("##### 📥 Upload das Fontes de Dados Obrigatórias (AGHU)")
        col1, col2 = st.columns(2)
        file_mov_alvo = col1.file_uploader("1. Movimento da Farmácia Alvo (.csv)", type=["csv"])
        file_est_geral = col2.file_uploader("2. Estoque Geral de todos os Almoxarifados (.csv)", type=["csv"])
        st.write("")
        files_mov_parceiras = st.file_uploader(
            "3. Movimentos das outras Farmácias, ativa a INTELIGÊNCIA LOGÍSTICA (Múltiplos .csv)", type=["csv"], accept_multiple_files=True
        )

    st.write("")

    if file_mov_alvo and file_est_geral:
        if st.button("🚀 ANALISAR OS DADOS COM INTELIGÊNCIA LOGÍSTICA", use_container_width=True):
            st.session_state['disparar_processamento_huufma'] = True

        if st.session_state.get('disparar_processamento_huufma', False):
            progress = st.progress(0, text="📂 Lendo arquivos...")

            try:
                mov = ler_csv_cached(file_mov_alvo.read(), file_mov_alvo.name)
                est_geral = ler_csv_cached(file_est_geral.read(), file_est_geral.name)

                progress.progress(10, text="🔍 Identificando colunas...")

                cols_mov = {
                    'código':      find_col(mov, ['material', 'cod', 'ca3']),
                    'quantidade':  find_col(mov, ['quant']),
                    'tipo':        find_col(mov, ['tipo']),
                    'data':        find_col(mov, ['data', 'ger']),
                    'almoxarifado':find_col(mov, ['almox']),
                }
                if not validar_colunas(cols_mov, file_mov_alvo.name): st.stop()

                cols_est = {
                    'código':      find_col(est_geral, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod']),
                    'quantidade':  find_col(est_geral, ['qtde disp', 'disponivel']),
                    'produto':     find_col(est_geral, ['material', 'produto', 'descri']),
                    'almoxarifado':find_col(est_geral, ['almox']),
                    'mínimo':      find_col(est_geral, ['qtde estq min', 'estoque minimo', 'minimo']),
                }
                cols_est_criticas = {k: v for k, v in cols_est.items() if k != 'mínimo'}
                if not validar_colunas(cols_est_criticas, file_est_geral.name): st.stop()

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

                est_geral = est_geral.copy()
                est_geral['key']         = est_geral[c_est_cod].apply(clean_key)
                est_geral['almox_limpo'] = est_geral[c_est_almox].apply(clean_key)
                est_geral['saldo_num']   = est_geral[c_est_qtd].apply(p_num)
                est_geral['min_num']     = est_geral[c_est_min].apply(p_num) if c_est_min else 0.0

                def saldo_por_almox(filtro_almox):
                    return est_geral[est_geral['almox_limpo'] == filtro_almox].groupby('key')['saldo_num'].sum().to_dict()

                est_farmacia_alvo = saldo_por_almox(cod_farmacia_alvo)
                est_min_alvo = est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo].groupby('key')['min_num'].sum().to_dict()

                centrais_alvo = ['1', '6', '9', '41', '43']
                est_centrais_filtrado = est_geral[est_geral['almox_limpo'].isin(centrais_alvo)].copy()
                
                dict_saldos_centrais = {}
                if not est_centrais_filtrado.empty:
                    dict_saldos_centrais = (
                        est_centrais_filtrado.groupby('key')
                        .apply(lambda g: dict(zip(g['almox_limpo'], g['saldo_num'])), include_groups=False)
                        .to_dict()
                    )

                cod_parceiras = [c for c in ['7', '13', '31', '34', '39'] if c != cod_farmacia_alvo]
                est_outras = est_geral[est_geral['almox_limpo'].isin(cod_parceiras)].copy()

                dict_saldos_parceiras = {}
                if not est_outras.empty:
                    dict_saldos_parceiras = (
                        est_outras.groupby('key')
                        .apply(lambda g: dict(zip(g['almox_limpo'], g['saldo_num'])), include_groups=False)
                        .to_dict()
                    )

                progress.progress(40, text="📊 Calculando consumo e auditando calendário...")

                mov = mov.copy()
                mov['dt_formatada'] = pd.to_datetime(mov[c_mov_data], dayfirst=True, errors='coerce')
                mov_filtrado = mov[
                    (mov['dt_formatada'].dt.date >= data_inicio) &
                    (mov['dt_formatada'].dt.date <= data_fim) &
                    (mov[c_mov_tipo].astype(str).str.upper() == 'RM')
                ].copy()

                # Auditoria de calendário preventivo para detectar dias zerados
                dias_ideais = pd.date_range(start=data_inicio, end=data_fim).date
                dias_com_movimento = mov_filtrado['dt_formatada'].dt.date.dropna().unique()
                dias_vazios = [d for d in dias_ideais if d not in dias_com_movimento]
                st.session_state['datas_sem_movimento_huufma'] = [d.strftime('%d/%m/%Y') for d in dias_vazios]

                dias_considerados = max((data_fim - data_inicio).days + 1, 1)

                consumo = mov_filtrado.copy().assign(
                    qtd_num=lambda df: df[c_mov_qtd].apply(p_num),
                    key=lambda df: df[c_mov_cod].apply(clean_key)
                ).groupby('key')['qtd_num'].sum().reset_index().rename(columns={'qtd_num': 'total_consumo'})
                
                consumo['cmd'] = consumo['total_consumo'].apply(lambda x: calcular_cmd(x, dias_considerados))

                progress.progress(55, text="🔄 Cruzando dados entre farmácias...")

                consumo_outras_total = {}
                if files_mov_parceiras:
                    for f_parc in files_mov_parceiras:
                        try:
                            df_p = ler_csv_cached(f_parc.read(), f_parc.name)
                            c_p_cod = find_col(df_p, ['material', 'cod', 'ca3'])
                            c_p_qtd = find_col(df_p, ['quant'])
                            c_p_tipo = find_col(df_p, ['tipo'])
                            c_p_almox = find_col(df_p, ['almox'])
                            c_p_data = find_col(df_p, ['data', 'ger'])

                            if not all([c_p_cod, c_p_qtd, c_p_tipo, c_p_almox, c_p_data]): continue

                            df_p = df_p.copy()
                            df_p['dt_formatada'] = pd.to_datetime(df_p[c_p_data], dayfirst=True, errors='coerce')
                            df_p_filt = df_p[
                                (df_p['dt_formatada'].dt.date >= data_inicio) &
                                (df_p['dt_formatada'].dt.date <= data_fim) &
                                (df_p[c_p_tipo].astype(str).str.upper() == 'RM')
                            ].copy()

                            df_p_filt['key'] = df_p_filt[c_p_cod].apply(clean_key)
                            df_p_filt['almox_limpo'] = df_p_filt[c_p_almox].apply(clean_key)
                            df_p_filt['qtd_num'] = df_p_filt[c_p_qtd].apply(p_num)

                            resumo_p_dict = (
                                df_p_filt.groupby('key')
                                .apply(lambda g: dict(zip(g['almox_limpo'], g['qtd_num'])), include_groups=False)
                                .to_dict()
                            )
                            for key_med, almox_dict in resumo_p_dict.items():
                                if key_med not in consumo_outras_total:
                                    consumo_outras_total[key_med] = {}
                                for a_id, q_val in almox_dict.items():
                                    consumo_outras_total[key_med][a_id] = consumo_outras_total[key_med].get(a_id, 0) + q_val
                        except Exception: continue

                progress.progress(70, text="🧠 Gerando DataFrames Finais...")

                mapa_produtos = est_geral.drop_duplicates(subset=['key']).set_index('key')[c_est_prod].to_dict()
                consumo_map = consumo.set_index('key')['cmd'].to_dict()

                todos_codigos = sorted(list(set(est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo]['key'].unique()) | set(consumo['key'])))

                final = pd.DataFrame({'Código MV': todos_codigos})
                final['Material'] = final['Código MV'].map(mapa_produtos).fillna('PRODUTO SEM DESCRIÇÃO')
                final['Saldo Atual Satélite'] = final['Código MV'].map(est_farmacia_alvo).fillna(0)
                final['Consumo Médio Diário'] = final['Código MV'].map(consumo_map).fillna(0)
                final['Estoque Mínimo'] = final['Código MV'].map(est_min_alvo).fillna(0)
                final['Necessidade de Ressuprimento'] = final.apply(lambda row: calcular_sugestao(row, dias_pedido), axis=1)

                enriquecer_e_auto_preencher_categorias(est_geral, c_est_cod, c_est_prod)

                mapa_cat = obter_mapa_categorias()
                final['Categoria'] = final['Código MV'].map(mapa_cat).fillna('OUTROS')
                
                mapa_descricoes_reparadas = dict(zip(st.session_state["df_categorias"]["Código"].astype(str), st.session_state["df_categorias"]["Material"].astype(str)))
                final['Material'] = final['Código MV'].map(mapa_descricoes_reparadas).fillna(final['Material'])

                final['Saldo Almox. Centrais Unificado'] = final['Código MV'].apply(lambda c: sum(dict_saldos_centrais.get(c, {}).values()))

                alertas_acoes = final.apply(lambda row: definir_alerta_e_acao(row, dict_saldos_centrais, dict_saldos_parceiras, consumo_outras_total), axis=1)
                final['Parecer Logístico / Alerta'] = [r[0] for r in alertas_acoes]
                final['Ação Logística Sugerida'] = [r[1] for r in alertas_acoes]

                st.session_state['df_final_huufma'] = final
                st.session_state['disparar_processamento_huufma'] = False
                progress.progress(100, text="✅ Processamento concluído!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erro crítico no processamento: {e}")
                st.session_state['disparar_processamento_huufma'] = False

        # --- SEÇÃO DE RENDERIZAÇÃO PAINEL GESTOR ---
        if 'df_final_huufma' in st.session_state:
            df_view = st.session_state['df_final_huufma'].copy()

            datas_vazias_detectadas = st.session_state.get('datas_sem_movimento_huufma', [])
            if datas_vazias_detectadas:
                st.warning(
                    f"⚠️ **Atenção Logística (Auditoria de Calendário):** Não foi encontrada nenhuma movimentação "
                    f"no arquivo de saídas para as seguintes datas do período selecionado: `{', '.join(datas_vazias_detectadas)}`. "
                    f"Verifique se o arquivo enviado está completo para evitar subdimensionar a meta do pedido."
                )

            st.write("---")
            df_desabast = df_view[df_view['Parecer Logístico / Alerta'] == "Desabastecimento Crítico"].sort_values('Material')
            df_remanej = df_view[df_view['Parecer Logístico / Alerta'] == "Remanejar"].sort_values('Material')
            
            df_caf_disp = df_view[df_view['Parecer Logístico / Alerta'].str.contains("Solicitar|Almoxarifado", na=False)].sort_values('Material')
            
            df_excesso = df_view[
                (df_view['Parecer Logístico / Alerta'].isin(["Estoque Excessivo", "Estoque Parado", "Sem Consumo"])) &
                (df_view['Parecer Logístico / Alerta'] != "Sem Consumo")
            ].sort_values('Material')

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("🚨 Desabastecimento Crítico", f"{len(df_desabast)} itens")
                st.download_button("📥 Extrair Lista", data=exportar_excel_padronizado(df_desabast, "Rupturas"), file_name=f"Rupturas_{cod_farmacia_alvo}.xlsx", key="ex_c1", use_container_width=True)
            with c2:
                st.metric("🔄 Remanejamento Potencial", f"{len(df_remanej)} itens")
                st.download_button("📥 Extrair Lista", data=exportar_excel_padronizado(df_remanej, "Remanejamento"), file_name=f"Remanejamento_{cod_farmacia_alvo}.xlsx", key="ex_c2", use_container_width=True)
            with c3:
                st.metric("📦 Disponível no Almoxarifado", f"{len(df_caf_disp)} itens")
                st.download_button("📥 Extrair Lista", data=exportar_excel_padronizado(df_caf_disp, "Disponiveis_CAF"), file_name=f"Disponiveis_Centrais_{cod_farmacia_alvo}.xlsx", key="ex_c3", use_container_width=True)
            with c4:
                st.metric("⚠️Atenção! (Excesso/Sem Giro)", f"{len(df_excesso)} itens")
                st.download_button("📥 Extrair Lista", data=exportar_excel_padronizado(df_excesso, "Riscos_Excesso"), file_name=f"Overstock_{cod_farmacia_alvo}.xlsx", key="ex_c4", use_container_width=True)

            st.write("")
            g1, g2 = st.columns(2)

            with g1:
                st.markdown("**Saúde Geral do Estoque (% Proporcional)**")
                df_g1 = df_view.copy()
                df_g1.loc[df_g1['Parecer Logístico / Alerta'].str.contains("Almoxarifado", na=False), 'Parecer Logístico / Alerta'] = "Estoque Crítico CAF"
                df_g1_grouped = df_g1.groupby('Parecer Logístico / Alerta')['Código MV'].count().reset_index().rename(columns={'Código MV': 'Quantidade', 'Parecer Logístico / Alerta': 'Status'})
                df_g1_grouped = df_g1_grouped[df_g1_grouped['Status'] != 'Sem Consumo'] 

                MAPA_CORES_GRAFICO = STATUS_CORES.copy()
                MAPA_CORES_GRAFICO["Estoque Crítico CAF"] = "#FFC499"

                if not df_g1_grouped.empty:
                    donut_chart = alt.Chart(df_g1_grouped).mark_arc(innerRadius=65, stroke='#fff').encode(
                        theta=alt.Theta(field="Quantidade", type="quantitative"),
                        color=alt.Color(field="Status", type="nominal", scale=alt.Scale(domain=list(MAPA_CORES_GRAFICO.keys()), range=list(MAPA_CORES_GRAFICO.values())), legend=alt.Legend(title="Parecer")),
                        tooltip=['Status', 'Quantidade']
                    ).properties(height=280)
                    st.altair_chart(donut_chart, use_container_width=True)
                else:
                    st.info("Sem dados suficientes para gerar a rosca estatística.")

            with g2:
                st.markdown("**Urgência do Ressuprimento por Categoria**")
                df_g2 = df_view[df_view['Categoria'] != 'OUTROS'].copy()
                df_g2.loc[df_g2['Parecer Logístico / Alerta'].str.contains("Almoxarifado", na=False), 'Parecer Logístico / Alerta'] = "Estoque Crítico CAF"
                df_g2_grouped = df_g2.groupby(['Categoria', 'Parecer Logístico / Alerta'])['Código MV'].count().reset_index().rename(columns={'Código MV': 'Itens', 'Parecer Logístico / Alerta': 'Parecer'})
                df_g2_grouped = df_g2_grouped[df_g2_grouped['Parecer'] != 'Sem Consumo'].copy()

                if not df_g2_grouped.empty:
                    stacked_chart = alt.Chart(df_g2_grouped).mark_bar().encode(
                        x=alt.X('Itens:Q', title="Quantidade de Insumos"),
                        y=alt.Y('Categoria:N', title=None, sort='-x'),
                        color=alt.Color('Parecer:N', scale=alt.Scale(domain=list(MAPA_CORES_GRAFICO.keys()), range=list(MAPA_CORES_GRAFICO.values())), legend=None),
                        tooltip=['Categoria', 'Parecer', 'Itens']
                    ).properties(height=280)
                    st.altair_chart(stacked_chart, use_container_width=True)
                else:
                    st.info("Nenhuma categoria vinculada ativa no filtro atual.")

            st.write("---")
            st.markdown("#### 📋 Resultado da Análise Inteligente")
            
            with st.container(border=True):
                st.markdown("##### 🔍 Filtros Dinâmicos")
                f1, f2, f3 = st.columns([2, 2, 1])
                busca_nome = f1.text_input("Filtrar por nome ou código do Insumo:", value="")
                opcoes_alertas_vivos = sorted(list(df_view['Parecer Logístico / Alerta'].unique()))
                busca_alerta = f2.multiselect("Filtrar por Parecer Logístico:", options=opcoes_alertas_vivos)
                busca_cat = f3.selectbox("Filtrar por Categoria:", options=["TODAS"] + list(df_view['Categoria'].unique()))

            if busca_nome:
                df_view = df_view[
                    (df_view['Material'].astype(str).str.contains(busca_nome, case=False, na=False)) |
                    (df_view['Código MV'].astype(str).str.contains(busca_nome, case=False, na=False))
                ]
            if busca_alerta:
                df_view = df_view[df_view['Parecer Logístico / Alerta'].isin(busca_alerta)]
            if busca_cat != "TODAS":
                df_view = df_view[df_view['Categoria'] == busca_cat]

            st.dataframe(
                df_view.sort_values('Material'),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Código MV":                    st.column_config.TextColumn("Código MV", width="small"),
                    "Material":                     st.column_config.TextColumn("Descrição do Insumo / Medicamento", width="large"),
                    "Saldo Atual Satélite":         st.column_config.NumberColumn("Saldo Atual Satélite", format="%d"),
                    "Consumo Médio Diário":         st.column_config.NumberColumn("CMD", format="%d"),
                    "Estoque Mínimo":               st.column_config.NumberColumn("Estq Mínimo", format="%d"),
                    "Necessidade de Ressuprimento": st.column_config.NumberColumn("Necessidade Ressuprimento", format="%d"),
                    "Saldo Almox. Centrais Unificado": st.column_config.NumberColumn("Saldo Centrais", format="%d"),
                    "Parecer Logístico / Alerta":   st.column_config.TextColumn("Parecer Logístico / Alerta", width="medium"),
                }
            )

            st.write("")
            b_p1, b_p2 = st.columns(2)
            
            with b_p1:
                st.download_button(
                    label="📥 BAIXAR RELATÓRIO COMPLETO DA ANÁLISE (.XLSX)",
                    data=exportar_excel_padronizado(st.session_state['df_final_huufma'], "Painel Geral"),
                    file_name=f"Painel_Geral_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )
            with b_p2:
                final_salvar = st.session_state['df_final_huufma'].copy()
                final_salvar = final_salvar.rename(columns={
                    'Necessidade de Ressuprimento': f'Necessidade de Ressuprimento ({dias_pedido} dias)'
                })
                
                ordem_colunas = [
                    'Código MV', 'Material', 'Categoria',
                    'Saldo Atual Satélite', 'Consumo Médio Diário', 'Estoque Mínimo',
                    f'Necessidade de Ressuprimento ({dias_pedido} dias)',
                    'Saldo Almox. Centrais Unificado',
                    'Parecer Logístico / Alerta', 'Ação Logística Sugerida',
                ]
                
                buffer_abas = io.BytesIO()
                
                with pd.ExcelWriter(buffer_abas, engine='xlsxwriter') as writer:
                    for cat, df_cat in final_salvar.groupby('Categoria'):
                        df_exportar = df_cat[df_cat['Ação Logística Sugerida'] != "Avaliar se é necessário inativar o item na farmácia."].copy()
                        if df_exportar.empty: continue
                        
                        df_exportar = df_exportar[ordem_colunas].sort_values('Material')
                        nome_aba = str(cat)[:31]
                        df_exportar.to_excel(writer, sheet_name=nome_aba, index=False)
                        
                        wb_m = writer.book
                        ws_m = writer.sheets[nome_aba]
                        
                        fmt_b = wb_m.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
                        fmt_t = wb_m.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
                        fmt_h = wb_m.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#1E3A8A', 'font_color': '#FFFFFF', 'border': 1, 'font_size': 11})
                        
                        # [CORREÇÃO APLICADA]: ws_m.write substituiu o antigo writer.write
                        for ci, cn in enumerate(df_exportar.columns): ws_m.write(0, ci, cn, fmt_h)
                        ws_m.set_row(0, 40)
                        
                        LARGURAS = {
                            'Código MV': 12, 'Material': 45, 'Categoria': 16, 'Saldo Atual Satélite': 20,
                            'Consumo Médio Diário': 20, 'Estoque Mínimo': 16, f'Necessidade de Ressuprimento ({dias_pedido} dias)': 24,
                            'Saldo Almox. Centrais Unificado': 28, 'Parecer Logístico / Alerta': 26, 'Ação Logística Sugerida': 50
                        }
                        for ci, cn in enumerate(df_exportar.columns):
                            ws_m.set_column(ci, ci, LARGURAS.get(cn, 20), fmt_t if cn in ('Material', 'Ação Logística Sugerida') else fmt_b)
                        
                        for ri in range(1, len(df_exportar) + 1): ws_m.set_row(ri, 30)
                        ws_m.freeze_panes(1, 0)
                        
                        if 'Parecer Logístico / Alerta' in df_exportar.columns:
                            idx_p = df_exportar.columns.get_loc('Parecer Logístico / Alerta')
                            l_col = chr(ord('A') + idx_p)
                            for status, (bg, fg) in EXCEL_CORES.items():
                                fmt_c = wb_m.add_format({'bg_color': bg, 'font_color': fg, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
                                if status == "Estoque Crítico CAF":
                                    ws_m.conditional_format(1, 0, len(df_exportar), len(df_exportar.columns)-1, {'type': 'cell', 'operator': 'containing', 'value': 'Estoque Crítico no Almox', 'format': fmt_c})
                                else:
                                    ws_m.conditional_format(1, 0, len(df_exportar), len(df_exportar.columns)-1, {'type': 'formula', 'criteria': f'=${l_col}2="{status}"', 'format': fmt_c})
                                
                            # [NOVA DECLARAÇÃO LOCAL DE SEGURANÇA]: Impede NameErrors ao rodar no cloud
                            fmt_critico_caf_local = wb_m.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
                            ws_m.conditional_format(1, 0, len(df_exportar), len(df_exportar.columns)-1, {'type': 'cell', 'operator': 'containing', 'value': 'Almoxarifado', 'format': fmt_critico_caf_local})
                                
                st.download_button(
                    label="📥 GERAR RELATÓRIO DO PEDIDO - ABAS POR CATEGORIAS (.XLSX)",
                    data=buffer_abas.getvalue(),
                    file_name=f"Pedido_Abas_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )

    else:
        st.session_state['disparar_processamento_huufma'] = False
        st.warning("⚠️ Aguardando o upload dos arquivos obrigatórios (Movimento e Estoque Geral) para iniciar o processamento.")
