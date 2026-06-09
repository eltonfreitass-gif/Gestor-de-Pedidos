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
# CONFIGURAÇÃO DA PÁGINA E ESTILOS
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
    div[data-testid="stFileUploaderFileData"] section { gap: 6px !important; }
    div[data-testid="stFileUploaderDropzone"] {
        padding: 1rem !important;
        border-radius: 8px !important;
    }
    </style>
""", unsafe_allow_html=True)

# =============================================================================
# CONSTANTES GLOBAIS
# =============================================================================

DIC_NOMES_FARMACIAS = {
    "7":  "Farmácia Centro Cirúrgico",
    "13": "Farmácia UMI",
    "31": "Farmácia Dutra",
    "34": "Farmácia UTI",
    "39": "Farmácia Oftalmologia",
}

# Fonte única: (cor_grafico, bg_excel, fg_excel)
MAPA_STATUS = {
    "Estoque Suficiente":       ("#A2E8A2", "#E2EFDA", "#375623"),
    "Solicitar":                ("#A6C8FF", "#DDEBF7", "#1F4E78"),
    "Remanejar":                ("#FFEAA6", "#FFF2CC", "#7F6000"),
    "Desabastecimento Crítico": ("#FFA6A6", "#F8CBAD", "#C00000"),
    "Estoque Excessivo":        ("#B2EBF2", "#E5F1F4", "#006666"),
    "Estoque Parado":           ("#F2F2F2", "#F2F2F2", "#595959"),
    "Sem Consumo":              ("#CFD8DC", "#F2F2F2", "#595959"),
}
STATUS_CORES = {k: v[0] for k, v in MAPA_STATUS.items()}
EXCEL_CORES  = {k: (v[1], v[2]) for k, v in MAPA_STATUS.items()}

# Cores do gráfico incluem o status dinâmico "Estoque Crítico CAF"
MAPA_CORES_GRAFICO = {**STATUS_CORES, "Estoque Crítico CAF": "#FFC499"}

CATEGORIAS_PADRAO = sorted([
    "MEDICAMENTO", "MMH", "SORO", "NUTRIÇÃO",
    "GASES MEDICINAIS", "MATERIAL DIAGNÓSTICO", "OUTROS",
])

ARQUIVO_CATEGORIAS = Path(__file__).parent / "Categorias_base.xlsx"

# Janela de tendência de consumo (últimos N dias com movimento)
JANELA_TENDENCIA_DIAS = 3


# =============================================================================
# FUNÇÕES UTILITÁRIAS
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
# EXPORTAÇÃO EXCEL
# =============================================================================

def exportar_excel_padronizado(df_dados: pd.DataFrame, nome_aba: str = "Dados") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
        df_dados.to_excel(wr, sheet_name=nome_aba, index=False)
        wb = wr.book
        ws = wr.sheets[nome_aba]

        # Alinhamento vertical centralizado para melhor UX com altura dinâmica
        fmt_base   = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                    'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_texto  = wb.add_format({'align': 'left',   'valign': 'vcenter', 'text_wrap': True,
                                    'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                                    'text_wrap': True, 'bg_color': '#1E3A8A',
                                    'font_color': '#FFFFFF', 'border': 1, 'font_size': 11})

        for ci, cn in enumerate(df_dados.columns):
            ws.write(0, ci, cn, fmt_header)
        ws.set_row(0, 40)

        col_mv     = find_col(df_dados, ['codigo', 'mv', 'id']) or df_dados.columns[0]
        col_mat    = find_col(df_dados, ['material', 'produto', 'descri']) or df_dados.columns[1]
        col_cat    = find_col(df_dados, ['categoria', 'grupo']) or df_dados.columns[2]
        col_alerta = find_col(df_dados, ['parecer', 'alerta', 'status'])

        LARGURAS_PAD = {
            col_mv: 12, col_mat: 45, col_cat: 16,
            'Saldo Atual Satélite': 20, 'Consumo Médio Diário': 20,
            'Estoque Mínimo': 16, 'Cobertura (dias)': 14,
            'CMD Últ. 3 dias': 16, 'Tendência': 12, 'Δ% Tendência': 12,
            'Saldo Almox. Centrais Unificado': 28, 'Ação Logística Sugerida': 50,
            'PEDIDO (X DIAS)': 20
        }
        for ci, cn in enumerate(df_dados.columns):
            larg    = LARGURAS_PAD.get(cn, 24)
            fmt_col = fmt_texto if cn in (col_mat, 'Ação Logística Sugerida') else fmt_base
            ws.set_column(ci, ci, larg, fmt_col)

        # Sem trava de altura para permitir autoajuste do Excel
        total_linhas = len(df_dados)
        ws.freeze_panes(1, 0)

        if col_alerta:
            idx_al  = df_dados.columns.get_loc(col_alerta)
            letra   = chr(ord('A') + idx_al)
            n_cols  = len(df_dados.columns) - 1
            for status, (bg, fg) in EXCEL_CORES.items():
                fmt_c = wb.add_format({'bg_color': bg, 'font_color': fg, 'align': 'center',
                                       'valign': 'vcenter', 'text_wrap': True, 'bold': True,
                                       'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
                ws.conditional_format(1, 0, total_linhas, n_cols,
                    {'type': 'formula', 'criteria': f'=${letra}2="{status}"', 'format': fmt_c})
            fmt_critico = wb.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911',
                                         'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                         'bold': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
            ws.conditional_format(1, 0, total_linhas, n_cols,
                {'type': 'cell', 'operator': 'containing', 'value': 'Almoxarifado', 'format': fmt_critico})
    return buf.getvalue()


def exportar_excel_multi_aba(df_total: pd.DataFrame, ordem_cols: list,
                              col_categoria: str, col_alerta: str,
                              larguras: dict, excluir_acoes: list = None) -> bytes:
    buf = io.BytesIO()
    excluir_acoes = excluir_acoes or []

    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        wb     = writer.book
        # Alinhamento vertical centralizado
        fmt_b  = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_t  = wb.add_format({'align': 'left',   'valign': 'vcenter', 'text_wrap': True,
                                 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_h  = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                                 'text_wrap': True, 'bg_color': '#1E3A8A',
                                 'font_color': '#FFFFFF', 'border': 1, 'font_size': 11})
        fmt_critico = wb.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911',
                                      'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                      'bold': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})

        for cat, df_cat in df_total.groupby(col_categoria):
            df_exp = df_cat.copy()
            if excluir_acoes and 'Ação Logística Sugerida' in df_exp.columns:
                df_exp = df_exp[~df_exp['Ação Logística Sugerida'].isin(excluir_acoes)]
                
            cols_ok = [c for c in ordem_cols if c in df_exp.columns]
            df_exp  = df_exp[cols_ok].sort_values('Material').copy()
            if df_exp.empty:
                continue

            nome_aba = str(cat)[:31]
            df_exp.to_excel(writer, sheet_name=nome_aba, index=False)
            ws = writer.sheets[nome_aba]

            for ci, cn in enumerate(df_exp.columns):
                ws.write(0, ci, cn, fmt_h)
            ws.set_row(0, 40)

            for ci, cn in enumerate(df_exp.columns):
                larg    = larguras.get(cn, 20)
                fmt_col = fmt_t if cn in ('Material', 'Ação Logística Sugerida') else fmt_b
                ws.set_column(ci, ci, larg, fmt_col)

            total_l = len(df_exp)
            ws.freeze_panes(1, 0)

            if col_alerta in df_exp.columns:
                idx_al = df_exp.columns.get_loc(col_alerta)
                letra  = chr(ord('A') + idx_al)
                n_cols = len(df_exp.columns) - 1
                for status, (bg, fg) in EXCEL_CORES.items():
                    fmt_c = wb.add_format({'bg_color': bg, 'font_color': fg, 'align': 'center',
                                           'valign': 'vcenter', 'text_wrap': True, 'bold': True,
                                           'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
                    ws.conditional_format(1, 0, total_l, n_cols,
                        {'type': 'formula', 'criteria': f'=${letra}2="{status}"', 'format': fmt_c})
                ws.conditional_format(1, 0, total_l, n_cols,
                    {'type': 'cell', 'operator': 'containing',
                     'value': 'Almoxarifado', 'format': fmt_critico})
    return buf.getvalue()


# =============================================================================
# PERSISTÊNCIA DE CATEGORIAS (SOMENTE LEITURA E EDIÇÃO MANUAL)
# =============================================================================

def carregar_categorias_do_disco() -> pd.DataFrame:
    if ARQUIVO_CATEGORIAS.exists():
        try:
            excel_file = pd.ExcelFile(ARQUIVO_CATEGORIAS)
            abas_ordenadas = sorted(
                excel_file.sheet_names,
                key=lambda x: '2' in x or 'plan' in x.lower(), reverse=True
            )
            for aba in abas_ordenadas:
                df = excel_file.parse(aba, dtype=str)
                if df.empty:
                    continue
                c_cod = find_col(df, ['codigo', 'cod', 'ca3'], forbidden=['material', 'descri', 'nome', 'produto'])
                c_mat = find_col(df, ['material', 'produto', 'insumo', 'descri', 'nome'])
                c_cat = find_col(df, ['categoria', 'grupo', 'classe', 'tipo'])
                if c_cod and c_cat:
                    df_clean = pd.DataFrame()
                    df_clean["Código"]   = df[c_cod].apply(clean_key)
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
        st.toast(
            "⚠️ Não foi possível gravar no disco (ambiente Cloud). "
            "Use o botão 'Exportar mapa atual' para salvar suas categorias "
            "e reimporte no próximo acesso.",
            icon="⚠️"
        )
        return False


def inicializar_categorias_session():
    if "df_categorias" not in st.session_state:
        st.session_state["df_categorias"] = carregar_categorias_do_disco()


def obter_mapa_categorias() -> dict:
    df = st.session_state.get("df_categorias", pd.DataFrame())
    if df.empty:
        return {}
    return dict(zip(df["Código"].astype(str), df["Categoria"].astype(str)))


# =============================================================================
# LÓGICA DE NEGÓCIO
# =============================================================================

def calcular_cmd(qtd_total: float, dias_com_movimento: int) -> float:
    cmd_bruto = qtd_total / max(dias_com_movimento, 1)
    return float(math.ceil(cmd_bruto)) if cmd_bruto > 0 else 0.0


def calcular_tendencia(mov_filtrado: pd.DataFrame, c_mov_cod: str,
                        c_mov_qtd: str, c_mov_data: str,
                        n_dias: int = JANELA_TENDENCIA_DIAS) -> pd.DataFrame:
    # 1. Identificar o período exato que estamos olhando (Lógica ajustada para 0)
    datas_unicas = sorted(mov_filtrado['dt_formatada'].dt.date.unique(), reverse=True)
    ultimos_n_dias = datas_unicas[:n_dias]

    if not ultimos_n_dias:
        return pd.DataFrame(columns=['key', 'cmd_tendencia'])

    # 2. Filtrar apenas o movimento desses dias
    mov_tend = mov_filtrado[mov_filtrado['dt_formatada'].dt.date.isin(ultimos_n_dias)].copy()

    # 3. Agrupar somando o consumo
    consumo_tend = (
        mov_tend
        .assign(qtd_num=lambda df: df[c_mov_qtd].apply(p_num))
        .assign(key=lambda df: df[c_mov_cod].apply(clean_key))
        .groupby('key')['qtd_num'].sum()
        .reset_index()
    )

    # 4. Tratamos o resultado matematicamente
    n_dias_reais = len(ultimos_n_dias)
    consumo_tend['cmd_tendencia'] = consumo_tend['qtd_num'].apply(
        lambda x: float(math.ceil(x / n_dias_reais))
    )
    
    return consumo_tend[['key', 'cmd_tendencia']]


def calcular_sugestao(row: pd.Series, dias_pedido: int) -> int:
    cmd        = row['Consumo Médio Diário']
    est_atual  = row['Saldo Atual Satélite']
    est_minimo = row['Estoque Mínimo']

    if cmd == 0 and est_minimo > 0:
        return max(0, round(est_minimo - est_atual))
    if cmd == 0:
        return 0

    meta_final = max(cmd * dias_pedido, est_minimo)
    return max(0, round(meta_final - est_atual))


def calcular_cobertura(row: pd.Series) -> str:
    cmd = row['Consumo Médio Diário']
    saldo = row['Saldo Atual Satélite']
    
    if cmd <= 0:
        if saldo <= 0:
            return "Sem dado"
        else:
            return "+∞" 
            
    dias = int(saldo / cmd)
    return str(dias)


def calcular_delta_tendencia(row: pd.Series) -> tuple:
    cmd_per  = row['Consumo Médio Diário']
    cmd_tend = row['CMD Últ. 3 dias']

    if cmd_tend == 0 and cmd_per == 0:
        return "⬜", "Sem dado"
    if cmd_tend == 0:
        return "⬜", "Sem dado recente"
    if cmd_per == 0:
        return "📈", "+∞%"

    delta = ((cmd_tend - cmd_per) / cmd_per) * 100
    if delta > 10:
        emoji = "📈"
    elif delta < -10:
        emoji = "📉"
    else:
        emoji = "➡️"

    sinal = "+" if delta >= 0 else ""
    return emoji, f"{sinal}{delta:.0f}%"


def definir_alerta_e_acao(
    row: pd.Series,
    dict_saldos_centrais: dict,
    dict_saldos_parceiras: dict,
    consumo_outras_total: dict,
) -> tuple:
    cod        = row['Código MV']
    sug        = row['Necessidade de Ressuprimento']
    cmd        = row['Consumo Médio Diário']
    est_un     = row['Saldo Atual Satélite']
    est_minimo = row['Estoque Mínimo']

    if cmd == 0 and est_minimo <= 0 and est_un == 0:
        return "Sem Consumo", "Avaliar se é necessário inativar o item na farmácia."

    if cmd > 0 and est_un > (cmd * 60):
        return "Estoque Excessivo", "Estoque acima da necessidade de 60 dias. Devolver!."

    if cmd == 0 and est_un > 0:
        if est_un <= est_minimo:
            return "Estoque Parado", "Item sem consumo, mas dentro do estoque mínimo parametrizado."
        excedente = int(est_un - est_minimo)
        return "Estoque Parado", (
            f"{excedente} unidades acima do estoque mínimo. "
            "Considerar devolver ou remanejar para outra farmácia."
        )

    prefixo_alerta = "[ALERTA - ABAIXO DO MÍNIMO] " if (sug > 0 and est_un < est_minimo) else ""

    if sug <= 0:
        return "Estoque Suficiente", "Estoque dentro da cobertura ideal."

    saldos_parceiras   = dict_saldos_parceiras.get(cod, {})
    consumos_parceiras = consumo_outras_total.get(cod, {})
    farmacias_paradas  = [
        f"Cód {fid} ({DIC_NOMES_FARMACIAS.get(str(fid), 'Farmácia Satélite')} - {int(sf)} un.)"
        for fid, sf in saldos_parceiras.items()
        if sf > 0 and (
            consumos_parceiras.get(fid, 0) == 0 or
            sf > consumos_parceiras.get(fid, 0) * 3
        )
    ]
    locais_remanejo = " | ".join(farmacias_paradas)

    saldos_nas_centrais = dict_saldos_centrais.get(cod, {})
    saldo_total_central = sum(saldos_nas_centrais.values())

    if saldo_total_central > 0:
        central_principal = max(saldos_nas_centrais, key=saldos_nas_centrais.get)
        if saldo_total_central >= sug:
            return "Solicitar", (
                f"{prefixo_alerta}Solicitar {int(sug)} un. "
                f"ao Almoxarifado Central {central_principal}."
            )
        faltante = int(sug - saldo_total_central)
        if farmacias_paradas:
            return (
                f"Estoque Crítico no Almoxarifado {central_principal}",
                f"{prefixo_alerta}Pegar {int(saldo_total_central)} un. no Almox Central "
                f"{central_principal} e remanejar o restante ({faltante} un.) de: {locais_remanejo}."
            )
        return (
            f"Estoque Crítico no Almoxarifado {central_principal}",
            f"{prefixo_alerta}Pegar {int(saldo_total_central)} un. no Almox Central "
            f"{central_principal}. ALERTA: Sem saldo nas parceiras para cobrir as {faltante} un. restantes."
        )

    if farmacias_paradas:
        return "Remanejar", f"{prefixo_alerta}Central Zerada! Transferir de: {locais_remanejo}."

    return "Desabastecimento Crítico", (
        f"{prefixo_alerta}Sem saldo nos almoxarifados e sem estoque parado nas farmácias."
    )


# =============================================================================
# INICIALIZAÇÃO
# =============================================================================
inicializar_categorias_session()


# =============================================================================
# SIDEBAR
# =============================================================================
with st.sidebar:
    st.markdown("### 🏥 Farmácia Ativa (Alvo)")
    farm_ativa = st.session_state.get('nome_farmacia_alvo', 'Nenhuma (Aguardando processamento...)')
    if 'cod_farmacia_alvo' in st.session_state:
        st.success(f"**{farm_ativa}**\n\n*(Detectada via arquivo)*")
    else:
        st.info(farm_ativa)

    st.write("---")

    st.markdown("### ⚙️ Parâmetros do Pedido")
    dias_pedido = st.number_input(
        "Defina quantos dias de ressuprimento será solicitado:", value=15, min_value=1
    )

    ontem       = datetime.now() - timedelta(days=1)
    data_inicio = st.date_input(
        "Início do Histórico de Consumo:", value=ontem - timedelta(days=6), format="DD/MM/YYYY"
    )
    data_fim = st.date_input(
        "Fim do Período de Consumo:", value=ontem, format="DD/MM/YYYY"
    )

    st.write("---")
    with st.expander("📄 Manual de Regras Logísticas", expanded=False):
        st.markdown("""
        ### Matriz de Alertas e Ações Sugeridas

        #### 1. ⚪ Sem Consumo
        * **Condição:** CMD = 0, Estoque Mínimo = 0 e Saldo = 0.
        * **Ação:** Avaliar inativação do item.

        #### 2. 🩵 Estoque Excessivo
        * **Condição:** CMD > 0 com cobertura > 60 dias.
        * **Ação:** Devolver ao almoxarifado.

        #### 3. 🔘 Estoque Parado
        * **Condição:** CMD = 0, mas há saldo físico.
        * **Ação:** Informa excedente acima do mínimo e sugere devolução/remanejamento.

        #### 4. 🟡 Prefixo [ALERTA - ABAIXO DO MÍNIMO]
        * **Condição:** Sugestão > 0 e saldo atual < estoque mínimo parametrizado.
        * **Ação:** O item segue normalmente para Solicitar / Remanejar / Ruptura,
          mas recebe a marcação de urgência no início da ação.

        #### 5. 🟢 Estoque Suficiente
        * **Condição:** Sugestão ≤ 0 e saldo acima da margem de segurança.

        #### 6. 🔵 Solicitar
        * **Condição:** Sugestão > 0 e centrais (1, 6, 9, 41, 43) atendem a demanda.
        * **Ação:** Aponta a central com maior saldo.

        #### 7. 🟠 Estoque Crítico no Almoxarifado X
        * **Condição:** Sugestão > 0, central X tem saldo parcial.
        * **Ação:** Orienta pegar o disponível na central e informa de onde remanejar o restante.

        #### 8. 🟡 Remanejar
        * **Condição:** Centrais zeradas, parceiras têm saldo parado ou excedente.
        * **Ação:** Aponta farmácia e quantidade disponível pelo nome.

        #### 9. 🔴 Desabastecimento Crítico
        * **Condição:** Sem saldo nas centrais e sem estoque parado nas parceiras.

        ---
        **CMD:** calculado pelos dias com qualquer RM no arquivo (não dias corridos).

        **Tendência:** CMD dos últimos 3 dias com movimento vs CMD do período completo.
        """)

    with st.expander("🎖️ Créditos do Sistema", expanded=False):
        st.markdown("""
        **Idealização e Desenvolvimento:**
        * Elton Jonh Freitas Santos
        * Farmacêutico - Chefe da UDIS/HUUFMA

        *HUUFMA — Gestão e Inteligência Logística © 2026*
        """)


# =============================================================================
# INTERFACE PRINCIPAL
# =============================================================================
st.title("🏥 Gestor do Estoque - Unidade de Dispensação Farmacêutica")

nome_farm_ativa = st.session_state.get('nome_farmacia_alvo', 'Aguardando arquivo de movimento alvo...')
st.markdown(
    f"**Farmácia Ativa:** `{nome_farm_ativa}` | "
    f"**Janela Histórica:** `{data_inicio.strftime('%d/%m/%Y')}` até `{data_fim.strftime('%d/%m/%Y')}`"
)
st.write("")

tab1, tab2 = st.tabs(["⚡ Processar Pedido com IA Logística", "🗂️ Gestão de Categorias de Insumos"])


# =============================================================================
# TAB 2 — GESTÃO DE CATEGORIAS
# =============================================================================
with tab2:
    st.subheader("🗂️ Mapeamento Global de Categorias")
    st.info(
        "A tabela serve como De/Para para classificar os itens no momento da análise."
        )

    df_cat_atual = st.session_state["df_categorias"].copy()
    lista_grupos_reais = sorted(df_cat_atual["Categoria"].unique().tolist()) \
        if not df_cat_atual.empty else CATEGORIAS_PADRAO
    if "OUTROS" not in lista_grupos_reais:
        lista_grupos_reais.append("OUTROS")

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("📦 Total de Insumos Cadastrados", len(df_cat_atual))
    mc2.metric("🏷️ Categorias Mapeadas",
               df_cat_atual["Categoria"].nunique() if not df_cat_atual.empty else 0)
    mc3.metric("❓ Sem Classificação (OUTROS)",
               len(df_cat_atual[df_cat_atual["Categoria"] == "OUTROS"]) if not df_cat_atual.empty else 0)

    st.write("---")
    fc1, fc2 = st.columns([3, 1])
    filtro_termo_cat = fc1.text_input("🔍 Pesquisar por nome ou código:", value="")
    filtro_sel_cat   = fc2.selectbox("Filtrar por Grupo:", ["TODOS"] + lista_grupos_reais)

    alguem_pesquisou = filtro_termo_cat.strip() != "" or filtro_sel_cat != "TODOS"

    if alguem_pesquisou:
        df_filtrado_cat = df_cat_atual.copy()
        if filtro_termo_cat:
            df_filtrado_cat = df_filtrado_cat[
                df_filtrado_cat["Código"].astype(str).str.contains(filtro_termo_cat, case=False, na=False) |
                df_filtrado_cat["Material"].astype(str).str.contains(filtro_termo_cat, case=False, na=False)
            ]
        if filtro_sel_cat != "TODOS":
            df_filtrado_cat = df_filtrado_cat[df_filtrado_cat["Categoria"] == filtro_sel_cat]

        st.markdown(f"##### 📋 Itens Encontrados ({len(df_filtrado_cat)} registros)")

        df_editor_output = st.data_editor(
            df_filtrado_cat.reset_index(drop=True),
            use_container_width=True, hide_index=True, num_rows="dynamic",
            column_config={
                "Código":    st.column_config.TextColumn("Código MV", required=True, width="small"),
                "Material":  st.column_config.TextColumn("Descrição do Insumo", required=True, width="large"),
                "Categoria": st.column_config.SelectboxColumn(
                    "Categoria Logística", options=lista_grupos_reais, required=True, width="medium"
                ),
            },
            key="editor_categorias",
        )

        col_salvar, col_reset = st.columns([4, 1])
        if col_salvar.button("💾 SALVAR ALTERAÇÕES PERMANENTEMENTE", use_container_width=True):
            codigos_visiveis = set(df_filtrado_cat["Código"].astype(str).tolist())
            df_base = st.session_state["df_categorias"].copy()
            df_base_limpo = df_base[~df_base["Código"].astype(str).isin(codigos_visiveis)]
            df_novo = pd.concat([df_base_limpo, df_editor_output], ignore_index=True)
            df_novo = df_novo.drop_duplicates("Código", keep="last")
            df_novo = df_novo[df_novo["Código"].astype(str).str.strip() != ""]
            st.session_state["df_categorias"] = df_novo.reset_index(drop=True)
            if salvar_categorias_no_disco(st.session_state["df_categorias"]):
                st.success("✅ Alterações salvas em 'Categorias_base.xlsx'.")
                st.rerun()

        if col_reset.button("🔄 Cancelar", use_container_width=True):
            st.session_state.pop("df_categorias", None)
            st.rerun()

        # Download de itens OUTROS para classificação em lote
        df_outros = df_cat_atual[df_cat_atual["Categoria"] == "OUTROS"]
        if not df_outros.empty:
            st.download_button(
                "📥 Exportar itens OUTROS para classificação em lote",
                data=exportar_excel_padronizado(df_outros, "Para Classificar"),
                file_name=f"Itens_OUTROS_{datetime.now().strftime('%d%m%y')}.xlsx",
                use_container_width=True,
            )
        # Botão exportar mapa completo — sempre disponível
        st.write("")
        df_completo_export = st.session_state["df_categorias"].copy()
        if not df_completo_export.empty:
            st.download_button(
                "📤 Exportar mapa COMPLETO de categorias (.xlsx)",
                data=exportar_excel_padronizado(df_completo_export, "Categorias"),
                file_name=f"Categorias_base_{datetime.now().strftime('%d%m%y')}.xlsx",
                use_container_width=True,
                help="Baixe este arquivo e guarde no repositório como 'Categorias_base.xlsx' "
                     "para que as categorias sejam carregadas automaticamente na próxima sessão."
            )
    else:
        st.warning("🔍 Utilize os filtros acima para visualizar e editar os registros.")
        st.write("")
        df_completo_export = st.session_state["df_categorias"].copy()
        if not df_completo_export.empty:
            st.download_button(
                "📤 Exportar mapa COMPLETO de categorias (.xlsx)",
                data=exportar_excel_padronizado(df_completo_export, "Categorias"),
                file_name=f"Categorias_base_{datetime.now().strftime('%d%m%y')}.xlsx",
                use_container_width=True,
                help="Baixe este arquivo e guarde no repositório como 'Categorias_base.xlsx' "
                     "para que as categorias sejam carregadas automaticamente na próxima sessão."
            )


# =============================================================================
# TAB 1 — PROCESSAR PEDIDO
# =============================================================================
with tab1:
    with st.container(border=True):
        st.markdown("##### 📥 Upload das Fontes de Dados Obrigatórias (AGHU)")
        col1, col2 = st.columns(2)
        file_mov_alvo  = col1.file_uploader("1. Movimento da Farmácia Alvo (.csv)", type=["csv"])
        file_est_geral = col2.file_uploader("2. Estoque Geral de todos os Almoxarifados (.csv)", type=["csv"])
        st.write("")
        files_mov_parceiras = st.file_uploader(
            "3. Movimentos das outras Farmácias — Ativa a análise da viabilidade de remanejamentos (opcional) (Múltiplos.csv)",
            type=["csv"], accept_multiple_files=True,
        )

    st.write("")

    if file_mov_alvo and file_est_geral:
        if st.button("🚀 ANALISAR OS DADOS COM INTELIGÊNCIA LOGÍSTICA", use_container_width=True):
            st.session_state['disparar_processamento_huufma'] = True

        if st.session_state.get('disparar_processamento_huufma', False):
            progress = st.progress(0, text="📂 Lendo arquivos...")

            try:
                mov       = ler_csv_cached(file_mov_alvo.read(), file_mov_alvo.name)
                est_geral = ler_csv_cached(file_est_geral.read(), file_est_geral.name)

                progress.progress(10, text="🔍 Identificando colunas...")

                cols_mov = {
                    'código':       find_col(mov, ['material', 'cod', 'ca3']),
                    'quantidade':   find_col(mov, ['quant']),
                    'tipo':         find_col(mov, ['tipo']),
                    'data':         find_col(mov, ['data', 'ger']),
                    'almoxarifado': find_col(mov, ['almox']),
                }
                if not validar_colunas(cols_mov, file_mov_alvo.name):
                    st.session_state['disparar_processamento_huufma'] = False
                    st.stop()

                cols_est = {
                    'código':       find_col(est_geral, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod']),
                    'quantidade':   find_col(est_geral, ['qtde disp', 'disponivel']),
                    'produto':      find_col(est_geral, ['material', 'produto', 'descri']),
                    'almoxarifado': find_col(est_geral, ['almox']),
                    'mínimo':       find_col(est_geral, ['qtde estq min', 'estoque minimo', 'minimo']),
                }
                if not validar_colunas(
                    {k: v for k, v in cols_est.items() if k != 'mínimo'}, file_est_geral.name
                ):
                    st.session_state['disparar_processamento_huufma'] = False
                    st.stop()

                c_mov_cod   = cols_mov['código']
                c_mov_qtd   = cols_mov['quantidade']
                c_mov_tipo  = cols_mov['tipo']
                c_mov_data  = cols_mov['data']
                c_mov_almox = cols_mov['almoxarifado']
                c_est_cod   = cols_est['código']
                c_est_qtd   = cols_est['quantidade']
                c_est_prod  = cols_est['produto']
                c_est_almox = cols_est['almoxarifado']
                c_est_min   = cols_est['mínimo']

                # --- TRAVA INTELIGENTE: DETECÇÃO AUTOMÁTICA DA FARMÁCIA ALVO ---
                codigos_almox = mov[c_mov_almox].dropna().astype(str).apply(clean_key)
                cod_farmacia_alvo = codigos_almox[codigos_almox != ""].mode()[0]
                
                st.session_state['cod_farmacia_alvo'] = cod_farmacia_alvo
                st.session_state['nome_farmacia_alvo'] = DIC_NOMES_FARMACIAS.get(cod_farmacia_alvo, f"Almoxarifado (Cód. {cod_farmacia_alvo})")

                progress.progress(20, text="🏗️ Processando estoque geral...")

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

                # Centrais vetorizadas
                centrais_alvo         = ['1', '6', '9', '41', '43']
                est_centrais_filtrado = est_geral[est_geral['almox_limpo'].isin(centrais_alvo)]
                dict_saldos_centrais  = {}
                if not est_centrais_filtrado.empty:
                    _c = (est_centrais_filtrado.groupby(['key', 'almox_limpo'])['saldo_num']
                          .sum().reset_index())
                    dict_saldos_centrais = (
                        _c.groupby('key')[['almox_limpo', 'saldo_num']]
                        .apply(lambda g: dict(zip(g['almox_limpo'], g['saldo_num'])),
                               include_groups=False)
                        .to_dict()
                    )

                # Parceiras vetorizadas
                cod_parceiras = [c for c in ['7', '13', '31', '34', '39'] if c != cod_farmacia_alvo]
                est_outras    = est_geral[est_geral['almox_limpo'].isin(cod_parceiras)]
                dict_saldos_parceiras = {}
                if not est_outras.empty:
                    _p = (est_outras.groupby(['key', 'almox_limpo'])['saldo_num']
                          .sum().reset_index())
                    dict_saldos_parceiras = (
                        _p.groupby('key')[['almox_limpo', 'saldo_num']]
                        .apply(lambda g: dict(zip(g['almox_limpo'], g['saldo_num'])),
                               include_groups=False)
                        .to_dict()
                    )

                progress.progress(40, text="📊 Calculando consumo e auditando calendário...")

                mov = mov.copy()
                # Correção do dayfirst=True
                mov['dt_formatada'] = pd.to_datetime(mov[c_mov_data], errors='coerce')
                mov_filtrado = mov[
                    (mov['dt_formatada'].dt.date >= data_inicio) &
                    (mov['dt_formatada'].dt.date <= data_fim) &
                    (mov[c_mov_tipo].astype(str).str.upper() == 'RM')
                ].copy()

                # ── AUDITORIA DE CALENDÁRIO ──────────────────────────────────
                dias_ideais       = set(pd.date_range(start=data_inicio, end=data_fim).date)
                dias_com_movimento = sorted(
                    mov_filtrado['dt_formatada'].dt.date.dropna().unique()
                )
                dias_vazios = sorted(dias_ideais - set(dias_com_movimento))
                st.session_state['datas_sem_movimento_huufma'] = [
                    d.strftime('%d/%m/%Y') for d in dias_vazios
                ]
                st.session_state['n_dias_com_movimento'] = len(dias_com_movimento)

                # ── CMD BASEADO NOS DIAS COM MOVIMENTO ───
                n_dias_efetivos = max(len(dias_com_movimento), 1)

                consumo = (
                    mov_filtrado
                    .assign(qtd_num=lambda df: df[c_mov_qtd].apply(p_num))
                    .assign(key=lambda df: df[c_mov_cod].apply(clean_key))
                    .groupby('key')['qtd_num'].sum()
                    .reset_index()
                    .rename(columns={'qtd_num': 'total_consumo'})
                )
                consumo['cmd'] = consumo['total_consumo'].apply(
                    lambda x: calcular_cmd(x, n_dias_efetivos)
                )

                # ── TENDÊNCIA ─────────────────
                df_tendencia = calcular_tendencia(
                    mov_filtrado, c_mov_cod, c_mov_qtd, c_mov_data,
                    n_dias=JANELA_TENDENCIA_DIAS
                )

                progress.progress(55, text="🔄 Cruzando dados entre farmácias...")

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
                            # Correção do dayfirst=True
                            df_p['dt_formatada'] = pd.to_datetime(
                                df_p[c_p_data], errors='coerce'
                            )
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

                mapa_produtos = (
                    est_geral.drop_duplicates(subset=['key'])
                    .set_index('key')[c_est_prod].to_dict()
                )
                consumo_map   = consumo.set_index('key')['cmd'].to_dict()
                tendencia_map = df_tendencia.set_index('key')['cmd_tendencia'].to_dict() \
                    if not df_tendencia.empty else {}

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

                final['Cobertura (dias)'] = final.apply(calcular_cobertura, axis=1)

                final['CMD Últ. 3 dias'] = final['Código MV'].map(tendencia_map).fillna(0)
                tend_resultados = final.apply(calcular_delta_tendencia, axis=1)
                final['Tendência']   = [r[0] for r in tend_resultados]
                final['Δ% Tendência'] = [r[1] for r in tend_resultados]

                mapa_cat = obter_mapa_categorias()
                final['Categoria'] = final['Código MV'].map(mapa_cat).fillna('OUTROS')

                total_itens   = len(final)
                itens_mapeados = final['Categoria'].ne('OUTROS').sum()
                itens_outros   = total_itens - itens_mapeados
                st.session_state['diagnostico_categorias'] = {
                    'total': total_itens,
                    'mapeados': int(itens_mapeados),
                    'outros': int(itens_outros),
                }

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

                st.session_state['df_final_huufma']        = final
                st.session_state['n_dias_efetivos_huufma'] = n_dias_efetivos
                st.session_state['disparar_processamento_huufma'] = False
                progress.progress(100, text="✅ Processamento concluído!")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Erro crítico no processamento: {e}")
                st.session_state['disparar_processamento_huufma'] = False

        # =================================================================
        # PAINEL DE RESULTADOS
        # =================================================================
        if 'df_final_huufma' in st.session_state:
            df_view         = st.session_state['df_final_huufma'].copy()
            n_dias_efetivos = st.session_state.get('n_dias_efetivos_huufma', '—')
            datas_vazias    = st.session_state.get('datas_sem_movimento_huufma', [])
            cod_farmacia_alvo = st.session_state.get('cod_farmacia_alvo', 'Alvo')

            if datas_vazias:
                st.warning(
                    f"⚠️ **Auditoria de Calendário:** Sem movimentação em "
                    f"`{len(datas_vazias)}` dia(s) do período: "
                    f"`{', '.join(datas_vazias)}`. "
                    f"O CMD foi calculado com **{n_dias_efetivos} dias efetivos** de movimento."
                )
            else:
                st.info(
                    f"✅ **Calendário completo:** Todos os dias do período tiveram movimento. "
                    f"CMD calculado com **{n_dias_efetivos} dias efetivos**."
                )

            st.write("---")

            diag = st.session_state.get('diagnostico_categorias', {})
            if diag:
                pct = round(diag['mapeados'] / max(diag['total'], 1) * 100, 1)
                if diag['outros'] > 0:
                    st.warning(
                        f"🏷️ **Cobertura de Categorias:** {diag['mapeados']} de {diag['total']} itens "
                        f"classificados ({pct}%). **{diag['outros']} item(ns) como 'OUTROS'.**"
                    )
                else:
                    st.success(
                        f"🏷️ **Categorias:** 100% dos {diag['total']} itens classificados."
                    )

            # ── DEFINIÇÃO DAS VARIÁVEIS DE ORDEM PARA TODOS OS DOWNLOADS E PAINEL ──
            COL_SUG = f'Necessidade de Ressuprimento ({dias_pedido} dias)'
            ordem_cols = [
                'Código MV', 'Material', 'Categoria',
                'Estoque Mínimo', 'Saldo Atual Satélite', 'Cobertura (dias)',
                'CMD Últ. 3 dias', 'Consumo Médio Diário', COL_SUG,
                'Tendência', 'Δ% Tendência',
                'Saldo Almox. Centrais Unificado',
                'Parecer Logístico / Alerta', 'Ação Logística Sugerida',
            ]
            larguras_rel = {
                'Código MV': 12, 'Material': 45, 'Categoria': 16,
                'Estoque Mínimo': 16, 'Saldo Atual Satélite': 18,
                'Cobertura (dias)': 14, 'CMD Últ. 3 dias': 16,
                'Consumo Médio Diário': 18, COL_SUG: 26, 'PEDIDO (X DIAS)': 20,
                'Tendência': 10, 'Δ% Tendência': 12,
                'Saldo Almox. Centrais Unificado': 28,
                'Parecer Logístico / Alerta': 26, 'Ação Logística Sugerida': 50,
            }

            def _preparar_df_card(df_raw: pd.DataFrame) -> pd.DataFrame:
                df = df_raw.rename(columns={'Necessidade de Ressuprimento': COL_SUG})
                cols_ok = [c for c in ordem_cols if c in df.columns]
                return df[cols_ok]

            # ── MÉTRICAS ─────────────────────────────────────────────────
            df_desabast = df_view[df_view['Parecer Logístico / Alerta'] == "Desabastecimento Crítico"].sort_values('Material')
            df_remanej  = df_view[df_view['Parecer Logístico / Alerta'] == "Remanejar"].sort_values('Material')
            df_caf      = df_view[df_view['Parecer Logístico / Alerta'].str.contains("Solicitar|Almoxarifado", na=False)].sort_values('Material')
            df_excesso  = df_view[df_view['Parecer Logístico / Alerta'].isin(
                ["Estoque Excessivo", "Estoque Parado"]
            )].sort_values('Material')

            c0, c1, c2, c3, c4 = st.columns(5)
            c0.metric("📅 Dias Efetivos de Consumo",
                      f"{n_dias_efetivos} dias",
                      delta=f"-{len(datas_vazias)} sem mov." if datas_vazias else "Período completo",
                      delta_color="inverse" if datas_vazias else "normal")
            with c1:
                st.metric("🚨 Desabastecimento Crítico", f"{len(df_desabast)} itens")
                st.download_button("📥 Extrair", data=exportar_excel_padronizado(_preparar_df_card(df_desabast), "Rupturas"),
                    file_name=f"Rupturas_{cod_farmacia_alvo}.xlsx", key="ex_c1", use_container_width=True)
            with c2:
                st.metric("🔄 Remanejamento Potencial", f"{len(df_remanej)} itens")
                st.download_button("📥 Extrair", data=exportar_excel_padronizado(_preparar_df_card(df_remanej), "Remanejamento"),
                    file_name=f"Remanejamento_{cod_farmacia_alvo}.xlsx", key="ex_c2", use_container_width=True)
            with c3:
                st.metric("📦 Disponível no Almoxarifado", f"{len(df_caf)} itens")
                st.download_button("📥 Extrair", data=exportar_excel_padronizado(_preparar_df_card(df_caf), "Disponiveis_CAF"),
                    file_name=f"Disponiveis_{cod_farmacia_alvo}.xlsx", key="ex_c3", use_container_width=True)
            with c4:
                st.metric("⚠️ Excesso / Sem Giro", f"{len(df_excesso)} itens")
                st.download_button("📥 Extrair", data=exportar_excel_padronizado(_preparar_df_card(df_excesso), "Overstock"),
                    file_name=f"Overstock_{cod_farmacia_alvo}.xlsx", key="ex_c4", use_container_width=True)

            # ── GRÁFICOS ──────────────────────────────────────
            st.write("")
            g1, g2 = st.columns(2)

            with g1:
                st.markdown("**Saúde Geral do Estoque**")
                df_g1 = df_view.copy()
                df_g1.loc[
                    df_g1['Parecer Logístico / Alerta'].str.contains("Almoxarifado", na=False),
                    'Parecer Logístico / Alerta'
                ] = "Estoque Crítico CAF"
                df_g1_grp = (
                    df_g1[df_g1['Parecer Logístico / Alerta'] != 'Sem Consumo']
                    .groupby('Parecer Logístico / Alerta')['Código MV'].count()
                    .reset_index()
                    .rename(columns={'Código MV': 'Quantidade', 'Parecer Logístico / Alerta': 'Status'})
                )
                if not df_g1_grp.empty:
                    st.altair_chart(
                        alt.Chart(df_g1_grp).mark_arc(innerRadius=65, stroke='#fff').encode(
                            theta=alt.Theta('Quantidade:Q'),
                            color=alt.Color('Status:N', scale=alt.Scale(
                                domain=list(MAPA_CORES_GRAFICO.keys()),
                                range=list(MAPA_CORES_GRAFICO.values())
                            ), legend=None),
                            tooltip=['Status:N', 'Quantidade:Q']
                        ).properties(height=350),
                        use_container_width=True
                    )

            with g2:
                st.markdown("**Matriz de Urgência por Categoria**")
                df_g2 = df_view[df_view['Categoria'] != 'OUTROS'].copy()
                df_g2.loc[
                    df_g2['Parecer Logístico / Alerta'].str.contains("Almoxarifado", na=False),
                    'Parecer Logístico / Alerta'
                ] = "Estoque Crítico CAF"
                
                df_g2_grp = (
                    df_g2[df_g2['Parecer Logístico / Alerta'] != 'Sem Consumo']
                    .groupby(['Categoria', 'Parecer Logístico / Alerta'])['Código MV'].count()
                    .reset_index()
                    .rename(columns={'Código MV': 'Itens', 'Parecer Logístico / Alerta': 'Parecer'})
                )
                
                if not df_g2_grp.empty:
                    base = alt.Chart(df_g2_grp).encode(
                        x=alt.X('Parecer:N', title=None, axis=alt.Axis(labels=False, ticks=False, domain=False)),
                        y=alt.Y('Categoria:N', title=None)
                    )

                    heatmap = base.mark_rect(cornerRadius=6, stroke='white', strokeWidth=3).encode(
                        color=alt.Color('Parecer:N', scale=alt.Scale(
                            domain=list(MAPA_CORES_GRAFICO.keys()),
                            range=list(MAPA_CORES_GRAFICO.values())
                        ), legend=None),
                        opacity=alt.Opacity('Itens:Q', scale=alt.Scale(range=[0.4, 1.0]), legend=None),
                        tooltip=[
                            alt.Tooltip('Categoria:N', title='Categoria'),
                            alt.Tooltip('Parecer:N', title='Status'),
                            alt.Tooltip('Itens:Q', title='Quantidade')
                        ]
                    )

                    textos = base.mark_text(baseline='middle', fontSize=15, fontWeight='bold', color='#1E293B').encode(
                        text='Itens:Q'
                    )

                    grafico_matriz = (heatmap + textos).properties(height=350)
                    st.altair_chart(grafico_matriz, use_container_width=True)
                else:
                    st.info("Nenhuma categoria vinculada ativa no filtro atual.")

            # ── LEGENDA UNIFICADA (RODAPÉ) ────────────────────────────────
            st.write("")
            legend_html = "<div style='display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; padding: 10px; background-color: #F8FAFC; border-radius: 8px; border: 1px solid #E2E8F0;'>"
            for status, color in MAPA_CORES_GRAFICO.items():
                legend_html += f"<div style='display: flex; align-items: center;'><div style='width: 14px; height: 14px; background-color: {color}; border-radius: 50%; margin-right: 6px;'></div><span style='font-size: 13px; color: #334155; font-weight: 500;'>{status}</span></div>"
            legend_html += "</div>"
            st.markdown(legend_html, unsafe_allow_html=True)


            # ── PAINEL INTERATIVO ─────────────────────────────────────────
            st.write("---")
            st.markdown("#### 📋 Resultado da Análise Inteligente")

            with st.container(border=True):
                st.markdown("##### 🔍 Filtros Dinâmicos")
                f1, f2, f3 = st.columns([2, 2, 1])
                busca_nome  = f1.text_input("Filtrar por nome ou código:", value="", key="busca_nome")
                opcoes_alertas = sorted(df_view['Parecer Logístico / Alerta'].unique().tolist())
                busca_alerta   = f2.multiselect("Filtrar por Parecer:", options=opcoes_alertas, key="busca_alerta")
                busca_cat      = f3.selectbox("Categoria:", ["TODAS"] + sorted(df_view['Categoria'].unique().tolist()), key="busca_cat")

            df_filtrado = df_view.copy()
            if busca_nome:
                df_filtrado = df_filtrado[
                    df_filtrado['Material'].astype(str).str.contains(busca_nome, case=False, na=False) |
                    df_filtrado['Código MV'].astype(str).str.contains(busca_nome, case=False, na=False)
                ]
            if busca_alerta:
                df_filtrado = df_filtrado[df_filtrado['Parecer Logístico / Alerta'].isin(busca_alerta)]
            if busca_cat != "TODAS":
                df_filtrado = df_filtrado[df_filtrado['Categoria'] == busca_cat]

            # Exibir DataFrame com a ordem definida em 'ordem_cols'
            df_filtrado = df_filtrado.rename(columns={'Necessidade de Ressuprimento': COL_SUG})
            cols_exibicao = [c for c in ordem_cols if c in df_filtrado.columns]
            
            st.dataframe(
                df_filtrado[cols_exibicao].sort_values('Material'),
                use_container_width=True, hide_index=True,
                column_config={
                    "Código MV":               st.column_config.TextColumn("Código MV", width="small"),
                    "Material":                st.column_config.TextColumn("Descrição", width="large"),
                    "Categoria":               st.column_config.TextColumn("Categoria", width="medium"),
                    "Estoque Mínimo":          st.column_config.NumberColumn("Mínimo", format="%d"),
                    "Saldo Atual Satélite":    st.column_config.NumberColumn("Saldo", format="%d"),
                    "Cobertura (dias)":        st.column_config.TextColumn("Cobertura", width="small"),
                    "CMD Últ. 3 dias":         st.column_config.NumberColumn("CMD 3 dias", format="%d"),                    
                    "Consumo Médio Diário":    st.column_config.NumberColumn("CMD", format="%d"),
                    COL_SUG:                   st.column_config.NumberColumn("Necessidade", format="%d"),                    
                    "Tendência":               st.column_config.TextColumn("Tend.", width="small"),
                    "Δ% Tendência":            st.column_config.TextColumn("Δ%", width="small"),
                    "Saldo Almox. Centrais Unificado": st.column_config.NumberColumn("Saldo Central", format="%d"),
                    "Parecer Logístico / Alerta": st.column_config.TextColumn("Parecer", width="medium"),
                    "Ação Logística Sugerida":    st.column_config.TextColumn("Ação Sugerida", width="large"),
                }
            )

            # ── DOWNLOADS ─────────────────────────────────────────────────
            st.write("")
            
            # DataFrame base de exportação para a Aba Única e Filtro
            df_export_geral = df_view.rename(columns={'Necessidade de Ressuprimento': COL_SUG})[ordem_cols]

            b1, b2, b3 = st.columns(3)
            with b1:
                st.download_button(
                    "📥 BAIXAR RELATÓRIO COMPLETO — ABA ÚNICA (.XLSX)",
                    data=exportar_excel_padronizado(df_export_geral, "Painel Geral"),
                    file_name=f"Painel_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )
            with b2:
                # Ordem e colunas exclusivas para o relatório de abas, conforme solicitado
                ordem_abas = [
                    'Código MV', 'Material', 'Categoria', 'Estoque Mínimo', 
                    'Parecer Logístico / Alerta', 'Saldo Atual Satélite', 'Cobertura (dias)', 
                    'Δ% Tendência', 'Necessidade de Ressuprimento'
                ]
                # Criando df exclusivo para as abas com a renomeação da necessidade
                df_pedido_abas = df_view.copy()
                df_pedido_abas = df_pedido_abas.rename(columns={'Necessidade de Ressuprimento': 'PEDIDO (X DIAS)'})
                
                # Atualiza a lista da ordem para a coluna renomeada
                ordem_abas_final = [c if c != 'Necessidade de Ressuprimento' else 'PEDIDO (X DIAS)' for c in ordem_abas]
                
                st.download_button(
                    "📥 BAIXAR PEDIDO - CLASSIFICADO POR CATEGORIA (.XLSX)",
                    data=exportar_excel_multi_aba(
                        df_pedido_abas, # Passa o DF completo para poder filtrar por "Ação Logística" se precisar
                        ordem_abas_final, # Garante que SÓ essas colunas sairão no arquivo
                        col_categoria='Categoria',
                        col_alerta='Parecer Logístico / Alerta',
                        larguras=larguras_rel,
                        excluir_acoes=["Avaliar se é necessário inativar o item na farmácia."],
                    ),
                    file_name=f"Pedido_Abas_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )
            with b3:
                df_filtrado_export = df_filtrado[cols_exibicao]
                st.download_button(
                    "📥 BAIXAR RESULTADO DA ANÁLISE ATUAL - FILTRADA (.XLSX)",
                    data=exportar_excel_padronizado(df_filtrado_export, "Filtro Atual"),
                    file_name=f"Filtro_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )

    else:
        st.session_state['disparar_processamento_huufma'] = False
        st.warning("⚠️ Aguardando upload dos arquivos obrigatórios para iniciar o processamento.")
