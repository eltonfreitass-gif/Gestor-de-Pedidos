import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import unicodedata
import math
import altair as alt
import requests
from datetime import datetime, timedelta, date
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
    div.stButton > button,
    div.stDownloadButton > button {
        background-color: #10B981 !important;
        color: white !important;
        font-weight: 600 !important;
        border: none !important;
        padding: 0.5rem 2rem !important;
        border-radius: 0.5rem !important;
        transition: all 0.3s ease;
    }
    div.stButton > button:hover,
    div.stDownloadButton > button:hover {
        background-color: #059669 !important;
        color: white !important;
        border: none !important;
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

# Fonte preferencial de categorias em Google Sheets público (leitura).
# A edição oficial deve ser feita diretamente na planilha Google;
# Categorias_base.xlsx permanece apenas como fallback de leitura.
USAR_GOOGLE_SHEETS_CATEGORIAS = True
GOOGLE_SHEETS_CATEGORIAS_ID = "122sjqkGtwta8MJwhlIJYy8lm9bFuQfTXc4MoTPBT3uk"
GOOGLE_SHEETS_CATEGORIAS_GID = "0"
GOOGLE_SHEETS_CATEGORIAS_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_CATEGORIAS_ID}"
    f"/export?format=csv&gid={GOOGLE_SHEETS_CATEGORIAS_GID}"
)
GOOGLE_SHEETS_CATEGORIAS_LINK = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_CATEGORIAS_ID}/edit?usp=sharing"
)

# Janela de tendência de consumo (últimos N dias com movimento)
JANELA_TENDENCIA_DIAS = 3

# Faixas de alerta de validade (dias até o vencimento)
VALIDADE_CRITICO_DIAS  = 30   # 🔴 vence em até 30 dias
VALIDADE_ATENCAO_DIAS  = 90   # 🟡 vence em até 90 dias

# Chave session_state para o link SharePoint de validades
SS_SHAREPOINT_URL = 'sharepoint_validades_url'

# Farmácias satélites (códigos de almoxarifado)
CODIGOS_FARMACIAS = ['7', '13', '31', '34', '39']

# Almoxarifados que devem ser ignorados em toda a lógica do aplicativo.
# O almoxarifado 45 contém medicamentos vinculados a pacientes, fora da gestão
# operacional da UDIS/farmácias satélites para pedidos, remanejamentos, validade
# e consolidação de estoque.
CODIGOS_ALMOXARIFADOS_EXCLUIR = {'45'}
DIC_ALMOXARIFADOS_EXCLUIDOS = {'45': 'Almoxarifado 45 - estoque de pacientes'}


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


def almoxarifado_excluido(v) -> bool:
    """Indica se o código de almoxarifado deve ser ignorado pelo app."""
    return clean_key(v) in CODIGOS_ALMOXARIFADOS_EXCLUIR


def filtrar_almoxarifados_excluidos(df: pd.DataFrame, col_almox: str | None = None) -> pd.DataFrame:
    """Remove linhas de almoxarifados fora da gestão do aplicativo.

    A regra é aplicada de forma centralizada para evitar que o almoxarifado 45
    influencie saldos, validades, consumo, consolidação, remanejamentos ou
    relatórios. Se a coluna de almoxarifado não for encontrada, retorna cópia
    do dataframe sem alteração para não quebrar fluxos que não possuem essa coluna.
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    if col_almox is None:
        col_almox = find_col(df, ['almox'])

    if not col_almox or col_almox not in df.columns:
        return df.copy()

    out = df.copy()
    mask_excluir = out[col_almox].apply(clean_key).isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)
    return out.loc[~mask_excluir].copy()


def p_num(v) -> float:
    try:
        if pd.isna(v) or str(v).strip() == "":
            return 0.0
        s = str(v).strip()
        # Preserva o sinal de estornos/devoluções: "-5" ou "(5)" => negativo
        negativo = s.startswith('-') or (s.startswith('(') and s.endswith(')'))
        l = re.sub(r'[^0-9,.]', '', s)
        if "," in l and "." in l:
            l = l.replace(".", "").replace(",", ".")
        elif "," in l:
            l = l.replace(",", ".")
        elif "." in l and re.fullmatch(r'\d{1,3}(\.\d{3})+', l):
            # Padrão BR de milhar sem decimais: "1.234" => 1234 (não 1,234)
            l = l.replace(".", "")
        valor = float(l) if l else 0.0
        return -valor if negativo else valor
    except Exception:
        return 0.0


def p_num_series(s: pd.Series) -> pd.Series:
    """Versão vetorizada de p_num para colunas inteiras — 5-15x mais rápida que apply().
    Trata separadores BR (vírgula decimal, ponto de milhar) e valores negativos com parênteses."""
    s = s.astype(str).str.strip()
    negativo = s.str.startswith('-') | (s.str.startswith('(') & s.str.endswith(')'))
    # Caso simples: só dígitos (maioria dos dados do AGHUx) — pd.to_numeric direto
    resultado = pd.to_numeric(s, errors='coerce')
    # Fallback para linhas que falharam (vírgula decimal, parênteses, etc.)
    mask_falhou = resultado.isna() & s.ne('') & s.ne('nan')
    if mask_falhou.any():
        s_fix = s[mask_falhou].str.replace(r'[^0-9,.]', '', regex=True)
        tem_ambos = s_fix.str.contains(',') & s_fix.str.contains('.')
        so_virgula = s_fix.str.contains(',') & ~s_fix.str.contains('.')
        milhar_br  = ~s_fix.str.contains(',') & s_fix.str.contains('.') & \
                     s_fix.str.match(r'^\d{1,3}(\.\d{3})+$')
        s_fix = s_fix.copy()
        s_fix[tem_ambos]  = s_fix[tem_ambos].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        s_fix[so_virgula] = s_fix[so_virgula].str.replace(',', '.', regex=False)
        s_fix[milhar_br]  = s_fix[milhar_br].str.replace('.', '', regex=False)
        resultado[mask_falhou] = pd.to_numeric(s_fix, errors='coerce')
    resultado = resultado.fillna(0.0)
    resultado[negativo] = resultado[negativo].abs() * -1
    return resultado


def find_col(df: pd.DataFrame, terms: list, forbidden: list = None):
    if forbidden is None:
        forbidden = []
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

def _col_letter(n: int) -> str:
    """Converte índice 0-based para letra(s) de coluna Excel (A, B, ..., AA, AB...)."""
    result = ""
    while n >= 0:
        result = chr(n % 26 + ord('A')) + result
        n = n // 26 - 1
    return result


def _altura_linha(row_values: list, col_widths: list) -> int:
    """Estima a altura da linha em pontos com base no texto mais longo vs largura da coluna."""
    max_lines = 1
    for val, w in zip(row_values, col_widths):
        texto = str(val) if val is not None and not (isinstance(val, float) and math.isnan(val)) else ""
        chars_por_linha = max(int(w * 1.1), 8)
        linhas = math.ceil(len(texto) / chars_por_linha) if texto else 1
        max_lines = max(max_lines, linhas)
    return max(18, min(int(max_lines * 14.5), 130))


# Larguras padrão por nome amigável de coluna — fonte única usada em ambas as funções
LARGURAS_PADRAO_EXCEL = {
    'Código MV': 12, 'Material': 45, 'Categoria': 16,
    'Saldo Atual Satélite': 20, 'Consumo Médio Diário': 20,
    'Estoque Mínimo': 16, 'Cobertura (dias)': 14,
    'CMD Últ. 3 dias': 16, 'Tendência': 12, 'Δ% Tendência': 12,
    'Saldo Almox. Centrais Unificado': 28, 'Ação Logística Sugerida': 50,
    'Parecer Logístico / Alerta': 26,
}


def _aplicar_estilo_aba(ws, wb, df: pd.DataFrame, col_alerta: str,
                         larguras: dict, fmt_h, fmt_b, fmt_t,
                         fmt_status_cache: dict, fmt_critico,
                         ajustar_altura_linhas: bool = True,
                         ocultar_colunas: list | None = None,
                         orientacao_impressao: str = "landscape") -> None:
    """Aplica cabeçalho, larguras, altura dinâmica e formatação condicional a uma aba.

    Quando ajustar_altura_linhas=False, não fixa a altura das linhas de dados.
    Com texto quebrado (text_wrap=True), o Excel recalcula a altura ao abrir o arquivo,
    o que evita cortes em relatórios com descrições/observações longas.
    """
    # Cabeçalho
    for ci, cn in enumerate(df.columns):
        ws.write(0, ci, cn, fmt_h)
    ws.set_row(0, 40)

    col_widths = []
    for ci, cn in enumerate(df.columns):
        larg = larguras.get(cn, 22)
        fmt_col = fmt_t if cn in ('Material', 'Ação Logística Sugerida', 'Observação') else fmt_b
        ws.set_column(ci, ci, larg, fmt_col)
        col_widths.append(larg)

    # Altura das linhas
    # Para relatórios operacionais com texto longo, pode ser melhor deixar o Excel
    # ajustar a altura automaticamente ao abrir o arquivo.
    if ajustar_altura_linhas:
        for ri, row_vals in enumerate(df.itertuples(index=False), start=1):
            h = _altura_linha(list(row_vals), col_widths)
            ws.set_row(ri, h)

    ws.freeze_panes(1, 0)

    # Colunas auxiliares podem ser usadas para manter cores/condicionais
    # sem aparecerem para o usuário final no relatório.
    for cn in (ocultar_colunas or []):
        if cn in df.columns:
            ci = df.columns.get_loc(cn)
            ws.set_column(ci, ci, None, None, {'hidden': True})

    # Configuração padronizada para impressão em A4.
    aplicar_configuracao_impressao_excel(ws, df, orientacao=orientacao_impressao)

    # Formatação condicional por status
    if col_alerta in df.columns:
        idx_al = df.columns.get_loc(col_alerta)
        letra  = _col_letter(idx_al)
        n_rows = len(df)
        n_cols = len(df.columns) - 1
        for status, (bg, fg) in EXCEL_CORES.items():
            fmt_c = fmt_status_cache.get(status)
            if fmt_c is None:
                fmt_c = wb.add_format({
                    'bg_color': bg, 'font_color': fg, 'align': 'center',
                    'valign': 'vcenter', 'text_wrap': True, 'bold': True,
                    'border': 1, 'border_color': '#D0D0D0', 'font_size': 10
                })
                fmt_status_cache[status] = fmt_c
            ws.conditional_format(1, 0, n_rows, n_cols, {
                'type': 'formula',
                'criteria': f'=${letra}2="{status}"',
                'format': fmt_c
            })
        # Regra por fórmula (como os demais status): pinta a linha inteira
        # quando o PARECER contém "Almoxarifado", sem afetar a coluna de Ação.
        ws.conditional_format(1, 0, n_rows, n_cols, {
            'type': 'formula',
            'criteria': f'=ISNUMBER(SEARCH("Almoxarifado",${letra}2))',
            'format': fmt_critico
        })


def normalizar_nome_aba_excel(nome_aba: str = "Dados") -> str:
    """Garante nome de aba compatível com Excel/XlsxWriter.

    Regras do Excel:
    - máximo de 31 caracteres;
    - não permite os caracteres especiais bloqueados pelo Excel;
    - não pode ficar vazio.
    """
    nome = str(nome_aba or "Dados").strip()
    for ch in ['[', ']', ':', '*', '?', '/', '\\']:
        nome = nome.replace(ch, '-')
    nome = nome.replace("'", "").strip()
    if not nome:
        nome = "Dados"
    return nome[:31]


def aplicar_configuracao_impressao_excel(
    ws,
    df: pd.DataFrame,
    orientacao: str = "landscape",
) -> None:
    """Padroniza a impressão dos relatórios Excel.

    Regras aplicadas a todos os relatórios:
    - papel A4;
    - cabeçalho da tabela repetido em todas as páginas impressas;
    - todas as colunas ajustadas para caber em uma página de largura;
    - altura livre, permitindo quantas páginas forem necessárias;
    - margens compactas e centralização horizontal.

    Orientação padrão: paisagem (horizontal).
    Exceção operacional: relatório de pedido classificado por categoria pode usar retrato (vertical).
    """
    try:
        orientacao_norm = clean(orientacao)
        if orientacao_norm in ("portrait", "retrato", "vertical"):
            ws.set_portrait()
        else:
            ws.set_landscape()

        # 9 = A4 no XlsxWriter/Excel.
        ws.set_paper(9)

        # Repete a linha de títulos das colunas em todas as páginas impressas.
        ws.repeat_rows(0)

        # Garante que todas as colunas caibam em uma única página de largura.
        # A altura fica livre para múltiplas páginas, mantendo o cabeçalho repetido.
        ws.fit_to_pages(1, 0)

        # Área de impressão restrita à tabela preenchida.
        if df is not None and not df.empty and len(df.columns) > 0:
            ws.print_area(0, 0, len(df), len(df.columns) - 1)

        # Margens compactas para melhorar legibilidade em A4.
        ws.set_margins(left=0.25, right=0.25, top=0.45, bottom=0.45)
        ws.center_horizontally()
        ws.set_footer('&C&P de &N')
    except Exception:
        # A configuração de impressão não deve impedir a geração do relatório.
        pass


@st.cache_data(show_spinner=False)
def exportar_excel_padronizado(
    df_dados: pd.DataFrame,
    nome_aba: str = "Dados",
    orientacao_impressao: str = "landscape",
) -> bytes:
    buf = io.BytesIO()
    nome_aba = normalizar_nome_aba_excel(nome_aba)
    with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
        df_dados.to_excel(wr, sheet_name=nome_aba, index=False)
        wb = wr.book
        ws = wr.sheets[nome_aba]

        fmt_b = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                               'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_t = wb.add_format({'align': 'left',   'valign': 'vcenter', 'text_wrap': True,
                               'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_h = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                               'text_wrap': True, 'bg_color': '#1E3A8A',
                               'font_color': '#FFFFFF', 'border': 1, 'font_size': 11})
        fmt_critico = wb.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911',
                                     'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                     'bold': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})

        col_alerta = find_col(df_dados, ['parecer', 'alerta', 'status']) or ''
        larguras   = {**LARGURAS_PADRAO_EXCEL}

        fmt_status_cache: dict = {}
        _aplicar_estilo_aba(ws, wb, df_dados, col_alerta, larguras,
                            fmt_h, fmt_b, fmt_t, fmt_status_cache, fmt_critico,
                            orientacao_impressao=orientacao_impressao)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def exportar_excel_multi_aba(df_total: pd.DataFrame, ordem_cols: list,
                              col_categoria: str, col_alerta: str,
                              larguras: dict, excluir_acoes: list = None,
                              ocultar_colunas: list | None = None,
                              ajustar_altura_linhas: bool = True,
                              orientacao_impressao: str = "landscape") -> bytes:
    buf = io.BytesIO()
    excluir_acoes = excluir_acoes or []
    ocultar_colunas = ocultar_colunas or []

    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        wb = writer.book

        # Formatos criados UMA única vez fora do loop de abas
        fmt_b = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                               'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_t = wb.add_format({'align': 'left',   'valign': 'vcenter', 'text_wrap': True,
                               'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})
        fmt_h = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                               'text_wrap': True, 'bg_color': '#1E3A8A',
                               'font_color': '#FFFFFF', 'border': 1, 'font_size': 11})
        fmt_critico = wb.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911',
                                     'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                     'bold': True, 'border': 1, 'border_color': '#D0D0D0', 'font_size': 10})

        # Cache de formatos por status — criado uma vez, reutilizado em todas as abas
        fmt_status_cache: dict = {}

        for cat, df_cat in df_total.groupby(col_categoria):
            df_exp = df_cat.copy()
            if excluir_acoes and 'Ação Logística Sugerida' in df_exp.columns:
                df_exp = df_exp[~df_exp['Ação Logística Sugerida'].isin(excluir_acoes)]

            cols_ok = [c for c in ordem_cols if c in df_exp.columns]
            df_exp  = df_exp[cols_ok].sort_values('Material').copy()
            if df_exp.empty:
                continue

            nome_aba = normalizar_nome_aba_excel(cat)
            df_exp.to_excel(writer, sheet_name=nome_aba, index=False)
            ws = writer.sheets[nome_aba]

            _aplicar_estilo_aba(ws, wb, df_exp, col_alerta, larguras,
                                fmt_h, fmt_b, fmt_t, fmt_status_cache, fmt_critico,
                                ajustar_altura_linhas=ajustar_altura_linhas,
                                ocultar_colunas=ocultar_colunas,
                                orientacao_impressao=orientacao_impressao)
    return buf.getvalue()


# =============================================================================
# OTIMIZAÇÃO DE DESEMPENHO — CACHE OPERACIONAL E RELATÓRIOS SOB DEMANDA
# =============================================================================

def _hash_dataframe_operacional(df: pd.DataFrame | None, cols: list[str] | None = None) -> str:
    """Gera assinatura leve para saber se um resultado precisa ser recalculado.

    Usado para evitar refazer remanejamentos/consolidações e também para invalidar
    relatórios em Excel quando filtros ou dados mudam. Não altera nenhuma regra de
    negócio; apenas controla reaproveitamento de resultados na sessão.
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return "empty"
    try:
        if cols:
            cols_ok = [c for c in cols if c in df.columns]
            dfx = df[cols_ok].copy() if cols_ok else df.copy()
        else:
            dfx = df.copy()
        # Mantém a assinatura estável sem depender de objetos complexos.
        dfx = dfx.replace([np.inf, -np.inf], np.nan)
        h = int(pd.util.hash_pandas_object(dfx, index=True).sum())
        return f"{dfx.shape}|{','.join(map(str, dfx.columns))}|{h}"
    except Exception:
        return f"fallback|{df.shape}|{','.join(map(str, df.columns))}"


def invalidar_cache_remanejamentos():
    """Remove resultados de remanejamento derivados da consolidação atual.

    Quando uma nova consolidação é gerada, os remanejamentos anteriores podem
    ficar defasados em relação aos novos saldos/CMDs. A invalidação centralizada
    evita que a aba Remanejamentos mostre resultado antigo e elimina o atraso de
    sincronização entre a geração da consolidação e a visualização dos remanejamentos.
    """
    chaves = [
        'df_remanejamento_geral_huufma',
        'assinatura_remanejamento_geral_huufma',
        'df_remanejamento_validade_huufma',
        'assinatura_remanejamento_validade_huufma',
        'remanejamento_geral_ultima_atualizacao',
        'remanejamento_validade_ultima_atualizacao',
    ]
    for k in chaves:
        st.session_state.pop(k, None)

    # Relatórios gerados a partir dos remanejamentos também precisam ser refeitos
    # quando a consolidação muda. Os demais relatórios continuam sendo invalidados
    # pela própria assinatura do dataframe na função sob demanda.
    for k in list(st.session_state.keys()):
        if k.startswith('relatorio_excel_remanejamento_'):
            st.session_state.pop(k, None)


def _limpar_resultados_derivados():
    """Remove caches dependentes dos arquivos/parâmetros quando há novo processamento."""
    prefixos = [
        'df_consolidado', 'assinatura_consolidacao_huufma',
        'df_remanejamento_geral_huufma', 'assinatura_remanejamento_geral_huufma',
        'df_remanejamento_validade_huufma', 'assinatura_remanejamento_validade_huufma',
        'remanejamento_geral_ultima_atualizacao', 'remanejamento_validade_ultima_atualizacao',
    ]
    for k in list(st.session_state.keys()):
        if k in prefixos or k.startswith('relatorio_excel_'):
            st.session_state.pop(k, None)


def obter_remanejamento_geral_session_cache(
    df_cons: pd.DataFrame,
    df_validades: pd.DataFrame | None = None,
    dias_cobertura: int | None = None,
    forcar_recalculo: bool = False,
) -> pd.DataFrame:
    """Reaproveita o remanejamento geral na Consolidação e na aba Remanejamentos.

    A mesma tabela calculada é usada em todos os pontos do app. Isso evita que a
    consolidação e a aba de remanejamentos façam o mesmo processamento pesado a
    cada rerun do Streamlit.
    """
    sig = "|".join([
        str(dias_cobertura),
        _hash_dataframe_operacional(df_cons, [
            'Código MV', 'Material', 'Categoria', 'Farmácia', 'Cód. Farmácia',
            'Saldo Atual', 'CMD', 'Necessidade', 'Saldo Central', 'Parecer'
        ]),
        _hash_dataframe_operacional(df_validades, [
            'key', 'Farmácia', 'Dias até Vencer', 'Saldo AGHU', 'Validade Fmt', 'Situação'
        ]),
    ])
    if (
        forcar_recalculo or
        st.session_state.get('assinatura_remanejamento_geral_huufma') != sig or
        'df_remanejamento_geral_huufma' not in st.session_state
    ):
        df_calc = calcular_remanejamento_equalizado(df_cons, df_validades, dias_cobertura=dias_cobertura)
        if not df_calc.empty and 'Material' in df_calc.columns:
            df_calc = df_calc.sort_values('Material', kind='mergesort').reset_index(drop=True)
        st.session_state['df_remanejamento_geral_huufma'] = df_calc
        st.session_state['assinatura_remanejamento_geral_huufma'] = sig
        st.session_state['remanejamento_geral_ultima_atualizacao'] = datetime.now().strftime('%d/%m/%Y %H:%M')
    return st.session_state.get('df_remanejamento_geral_huufma', pd.DataFrame()).copy()


def obter_remanejamento_validade_session_cache(
    df_validades: pd.DataFrame,
    df_cons: pd.DataFrame,
    forcar_recalculo: bool = False,
) -> pd.DataFrame:
    """Cache em sessão do remanejamento preventivo por validade."""
    sig = "|".join([
        _hash_dataframe_operacional(df_validades, [
            'key', 'Farmácia', 'Dias até Vencer', 'Saldo AGHU', 'Validade Fmt', 'Situação'
        ]),
        _hash_dataframe_operacional(df_cons, [
            'Código MV', 'Farmácia', 'Cód. Farmácia', 'Saldo Atual', 'CMD', 'Necessidade', 'Parecer'
        ]),
    ])
    if (
        forcar_recalculo or
        st.session_state.get('assinatura_remanejamento_validade_huufma') != sig or
        'df_remanejamento_validade_huufma' not in st.session_state
    ):
        df_calc = calcular_remanejamento_preventivo_validade(df_validades, df_cons)
        st.session_state['df_remanejamento_validade_huufma'] = df_calc
        st.session_state['assinatura_remanejamento_validade_huufma'] = sig
        st.session_state['remanejamento_validade_ultima_atualizacao'] = datetime.now().strftime('%d/%m/%Y %H:%M')
    return st.session_state.get('df_remanejamento_validade_huufma', pd.DataFrame()).copy()


def _assinatura_relatorio(df: pd.DataFrame, nome_aba: str, extra: str = "") -> str:
    return "|".join([nome_aba, extra, _hash_dataframe_operacional(df)])


def gerar_download_excel_sob_demanda(
    label: str,
    df_dados: pd.DataFrame,
    nome_aba: str,
    file_name: str,
    key: str,
    help: str | None = None,
    use_container_width: bool = True,
    orientacao_impressao: str = "landscape",
):
    """Gera o Excel apenas quando o usuário solicita, evitando travar filtros/abas."""
    key_sig = f"relatorio_excel_{key}_sig"
    key_bytes = f"relatorio_excel_{key}_bytes"
    key_name = f"relatorio_excel_{key}_file"
    sig = _assinatura_relatorio(df_dados, nome_aba, f"{file_name}|{orientacao_impressao}")
    if st.session_state.get(key_sig) != sig:
        st.session_state.pop(key_bytes, None)
        st.session_state[key_sig] = sig
        st.session_state[key_name] = file_name

    if st.button(
        label,
        key=f"relatorio_excel_{key}_preparar",
        use_container_width=use_container_width,
        help=help,
        type="primary",
    ):
        with st.spinner("📄 Gerando arquivo Excel..."):
            st.session_state[key_bytes] = exportar_excel_padronizado(
                df_dados.copy(), nome_aba, orientacao_impressao=orientacao_impressao
            )
            st.session_state[key_name] = file_name

    if key_bytes in st.session_state:
        st.download_button(
            "⬇️ Baixar relatório pronto",
            data=st.session_state[key_bytes],
            file_name=st.session_state.get(key_name, file_name),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"relatorio_excel_{key}_baixar",
            use_container_width=use_container_width,
            type="primary",
        )


def gerar_download_multi_aba_sob_demanda(
    label: str,
    df_total: pd.DataFrame,
    ordem_cols: list,
    col_categoria: str,
    col_alerta: str,
    larguras: dict,
    file_name: str,
    key: str,
    excluir_acoes: list | None = None,
    ocultar_colunas: list | None = None,
    ajustar_altura_linhas: bool = True,
    help: str | None = None,
    use_container_width: bool = True,
    orientacao_impressao: str = "landscape",
):
    """Versão sob demanda para relatórios Excel com múltiplas abas."""
    excluir_acoes = excluir_acoes or []
    ocultar_colunas = ocultar_colunas or []
    key_sig = f"relatorio_excel_{key}_sig"
    key_bytes = f"relatorio_excel_{key}_bytes"
    key_name = f"relatorio_excel_{key}_file"
    extra = repr((ordem_cols, col_categoria, col_alerta, larguras, excluir_acoes, ocultar_colunas, ajustar_altura_linhas, orientacao_impressao, file_name))
    sig = _assinatura_relatorio(df_total, "multi_aba", extra)
    if st.session_state.get(key_sig) != sig:
        st.session_state.pop(key_bytes, None)
        st.session_state[key_sig] = sig
        st.session_state[key_name] = file_name

    if st.button(
        label,
        key=f"relatorio_excel_{key}_preparar",
        use_container_width=use_container_width,
        help=help,
        type="primary",
    ):
        with st.spinner("📄 Gerando arquivo Excel..."):
            st.session_state[key_bytes] = exportar_excel_multi_aba(
                df_total.copy(), ordem_cols, col_categoria, col_alerta, larguras,
                excluir_acoes=excluir_acoes,
                ocultar_colunas=ocultar_colunas,
                ajustar_altura_linhas=ajustar_altura_linhas,
                orientacao_impressao=orientacao_impressao,
            )
            st.session_state[key_name] = file_name

    if key_bytes in st.session_state:
        st.download_button(
            "⬇️ Baixar relatório pronto",
            data=st.session_state[key_bytes],
            file_name=st.session_state.get(key_name, file_name),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"relatorio_excel_{key}_baixar",
            use_container_width=use_container_width,
            type="primary",
        )


# =============================================================================
# CONTROLE DE VALIDADE — FUNÇÕES
# =============================================================================

# Localização do arquivo no SharePoint pessoal do responsável
_SP_SITE    = "https://ebserhnet-my.sharepoint.com"
_SP_CAMINHO = "/personal/elton_freitas_ebserh_gov_br/Documents/Controle de Validade final.xlsx"


def carregar_validades_sharepoint() -> tuple[pd.DataFrame | None, str]:
    """Autentica com conta Outlook convidada e baixa a planilha de validades.
    Credenciais lidas dos Streamlit Secrets — nunca expostas no código.
    Retorna (DataFrame, '') em sucesso ou (None, mensagem_erro) em falha."""
    try:
        from office365.sharepoint.client_context import ClientContext
        from office365.runtime.auth.user_credential import UserCredential

        email = st.secrets.get("SHAREPOINT_EMAIL", "")
        senha = st.secrets.get("SHAREPOINT_PASSWORD", "")

        if not email or not senha:
            return None, (
                "Credenciais não configuradas. "
                "Acesse o painel do Streamlit → Settings → Secrets e adicione "
                "SHAREPOINT_EMAIL e SHAREPOINT_PASSWORD."
            )

        ctx = ClientContext(_SP_SITE).with_credentials(
            UserCredential(email, senha)
        )

        conteudo = io.BytesIO()
        (ctx.web
           .get_file_by_server_relative_url(_SP_CAMINHO)
           .download(conteudo)
           .execute_query())

        conteudo.seek(0)
        df = carregar_planilha_validades_excel(conteudo)

        if df.empty:
            return None, "O arquivo foi baixado mas está vazio."

        st.session_state['sharepoint_ultima_carga'] = datetime.now().strftime('%d/%m/%Y %H:%M')
        return df, ""

    except ImportError:
        return None, "Biblioteca Office365 não instalada. Execute: pip install Office365-REST-Python-Client"
    except Exception as e:
        msg = str(e)
        if '401' in msg or 'Unauthorized' in msg:
            return None, "Acesso negado (401). Verifique se a conta ainda tem acesso ao arquivo."
        if '403' in msg or 'Forbidden' in msg:
            return None, "Permissão negada (403). O arquivo pode ter sido movido ou o acesso revogado."
        if '404' in msg or 'NotFound' in msg:
            return None, "Arquivo não encontrado (404). Verifique se o caminho do arquivo mudou."
        if 'timeout' in msg.lower():
            return None, "Tempo esgotado. Verifique a conexão com o SharePoint."
        return None, f"Erro ao acessar o SharePoint: {msg}"


def _coluna_lote_validades(df: pd.DataFrame):
    """Evita confundir 'Quantos lotes disponíveis?' com número/identificação do lote."""
    candidatos = []
    for col in df.columns:
        c = clean(col)
        if any(t in c for t in ['lote', 'batch', 'lot']) and not any(
            f in c for f in ['quant', 'quanto', 'quantos', 'disponivel', 'disponiveis']
        ):
            candidatos.append(col)
    return candidatos[0] if candidatos else None


def _coluna_quantidade_validade(df: pd.DataFrame):
    """Localiza a quantidade informada na planilha de validade.

    Essa quantidade é usada apenas como fallback, pois o saldo oficial deve vir
    preferencialmente do arquivo de Estoque Geral do AGHU.
    """
    prioritarios = []
    gerais = []

    for col in df.columns:
        c = clean(col)
        if not any(t in c for t in ['qtd', 'qtde', 'quant', 'quantidade', 'saldo', 'estoque']):
            continue
        # Não confundir quantidade de lotes nem a coluna isolada de dias restantes
        # com quantidade em estoque/a expirar.
        if any(f in c for f in ['lote', 'lotes']) or c.startswith('dias') or 'dias restantes' in c:
            continue

        if any(t in c for t in ['expirar', 'vencer', 'validade', 'prazo']):
            prioritarios.append(col)
        else:
            gerais.append(col)

    return prioritarios[0] if prioritarios else (gerais[0] if gerais else None)


def extrair_saldos_estoque_geral(est_geral: pd.DataFrame) -> pd.DataFrame:
    """Extrai saldo atual por código + farmácia do arquivo de Estoque Geral do AGHU.

    Essa é a fonte oficial para decidir se um item vencido/a vencer ainda exige ação.
    Quando o Estoque Geral está disponível, itens com saldo AGHU igual a zero não devem
    aparecer no painel nem nos relatórios de validade.
    """
    if est_geral is None or est_geral.empty:
        return pd.DataFrame(columns=['key', 'Farmácia', 'Saldo AGHU'])

    c_cod = find_col(est_geral, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod'])
    c_qtd = find_col(est_geral, ['qtde disp', 'disponivel', 'saldo', 'quant'])
    c_alm = find_col(est_geral, ['almox'])

    if not all([c_cod, c_qtd, c_alm]):
        return pd.DataFrame(columns=['key', 'Farmácia', 'Saldo AGHU'])

    est = est_geral.copy()
    est['key'] = est[c_cod].apply(clean_key)
    est['Farmácia'] = est[c_alm].apply(clean_key)
    est['Saldo AGHU'] = p_num_series(est[c_qtd])

    # Regra global: o almoxarifado 45 representa estoque de pacientes e não deve
    # compor saldo AGHU, validade, pedido ou remanejamento.
    est = est[~est['Farmácia'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)]

    est = est[(est['key'] != '') & (est['Farmácia'] != '')]
    if est.empty:
        return pd.DataFrame(columns=['key', 'Farmácia', 'Saldo AGHU'])

    return (
        est.groupby(['key', 'Farmácia'], as_index=False)['Saldo AGHU']
        .sum()
    )


def aplicar_saldos_validades(df_val: pd.DataFrame, est_geral_raw: pd.DataFrame | None = None) -> pd.DataFrame:
    """Adiciona saldos separados ao painel de validade e aplica a regra operacional.

    Colunas geradas:
    - Saldo AGHU: saldo atual oficial extraído do Estoque Geral do AGHU.
    - Saldo Planilha Validade: quantidade informada no controle de validade da equipe.

    Regra de exibição:
    - Se o Estoque Geral do AGHU estiver disponível, somente permanecem itens com
      Saldo AGHU > 0 para a respectiva farmácia.
    - Se o Estoque Geral não estiver disponível, mantém a lista com base na planilha
      de validade, mas deixa Saldo AGHU em branco para evidenciar que falta confirmação.

    Sem cache proposital: esta função é chamada a cada renderização da aba Controle de
    Validade (guarda de segurança) recebendo DataFrames potencialmente grandes. O custo
    de hash do st.cache_data nesse caso supera o ganho, já que o resultado raramente é
    reaproveitado entre chamadas na mesma sessão.
    """
    if df_val is None or df_val.empty:
        return df_val

    df = df_val.copy()
    if 'Farmácia' in df.columns:
        df['Farmácia'] = df['Farmácia'].apply(clean_key)
        df = df[~df['Farmácia'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()

    # Compatibilidade com versões anteriores: a quantidade da planilha era chamada
    # de Qtd Controle Validade. Agora ela fica visível como Saldo Planilha Validade.
    if 'Saldo Planilha Validade' not in df.columns:
        if 'Qtd Controle Validade' in df.columns:
            df['Saldo Planilha Validade'] = df['Qtd Controle Validade']
        else:
            df['Saldo Planilha Validade'] = np.nan

    df['Saldo Planilha Validade'] = pd.to_numeric(
        df['Saldo Planilha Validade'], errors='coerce'
    )

    # Remove colunas antigas/derivadas antes de recalcular, evitando sufixos após merges.
    for c in [
        'Saldo Estoque Geral', 'Saldo Atual', 'Fonte Saldo', 'Saldo AGHU',
        'AGHU disponível para filtro'
    ]:
        if c in df.columns:
            df = df.drop(columns=c)

    saldos_aghu = extrair_saldos_estoque_geral(est_geral_raw)
    aghu_disponivel = not saldos_aghu.empty

    if aghu_disponivel:
        df = df.merge(saldos_aghu, on=['key', 'Farmácia'], how='left')
        # Se o Estoque Geral foi carregado e o item não apareceu para aquela farmácia,
        # operacionalmente o saldo atual é zero.
        df['Saldo AGHU'] = pd.to_numeric(df['Saldo AGHU'], errors='coerce').fillna(0)
        df = df[df['Saldo AGHU'] > 0].copy()
        df['AGHU disponível para filtro'] = True
    else:
        df['Saldo AGHU'] = pd.NA
        df['AGHU disponível para filtro'] = False

    # Int64 permite valores ausentes sem quebrar a exibição no Streamlit.
    df['Saldo AGHU'] = pd.to_numeric(df['Saldo AGHU'], errors='coerce').round().astype('Int64')
    df['Saldo Planilha Validade'] = pd.to_numeric(
        df['Saldo Planilha Validade'], errors='coerce'
    ).round().astype('Int64')

    return df


# Alias para manter compatibilidade com chamadas antigas dentro do arquivo.
# Mesmo motivo: sem cache, ver docstring de aplicar_saldos_validades.
def aplicar_saldo_atual_validades(df_val: pd.DataFrame, est_geral_raw: pd.DataFrame | None = None) -> pd.DataFrame:
    return aplicar_saldos_validades(df_val, est_geral_raw)


def normalizar_farmacia_validade(v) -> str:
    """Converte farmácia da planilha de validade para código de almoxarifado usado pelo app."""
    if pd.isna(v):
        return ""

    # Se já vier como código, mantém.
    cod = clean_key(v)
    if cod:
        return cod

    txt = clean(v)

    # Mapa principal com os nomes do próprio sistema
    mapa_nomes = {clean(nome): cod for cod, nome in DIC_NOMES_FARMACIAS.items()}

    # Apelidos observados na planilha de validade
    aliases = {
        "farmacia upd": "31",
        "farmacia dutra": "31",
        "farmacia unidade de dispensacao": "31",
        "farmacia unidade de dispensacao farmaceutica": "31",
        "farmacia centro cirurgico": "7",
        "centro cirurgico": "7",
        "farmacia umi": "13",
        "umi": "13",
        "farmacia uti": "34",
        "uti": "34",
        "farmacia oftalmologia": "39",
        "oftalmologia": "39",
        # Ambulatorial não está entre as farmácias-alvo atuais do app.
        # Mantém sem código para aparecer como "Não informado" ou ser tratado futuramente.
        "farmacia ambulatorial": "",
    }

    if txt in aliases:
        return aliases[txt]
    if txt in mapa_nomes:
        return mapa_nomes[txt]

    # Tenta correspondência parcial como último recurso.
    for nome_limpo, cod_mapa in mapa_nomes.items():
        if nome_limpo in txt or txt in nome_limpo:
            return cod_mapa

    return ""


def converter_validade_series(s: pd.Series) -> pd.Series:
    """Converte validade em formatos comuns: data, texto DD/MM/AAAA, ISO ou serial Excel."""
    dt = pd.to_datetime(s, dayfirst=True, errors='coerce')

    # Fallback para número serial do Excel, se houver.
    mask = dt.isna()
    if mask.any():
        nums = pd.to_numeric(s[mask], errors='coerce')
        mask_num = nums.notna()
        if mask_num.any():
            dt.loc[nums[mask_num].index] = pd.to_datetime(
                nums[mask_num], unit='D', origin='1899-12-30', errors='coerce'
            )

    return dt


def carregar_planilha_validades_excel(fonte) -> pd.DataFrame:
    """Lê a planilha de validade localizando automaticamente aba e linha de cabeçalho.

    Necessário porque o arquivo 'Controle de Validade final.xlsx' tem abas vazias
    antes da aba real, e o cabeçalho da aba de controle começa na linha 6.
    """
    if hasattr(fonte, "seek"):
        fonte.seek(0)

    xls = pd.ExcelFile(fonte)
    candidatos = []

    for aba in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=aba, header=None, dtype=str, nrows=20)
        if preview.dropna(how='all').empty:
            continue

        for idx, row in preview.iterrows():
            headers = [("" if pd.isna(x) else str(x).strip()) for x in row.tolist()]
            tmp = pd.DataFrame(columns=headers)

            c_cod = find_col(tmp, ['codigo', 'cod', 'mv', 'ca3'], forbidden=['material', 'descri'])
            c_val = find_col(tmp, ['valid', 'vencim', 'expir', 'data val'])

            if c_cod and c_val:
                candidatos.append((aba, idx))
                break

    if not candidatos:
        # Fallback: mantém comportamento antigo, mas provavelmente será rejeitado pela normalização.
        if hasattr(fonte, "seek"):
            fonte.seek(0)
        return pd.read_excel(fonte, dtype=str)

    # Dá preferência à aba de controle de validade, quando existir.
    aba, header_row = next(
        ((a, h) for a, h in candidatos if 'validade' in clean(a) or 'controle' in clean(a)),
        candidatos[0]
    )

    if hasattr(fonte, "seek"):
        fonte.seek(0)
    df = pd.read_excel(fonte, sheet_name=aba, header=header_row, dtype=str)
    return df.dropna(how='all').reset_index(drop=True)


def normalizar_planilha_validades(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza colunas da planilha de validades (SharePoint ou upload manual)."""
    c_cod = find_col(df, ['codigo', 'cod', 'mv', 'ca3'], forbidden=['material', 'descri'])
    c_mat = find_col(df, ['material', 'produto', 'descri', 'nome'])
    c_lot = _coluna_lote_validades(df)
    c_val = find_col(df, ['valid', 'vencim', 'expir', 'data val'])
    c_far = find_col(df, ['farmacia', 'almox', 'unidade', 'setor'])
    c_qtd = _coluna_quantidade_validade(df)

    if not c_cod or not c_val:
        return pd.DataFrame()

    out = pd.DataFrame()
    out['key']      = df[c_cod].apply(clean_key)
    out['Material'] = df[c_mat].fillna('').astype(str).str.strip() if c_mat else ''
    out['Lote']     = df[c_lot].fillna('').astype(str).str.strip() if c_lot else ''
    out['Validade'] = converter_validade_series(df[c_val])
    out['Farmácia'] = df[c_far].apply(normalizar_farmacia_validade) if c_far else ''
    out['Saldo Planilha Validade'] = p_num_series(df[c_qtd]) if c_qtd else np.nan
    # Mantém o nome antigo apenas por compatibilidade interna, se alguma sessão antiga reutilizar o dado.
    out['Qtd Controle Validade'] = out['Saldo Planilha Validade']

    # Visibilidade operacional: registra quantos registros tinham linha de validade válida
    # (código + data preenchidos), mas a farmácia não pôde ser mapeada para um código conhecido
    # (ex.: "Farmácia Ambulatorial", erro de digitação, célula vazia). Esses registros são
    # descartados silenciosamente pelo filtro abaixo; o contador fica disponível para a UI
    # avisar o usuário, sem mudar o comportamento de filtragem já validado.
    mask_base_valida = (out['key'] != '') & out['Validade'].notna()
    mask_farmacia_vazia = mask_base_valida & (out['Farmácia'].astype(str).str.strip() == '')
    n_descartados_farmacia = int(mask_farmacia_vazia.sum())
    if n_descartados_farmacia > 0:
        st.session_state['validade_descartados_farmacia_nao_mapeada'] = n_descartados_farmacia
    else:
        st.session_state.pop('validade_descartados_farmacia_nao_mapeada', None)

    # Regra global: remover almoxarifados fora da gestão do aplicativo.
    out = out[~out['Farmácia'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)]

    return out[out['key'] != ''].dropna(subset=['Validade']).reset_index(drop=True)


def extrair_validades_aghu(est_geral: pd.DataFrame) -> pd.DataFrame:
    """Extrai colunas de lote e validade do AGDA2, quando preenchidas."""
    c_cod = find_col(est_geral, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod'])
    c_lot = find_col(est_geral, ['lote', 'batch'])
    c_val = find_col(est_geral, ['valid', 'vencim', 'expir'])
    c_alm = find_col(est_geral, ['almox'])
    c_mat = find_col(est_geral, ['material', 'produto', 'descri'])

    if not c_cod or not (c_lot or c_val):
        return pd.DataFrame()

    df = est_geral.copy()
    out = pd.DataFrame()
    out['key']      = df[c_cod].apply(clean_key)
    out['Material'] = df[c_mat].fillna('').astype(str).str.strip() if c_mat else ''
    out['Lote']     = df[c_lot].fillna('').astype(str).str.strip() if c_lot else ''
    out['Validade'] = pd.to_datetime(df[c_val], dayfirst=True, errors='coerce') if c_val else pd.NaT
    out['Farmácia'] = df[c_alm].apply(clean_key) if c_alm else ''
    out['Fonte']    = 'AGHU'

    # Regra global: remover almoxarifados fora da gestão do aplicativo.
    out = out[~out['Farmácia'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)]

    # Só retorna linhas com lote ou validade efetivamente preenchidos
    mask = out['Lote'].ne('') | out['Validade'].notna()
    return out[mask & out['key'].ne('')].reset_index(drop=True)


def mesclar_validades(df_aghu: pd.DataFrame, df_sp: pd.DataFrame) -> pd.DataFrame:
    """Mescla validades: AGHU prevalece; SharePoint complementa os vazios.
    Regra: se AGHU tem validade para key+farmácia → usa AGHU. Senão → SharePoint."""
    if df_aghu.empty and df_sp.empty:
        return pd.DataFrame()

    if df_sp.empty:
        df_aghu['Fonte'] = 'AGHU'
        return df_aghu

    df_sp_norm = df_sp.copy()
    df_sp_norm['Fonte'] = 'Planilha Equipe'

    if df_aghu.empty:
        return df_sp_norm

    # Chave composta: item + farmácia
    chave_aghu = set(zip(df_aghu['key'], df_aghu['Farmácia']))
    mask_novos = ~df_sp_norm.apply(
        lambda r: (r['key'], r['Farmácia']) in chave_aghu, axis=1
    )
    return pd.concat([df_aghu, df_sp_norm[mask_novos]], ignore_index=True)


def classificar_validade(val, hoje: date) -> tuple[str, str]:
    """Retorna (emoji_semaforo, faixa_texto) para uma data de validade."""
    if pd.isna(val):
        return "⚫", "Sem data"
    dias = (val.date() - hoje).days
    if dias < 0:
        return "💀", "VENCIDO"
    if dias <= VALIDADE_CRITICO_DIAS:
        return "🔴", f"Crítico ({dias}d)"
    if dias <= VALIDADE_ATENCAO_DIAS:
        return "🟡", f"Atenção ({dias}d)"
    return "🟢", f"OK ({dias}d)"


def preparar_painel_validades(
    df_val: pd.DataFrame,
    mapa_cat: dict,
    est_geral_raw: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Monta o DataFrame final do painel de validades com semáforo, categoria e saldo atual."""
    hoje = datetime.now().date()
    df = aplicar_saldo_atual_validades(df_val.copy(), est_geral_raw)
    df['Categoria'] = df['key'].map(mapa_cat).fillna('OUTROS')
    df['Nome Farmácia'] = df['Farmácia'].map(DIC_NOMES_FARMACIAS).fillna(
        df['Farmácia'].apply(lambda x: f'Almox {x}' if x else 'Não informado')
    )
    semaforo = df['Validade'].apply(lambda v: classificar_validade(v, hoje))
    df['🚦'] = [s[0] for s in semaforo]
    df['Situação'] = [s[1] for s in semaforo]
    df['Dias até Vencer'] = df['Validade'].apply(
        lambda v: (v.date() - hoje).days if pd.notna(v) else None
    )
    df['Validade Fmt'] = df['Validade'].dt.strftime('%d/%m/%Y').fillna('—')

    ordem_situacao = {'VENCIDO': 0, 'Crítico': 1, 'Atenção': 2, 'OK': 3, 'Sem data': 4}
    df['_ord'] = df['Situação'].apply(lambda s: next(
        (v for k, v in ordem_situacao.items() if k in s), 5
    ))
    return df.sort_values(['_ord', 'Dias até Vencer']).drop(columns='_ord')


def processar_validades_para_sessao(df_sp_raw: pd.DataFrame, origem: str = "Planilha de validade") -> tuple[bool, str]:
    """Normaliza, mescla e salva as validades no session_state.

    Mantém a lógica já validada do app, mas centraliza o carregamento para que
    a Central de Processamento possa alimentar o painel de validade, o pedido e
    os remanejamentos FEFO no mesmo fluxo operacional.
    """
    if df_sp_raw is None or df_sp_raw.empty:
        return False, "A fonte de validade está vazia."

    df_sp_norm = normalizar_planilha_validades(df_sp_raw)
    if df_sp_norm.empty:
        return False, (
            "A planilha de validades não pôde ser interpretada. "
            "Verifique se há colunas identificáveis de Código MV e Validade."
        )

    df_aghu_val = pd.DataFrame()
    if 'est_geral_raw' in st.session_state:
        df_aghu_val = extrair_validades_aghu(st.session_state['est_geral_raw'])

    df_val_final = mesclar_validades(df_aghu_val, df_sp_norm)
    mapa_cat_val = obter_mapa_categorias()
    est_geral_para_saldo = st.session_state.get('est_geral_raw')
    painel_validades = preparar_painel_validades(
        df_val_final, mapa_cat_val, est_geral_para_saldo
    )

    st.session_state['df_validades_mescladas'] = painel_validades
    st.session_state['validade_ultima_origem'] = origem
    st.session_state['validade_ultima_carga'] = datetime.now().strftime('%d/%m/%Y %H:%M')

    n_aghu = len(df_aghu_val)
    n_sp = len(df_sp_norm)
    n_painel = len(painel_validades)
    tem_estoque_aghu = (
        est_geral_para_saldo is not None and
        isinstance(est_geral_para_saldo, pd.DataFrame) and
        not est_geral_para_saldo.empty
    )
    complemento = (
        f" Após filtro por Saldo AGHU > 0, permanecem {n_painel} registros no painel."
        if tem_estoque_aghu else
        f" {n_painel} registros foram mantidos no painel provisório, sem filtro por AGHU."
    )
    return True, (
        f"Validades carregadas de {origem}: {n_aghu} registros extraídos do AGHU + "
        f"{n_sp} registros da planilha da equipe → {len(df_val_final)} registros candidatos."
        + complemento
    )


def calcular_remanejamento_preventivo_validade(
    df_validades: pd.DataFrame,
    df_cons: pd.DataFrame,
    dias_max: int = VALIDADE_ATENCAO_DIAS,
) -> pd.DataFrame:
    """Identifica oportunidades de remanejamento preventivo por validade entre farmácias.

    Objetivo: localizar item com validade próxima em uma farmácia de origem e apontar
    outra farmácia que consome o mesmo item e não possui alerta de validade para ele.

    Regras principais:
    - Entram apenas itens NÃO vencidos, com vencimento em até ``dias_max`` dias.
    - A origem precisa ter Saldo AGHU > 0 na respectiva farmácia.
    - O destino precisa ser outra farmácia, com CMD > 0.
    - O destino não pode ter o mesmo item em alerta de validade até ``dias_max`` dias.
    - A quantidade sugerida é limitada pelo saldo em risco na origem e pela capacidade
      estimada de consumo do destino até a data de validade.

    Sem decorator de cache: o reaproveitamento de resultado já é controlado por
    obter_remanejamento_validade_session_cache, que mantém o resultado em session_state
    por assinatura de dados. Ter os dois caches simultâneos seria redundante.
    """
    if df_validades is None or df_validades.empty or df_cons is None or df_cons.empty:
        return pd.DataFrame()

    obrig_val = {'key', 'Farmácia', 'Validade'}
    obrig_cons = {'Código MV', 'Cód. Farmácia', 'Farmácia', 'CMD', 'Saldo Atual'}
    if not obrig_val.issubset(df_validades.columns) or not obrig_cons.issubset(df_cons.columns):
        return pd.DataFrame()

    hoje = datetime.now().date()
    dfv = df_validades.copy()
    dfv['key'] = dfv['key'].apply(clean_key)
    dfv['Farmácia'] = dfv['Farmácia'].apply(clean_key)
    dfv = dfv[~dfv['Farmácia'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()
    dfv['Validade'] = pd.to_datetime(dfv['Validade'], errors='coerce', dayfirst=True)

    if 'Dias até Vencer' not in dfv.columns:
        dfv['Dias até Vencer'] = dfv['Validade'].apply(
            lambda v: (v.date() - hoje).days if pd.notna(v) else np.nan
        )
    dfv['Dias até Vencer'] = pd.to_numeric(dfv['Dias até Vencer'], errors='coerce')

    if 'Saldo AGHU' not in dfv.columns:
        dfv['Saldo AGHU'] = 0
    dfv['Saldo AGHU'] = pd.to_numeric(dfv['Saldo AGHU'], errors='coerce').fillna(0)

    if 'Material' not in dfv.columns:
        dfv['Material'] = ''
    if 'Categoria' not in dfv.columns:
        dfv['Categoria'] = 'OUTROS'
    if 'Nome Farmácia' not in dfv.columns:
        dfv['Nome Farmácia'] = dfv['Farmácia'].map(DIC_NOMES_FARMACIAS).fillna(dfv['Farmácia'])
    if 'Validade Fmt' not in dfv.columns:
        dfv['Validade Fmt'] = dfv['Validade'].dt.strftime('%d/%m/%Y').fillna('—')
    if 'Situação' not in dfv.columns:
        semaforo = dfv['Validade'].apply(lambda v: classificar_validade(v, hoje))
        dfv['Situação'] = [s[1] for s in semaforo]

    # Para remanejamento preventivo, vencidos não entram: vencido deve ser segregado/retirado.
    df_risco = dfv[
        (dfv['Dias até Vencer'] >= 0) &
        (dfv['Dias até Vencer'] <= dias_max) &
        (dfv['Saldo AGHU'] > 0) &
        (dfv['key'] != '') &
        (dfv['Farmácia'] != '')
    ].copy()
    if df_risco.empty:
        return pd.DataFrame()

    # Conjunto de farmácias que também têm o mesmo item a vencer; elas não são bons destinos.
    chaves_em_risco = set(zip(df_risco['key'], df_risco['Farmácia']))

    dfc = df_cons.copy()
    dfc['key'] = dfc['Código MV'].apply(clean_key)
    dfc['cod_farm'] = dfc['Cód. Farmácia'].apply(clean_key)
    dfc = dfc[~dfc['cod_farm'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()
    dfc['CMD'] = pd.to_numeric(dfc['CMD'], errors='coerce').fillna(0)
    dfc['Saldo Atual'] = pd.to_numeric(dfc['Saldo Atual'], errors='coerce').fillna(0)
    if 'Necessidade' not in dfc.columns:
        dfc['Necessidade'] = 0
    dfc['Necessidade'] = pd.to_numeric(dfc['Necessidade'], errors='coerce').fillna(0)
    if 'Parecer' not in dfc.columns:
        dfc['Parecer'] = ''

    origem_consumo = dfc[['key', 'cod_farm', 'CMD', 'Saldo Atual']].rename(columns={
        'cod_farm': 'Farmácia',
        'CMD': 'CMD Origem',
        'Saldo Atual': 'Saldo Atual Origem Consolidado',
    })

    base = df_risco.merge(origem_consumo, on=['key', 'Farmácia'], how='left')
    base['CMD Origem'] = pd.to_numeric(base['CMD Origem'], errors='coerce').fillna(0)

    destinos = dfc[['key', 'cod_farm', 'Farmácia', 'CMD', 'Saldo Atual', 'Necessidade', 'Parecer']].rename(columns={
        'cod_farm': 'Farmácia Destino Código',
        'Farmácia': 'Transferir PARA',
        'CMD': 'CMD Destino',
        'Saldo Atual': 'Saldo Destino',
        'Necessidade': 'Necessidade Destino',
        'Parecer': 'Parecer Destino',
    })

    op = base.merge(destinos, on='key', how='left')
    op = op[op['Farmácia Destino Código'].notna()].copy()
    op = op[op['Farmácia Destino Código'] != op['Farmácia']].copy()

    # Destino não deve estar em alerta de validade para o mesmo item.
    op['_destino_tem_risco_validade'] = op.apply(
        lambda r: (r['key'], r['Farmácia Destino Código']) in chaves_em_risco,
        axis=1
    )
    op = op[~op['_destino_tem_risco_validade']].copy()

    op['CMD Destino'] = pd.to_numeric(op['CMD Destino'], errors='coerce').fillna(0)
    op['Saldo Destino'] = pd.to_numeric(op['Saldo Destino'], errors='coerce').fillna(0)
    op['Necessidade Destino'] = pd.to_numeric(op['Necessidade Destino'], errors='coerce').fillna(0)

    # Só faz sentido transferir para farmácia que consome o item.
    op = op[op['CMD Destino'] > 0].copy()
    if op.empty:
        return pd.DataFrame()

    op['Consumo Origem até Validade'] = np.floor(op['CMD Origem'] * op['Dias até Vencer']).clip(lower=0)
    op['Qtd em Risco na Origem'] = np.ceil(
        np.maximum(op['Saldo AGHU'] - op['Consumo Origem até Validade'], 0)
    )
    op['Consumo Possível no Destino até Validade'] = np.floor(
        op['CMD Destino'] * op['Dias até Vencer']
    ).clip(lower=0)
    op['Qtd Sugerida Remanejar'] = np.minimum(
        op['Qtd em Risco na Origem'],
        op['Consumo Possível no Destino até Validade']
    )

    op = op[op['Qtd Sugerida Remanejar'] > 0].copy()
    if op.empty:
        return pd.DataFrame()

    def _prioridade(row):
        parecer_destino = str(row.get('Parecer Destino', ''))
        if row['Dias até Vencer'] <= VALIDADE_CRITICO_DIAS:
            return 'Alta — vence em até 30 dias'
        if row['Necessidade Destino'] > 0 or any(t in parecer_destino for t in ['Desabastecimento', 'Remanejar', 'Almoxarifado', 'Solicitar']):
            return 'Alta — destino com necessidade'
        return 'Preventiva — destino consome o item'

    op['Prioridade'] = op.apply(_prioridade, axis=1)
    op['Transferir DE'] = op['Nome Farmácia']
    op['Validade'] = op['Validade Fmt']
    op['Justificativa'] = op.apply(
        lambda r: (
            f"Origem possui {int(r['Saldo AGHU'])} un. com validade em {int(r['Dias até Vencer'])} dia(s) "
            f"e CMD origem {r['CMD Origem']:.0f}. Destino consome {r['CMD Destino']:.0f} un./dia "
            f"e não consta com alerta de validade para este item."
        ),
        axis=1
    )

    cols = [
        'Prioridade', 'Código MV', 'Material', 'Categoria',
        'Transferir DE', 'Situação', 'Validade', 'Dias até Vencer',
        'Saldo AGHU', 'CMD Origem', 'Qtd em Risco na Origem',
        'Transferir PARA', 'CMD Destino', 'Saldo Destino', 'Necessidade Destino',
        'Consumo Possível no Destino até Validade', 'Qtd Sugerida Remanejar',
        'Parecer Destino', 'Justificativa'
    ]
    op = op.rename(columns={'key': 'Código MV'})

    for c in cols:
        if c not in op.columns:
            op[c] = ''

    # Evita múltiplas sugestões idênticas; mantém a melhor opção por origem+destino+item.
    op = op.sort_values(
        ['Dias até Vencer', 'Qtd Sugerida Remanejar', 'CMD Destino'],
        ascending=[True, False, False]
    ).drop_duplicates(['Código MV', 'Transferir DE', 'Transferir PARA'], keep='first')

    col_num = [
        'Dias até Vencer', 'Saldo AGHU', 'CMD Origem', 'Qtd em Risco na Origem',
        'CMD Destino', 'Saldo Destino', 'Necessidade Destino',
        'Consumo Possível no Destino até Validade', 'Qtd Sugerida Remanejar'
    ]
    for c in col_num:
        op[c] = pd.to_numeric(op[c], errors='coerce').fillna(0).round().astype(int)

    return op[cols].sort_values(
        ['Prioridade', 'Dias até Vencer', 'Qtd Sugerida Remanejar'],
        ascending=[True, True, False]
    ).reset_index(drop=True)


def _cobertura_texto(saldo, cmd) -> str:
    """Formata cobertura em dias de forma segura para relatórios gerenciais."""
    try:
        saldo = float(saldo)
    except Exception:
        saldo = 0.0
    try:
        cmd = float(cmd)
    except Exception:
        cmd = 0.0
    if cmd <= 0:
        return "+∞" if saldo > 0 else "Sem consumo"
    dias = saldo / cmd
    if not np.isfinite(dias):
        return "Sem dado"
    return f"{dias:.1f} dia(s)"


def calcular_remanejamento_equalizado(df_cons: pd.DataFrame,
                                      df_validades: pd.DataFrame | None = None,
                                      dias_cobertura: int | None = None) -> pd.DataFrame:
    """Sugere remanejamentos gerais apenas em cenário de contingência real.

    Esta análise NÃO equaliza todos os estoques indiscriminadamente. Ela só entra
    quando há, para o mesmo Código MV:
      1) pelo menos uma farmácia com consumo real e necessidade por consumo;
      2) saldo dos almoxarifados fornecedores zerado ou insuficiente para cobrir
         a necessidade agregada calculada por consumo, sem usar estoque mínimo;
      3) saldo disponível em outra farmácia satélite.

    Quando esses gatilhos são atendidos, o saldo existente nas farmácias é
    redistribuído de forma proporcional ao CMD, buscando que as farmácias
    consumidoras fiquem com coberturas semelhantes. Farmácias sem consumo
    recente têm estoque-alvo zero e podem doar todo o saldo.

    Sem decorator de cache: o reaproveitamento de resultado já é controlado por
    obter_remanejamento_geral_session_cache, que mantém o resultado em session_state
    por assinatura de dados. Ter os dois caches simultâneos seria redundante.
    """
    if df_cons is None or df_cons.empty:
        return pd.DataFrame()

    obrig = {'Código MV', 'Material', 'Categoria', 'Farmácia', 'Cód. Farmácia', 'Saldo Atual', 'CMD'}
    if not obrig.issubset(df_cons.columns):
        return pd.DataFrame()

    try:
        dias_cobertura = int(dias_cobertura or 15)
    except Exception:
        dias_cobertura = 15
    dias_cobertura = max(dias_cobertura, 1)

    df = df_cons.copy()
    df['key'] = df['Código MV'].apply(clean_key)
    df['cod_farm'] = df['Cód. Farmácia'].apply(clean_key)
    df = df[~df['cod_farm'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()
    if df.empty:
        return pd.DataFrame()
    df['Saldo Atual'] = (
        pd.to_numeric(df['Saldo Atual'], errors='coerce')
        .replace([np.inf, -np.inf], 0).fillna(0).clip(lower=0)
    )
    df['CMD'] = (
        pd.to_numeric(df['CMD'], errors='coerce')
        .replace([np.inf, -np.inf], 0).fillna(0).clip(lower=0)
    )
    if 'Parecer' not in df.columns:
        df['Parecer'] = ''
    if 'Saldo Central' not in df.columns:
        df['Saldo Central'] = 0

    df['Saldo Central'] = (
        pd.to_numeric(df['Saldo Central'], errors='coerce')
        .replace([np.inf, -np.inf], 0).fillna(0).clip(lower=0)
    )

    # Necessidade REAL para esta análise: baseada apenas no consumo médio diário
    # e no período de cobertura desejado. Não usa estoque mínimo, para evitar
    # sugerir remanejamento para farmácia sem demanda real.
    df['Necessidade consumo real'] = np.where(
        df['CMD'] > 0,
        np.ceil(np.maximum((df['CMD'] * dias_cobertura) - df['Saldo Atual'], 0)),
        0
    )
    df['Necessidade consumo real'] = (
        pd.to_numeric(df['Necessidade consumo real'], errors='coerce')
        .replace([np.inf, -np.inf], 0).fillna(0).clip(lower=0).astype(int)
    )

    # Consolidar eventual duplicidade por item+farmácia antes de calcular a distribuição.
    df = (df.groupby(['key', 'cod_farm'], as_index=False)
            .agg({
                'Código MV': 'first',
                'Material': 'first',
                'Categoria': 'first',
                'Farmácia': 'first',
                'Saldo Atual': 'sum',
                'CMD': 'max',
                'Necessidade consumo real': 'max',
                'Parecer': 'first',
                'Saldo Central': 'max',
            }))

    # Chaves com alerta de validade para observação, sem excluir automaticamente da análise geral.
    chaves_validade = set()
    if df_validades is not None and not df_validades.empty:
        dfv = df_validades.copy()
        if {'key', 'Farmácia'}.issubset(dfv.columns):
            dfv['key'] = dfv['key'].apply(clean_key)
            dfv['Farmácia'] = dfv['Farmácia'].apply(clean_key)
            dfv = dfv[~dfv['Farmácia'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()
            if 'Dias até Vencer' in dfv.columns:
                dias = pd.to_numeric(dfv['Dias até Vencer'], errors='coerce')
                mask = dias.between(0, VALIDADE_ATENCAO_DIAS, inclusive='both')
            else:
                mask = pd.Series(True, index=dfv.index)
            if 'Saldo AGHU' in dfv.columns:
                saldo_v = pd.to_numeric(dfv['Saldo AGHU'], errors='coerce').fillna(0)
                mask = mask & (saldo_v > 0)
            chaves_validade = set(zip(dfv.loc[mask, 'key'], dfv.loc[mask, 'Farmácia']))

    # Ordem de prioridade ancorada nas chaves conhecidas de MAPA_STATUS, em vez de
    # comparação solta por substring. Reduz o risco de quebra silenciosa caso o texto
    # do parecer mude (acento, espaço extra, etc.) em alguma chamada futura.
    # "Estoque Crítico no Almoxarifado X" é gerado dinamicamente com o nome do
    # almoxarifado central (ver definir_alerta_e_acao), por isso mantém checagem
    # por prefixo fixo para esse caso específico.
    _ORDEM_PRIORIDADE_PARECER = {
        "Desabastecimento Crítico": 0,
        "Remanejar": 1,
        "Estoque Crítico CAF": 2,
        "Solicitar": 3,
    }
    _PREFIXO_ALMOX_CRITICO = "Estoque Crítico no Almoxarifado"

    def _prioridade_parecer(parecer: str) -> int:
        p = str(parecer).strip()
        if p in _ORDEM_PRIORIDADE_PARECER:
            return _ORDEM_PRIORIDADE_PARECER[p]
        if p.startswith(_PREFIXO_ALMOX_CRITICO):
            return 2
        return 4

    resultados = []

    for cod, g in df.groupby('key'):
        g = g.copy().reset_index(drop=True)
        if cod == '' or g.empty:
            continue

        # Gatilho 1: precisa haver demanda real em pelo menos uma farmácia consumidora.
        # A necessidade não usa estoque mínimo; usa apenas CMD x dias_cobertura.
        g['precisa_receber'] = (g['CMD'] > 0) & (g['Necessidade consumo real'] > 0)
        if not g['precisa_receber'].any():
            continue

        necessidade_total = int(math.ceil(float(g.loc[g['precisa_receber'], 'Necessidade consumo real'].sum())))
        if necessidade_total <= 0:
            continue

        saldo_central = int(round(float(g['Saldo Central'].max())))
        central_insuficiente = saldo_central < necessidade_total

        # Gatilho 2: se os almoxarifados conseguem atender a necessidade agregada real,
        # a ação logística preferencial continua sendo solicitar ao almoxarifado.
        if not central_insuficiente:
            continue

        saldo_total_farmacias = int(round(float(g['Saldo Atual'].sum())))
        cmd_total_consumidoras = float(g.loc[g['CMD'] > 0, 'CMD'].sum())

        # Sem estoque nas farmácias ou sem consumo recente não há redistribuição segura.
        if saldo_total_farmacias <= 0 or cmd_total_consumidoras <= 0:
            continue

        # Gatilho 3: precisa existir saldo disponível em alguma farmácia satélite.
        if not (g['Saldo Atual'] > 0).any():
            continue

        cobertura_alvo = saldo_total_farmacias / cmd_total_consumidoras
        if not np.isfinite(cobertura_alvo) or cobertura_alvo <= 0:
            continue

        # Estoque ideal inteiro preservando o saldo total: floor + distribuição por maior fração.
        # Farmácias com CMD=0 ficam com alvo zero. Farmácias com consumo entram na conta para
        # preservar a assistência, mesmo se naquele momento não estão abaixo do parâmetro de pedido.
        g['ideal_raw'] = np.where(g['CMD'] > 0, g['CMD'] * cobertura_alvo, 0.0)
        g['ideal_floor'] = np.floor(g['ideal_raw']).astype(int)
        g.loc[g['CMD'] <= 0, 'ideal_floor'] = 0
        g['frac'] = g['ideal_raw'] - np.floor(g['ideal_raw'])

        restante = int(saldo_total_farmacias - g['ideal_floor'].sum())
        g['estoque_ideal'] = g['ideal_floor'].astype(int)
        if restante > 0:
            idx_consumo = g[g['CMD'] > 0].sort_values('frac', ascending=False).index.tolist()
            if idx_consumo:
                for idx in idx_consumo[:restante]:
                    g.loc[idx, 'estoque_ideal'] += 1

        g['saldo_int'] = g['Saldo Atual'].round().astype(int).clip(lower=0)
        g['delta'] = g['estoque_ideal'] - g['saldo_int']
        g['cobertura_atual_num'] = np.where(g['CMD'] > 0, g['Saldo Atual'] / g['CMD'], np.inf)

        donors = []
        receivers = []
        for _, r in g.iterrows():
            saldo = int(r['saldo_int'])
            cmd = float(r['CMD'])
            delta = int(round(float(r['delta'])))
            cobertura_atual = float(r['cobertura_atual_num']) if np.isfinite(r['cobertura_atual_num']) else np.inf

            item = {
                'cod_farm': r['cod_farm'],
                'farmacia': r['Farmácia'],
                'saldo': saldo,
                'cmd': cmd,
                'parecer': r.get('Parecer', ''),
                'necessidade_consumo_real': int(r.get('Necessidade consumo real', 0) or 0),
                'precisa_receber': bool(r.get('precisa_receber', False)),
                'cobertura_atual': cobertura_atual,
                'estoque_ideal': int(round(float(r['estoque_ideal']))),
            }

            # Receptores: farmácias consumidoras abaixo do estoque ideal após redistribuição.
            # O gatilho do item já exige que exista necessidade real em pelo menos uma farmácia.
            if cmd > 0 and delta > 0:
                item['necessidade_equalizar'] = int(delta)
                receivers.append(item)

            # Doadoras: farmácia sem consumo doa todo o saldo; farmácia com consumo doa
            # apenas o excedente acima do estoque ideal.
            if saldo > 0:
                if cmd <= 0:
                    excedente = saldo
                else:
                    excedente = max(0, saldo - int(item['estoque_ideal']))
                if excedente > 0:
                    item['excedente'] = int(excedente)
                    donors.append(item)

        if not donors or not receivers:
            continue

        donors.sort(key=lambda d: (0 if d['cmd'] <= 0 else 1, -d['cobertura_atual'], -d['excedente']))
        receivers.sort(key=lambda d: (_prioridade_parecer(d['parecer']), d['cobertura_atual'], -d['necessidade_equalizar']))

        for rec in receivers:
            necessidade_restante = int(rec['necessidade_equalizar'])
            saldo_destino_corrente = int(rec['saldo'])
            if necessidade_restante <= 0:
                continue

            for donor in donors:
                if necessidade_restante <= 0:
                    break
                if donor['cod_farm'] == rec['cod_farm']:
                    continue
                excedente_disponivel = int(donor.get('excedente', 0))
                if excedente_disponivel <= 0:
                    continue

                qtd = int(min(excedente_disponivel, necessidade_restante))
                if qtd <= 0:
                    continue

                cobertura_destino_antes_num = (saldo_destino_corrente / rec['cmd']) if rec['cmd'] > 0 else np.nan
                cobertura_destino_apos = (saldo_destino_corrente + qtd) / rec['cmd'] if rec['cmd'] > 0 else np.nan
                obs_validade = ''
                if (cod, donor['cod_farm']) in chaves_validade:
                    obs_validade = 'Origem também possui alerta de validade; avaliar prioridade FEFO.'
                elif (cod, rec['cod_farm']) in chaves_validade:
                    obs_validade = 'Destino possui alerta de validade para este item; avaliar antes de transferir.'

                if donor['cmd'] <= 0:
                    motivo_origem = 'a origem não apresentou consumo recente e pode doar o saldo disponível'
                else:
                    motivo_origem = (
                        f'a origem preserva estoque proporcional ao seu consumo, mantendo alvo aproximado de '
                        f'{donor["estoque_ideal"]} un.'
                    )

                cobertura_origem_txt = _cobertura_texto(donor['saldo'], donor['cmd'])
                cobertura_destino_txt = _cobertura_texto(saldo_destino_corrente, rec['cmd'])
                justificativa = (
                    f"Almoxarifados fornecedores com saldo insuficiente ({saldo_central} un. para "
                    f"{necessidade_total} un. de necessidade real por consumo em {dias_cobertura} dia(s)). "
                    f"Origem: {donor['farmacia']} possui {donor['saldo']} un., CMD {donor['cmd']:.2f} un./dia, "
                    f"cobertura {cobertura_origem_txt} e excedente remanejável de {excedente_disponivel} un.; "
                    f"{motivo_origem}. Destino: {rec['farmacia']} possui {saldo_destino_corrente} un., "
                    f"CMD {rec['cmd']:.2f} un./dia, cobertura {cobertura_destino_txt} e necessidade para equalizar "
                    f"{necessidade_restante} un. Sugere-se transferir {qtd} un. para estimar cobertura de "
                    f"{cobertura_destino_apos:.1f} dia(s), aproximando-se da cobertura hospitalar alvo de "
                    f"{cobertura_alvo:.1f} dia(s)."
                )
                if obs_validade:
                    justificativa += f" {obs_validade}"

                resultados.append({
                    'Código MV': str(g['Código MV'].iloc[0]),
                    'Material': str(g['Material'].iloc[0]),
                    'Categoria': str(g['Categoria'].iloc[0]),
                    'Saldo almoxarifados fornecedores': saldo_central,
                    'Necessidade agregada das farmácias': necessidade_total,
                    'Transferir DE': donor['farmacia'],
                    'Saldo origem': donor['saldo'],
                    'CMD origem': donor['cmd'],
                    'Cobertura origem antes': cobertura_origem_txt,
                    'Excedente disponível': excedente_disponivel,
                    'Transferir PARA': rec['farmacia'],
                    'Saldo destino': saldo_destino_corrente,
                    'CMD destino': rec['cmd'],
                    'Cobertura destino antes': cobertura_destino_txt,
                    'Necessidade para equalizar': necessidade_restante,
                    'Quantidade sugerida remanejar': qtd,
                    'Cobertura alvo hospitalar': f"{cobertura_alvo:.1f} dia(s)",
                    'Cobertura estimada destino após remanejamento': f"{cobertura_destino_apos:.1f} dia(s)",
                    'Observação validade': obs_validade,
                    'Justificativa': justificativa,
                })

                donor['excedente'] -= qtd
                necessidade_restante -= qtd
                saldo_destino_corrente += qtd

    if not resultados:
        return pd.DataFrame()

    out = pd.DataFrame(resultados)
    col_int = [
        'Saldo almoxarifados fornecedores', 'Necessidade agregada das farmácias',
        'Saldo origem', 'Excedente disponível', 'Saldo destino',
        'Necessidade para equalizar', 'Quantidade sugerida remanejar'
    ]
    for c in col_int:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors='coerce').replace([np.inf, -np.inf], 0).fillna(0).astype(int)
    for c in ['CMD origem', 'CMD destino']:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors='coerce').replace([np.inf, -np.inf], 0).fillna(0).round(2)

    out = out[out['Quantidade sugerida remanejar'] > 0].copy()
    if out.empty:
        return pd.DataFrame()

    ordem_cols = [
        'Código MV', 'Material', 'Categoria',
        'Saldo almoxarifados fornecedores', 'Necessidade agregada das farmácias',
        'Transferir DE', 'Saldo origem', 'CMD origem', 'Cobertura origem antes', 'Excedente disponível',
        'Transferir PARA', 'Saldo destino', 'CMD destino', 'Cobertura destino antes',
        'Necessidade para equalizar', 'Quantidade sugerida remanejar',
        'Cobertura alvo hospitalar', 'Cobertura estimada destino após remanejamento',
        'Observação validade', 'Justificativa'
    ]
    for c in ordem_cols:
        if c not in out.columns:
            out[c] = ''

    # Ordenação padrão de abertura: Material em ordem alfabética.
    # Mantém a leitura operacional mais intuitiva e facilita localizar o item.
    return out[ordem_cols].sort_values(
        ['Material', 'Código MV', 'Transferir DE', 'Transferir PARA'],
        ascending=[True, True, True, True],
        kind='mergesort'
    ).reset_index(drop=True)

@st.cache_data(show_spinner=False)
def calcular_status_farmacia(df_estoque: pd.DataFrame, df_mov: pd.DataFrame,
                              cod_farm: str, mapa_cat: dict,
                              data_ini: date, data_fim: date,
                              dias_pedido: int,
                              mapa_antimicrobianos: dict | None = None) -> pd.DataFrame:
    """Roda a lógica de análise completa para UMA farmácia e retorna df_final."""
    # Colunas estoque
    c_ec  = find_col(df_estoque, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod'])
    c_eq  = find_col(df_estoque, ['qtde disp', 'disponivel'])
    c_ep  = find_col(df_estoque, ['material', 'produto', 'descri'])
    c_ea  = find_col(df_estoque, ['almox'])
    c_em  = find_col(df_estoque, ['qtde estq min', 'estoque minimo', 'minimo'])
    if not all([c_ec, c_eq, c_ep, c_ea]):
        return pd.DataFrame()

    est = df_estoque.copy()
    est['key']         = est[c_ec].apply(clean_key)
    est['almox_limpo'] = est[c_ea].apply(clean_key)
    est['saldo_num']   = p_num_series(est[c_eq])
    est['min_num']     = p_num_series(est[c_em]) if c_em else 0.0
    est = est[~est['almox_limpo'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()

    _alvo = (est[est['almox_limpo'] == cod_farm]
             .groupby('key')
             .agg(saldo_num=('saldo_num', 'sum'), min_num=('min_num', 'max')))
    if _alvo.empty:
        return pd.DataFrame()

    est_farm  = _alvo['saldo_num'].to_dict()
    est_min   = _alvo['min_num'].to_dict()
    mapa_prod = est.drop_duplicates('key').set_index('key')[c_ep].to_dict()

    # Centrais
    cent = est[est['almox_limpo'].isin(['1','6','9','41','43'])]
    dict_centrais = {}
    if not cent.empty:
        _c = cent.groupby(['key','almox_limpo'])['saldo_num'].sum().reset_index()
        dict_centrais = (_c.groupby('key')[['almox_limpo','saldo_num']]
                         .apply(lambda g: dict(zip(g['almox_limpo'], g['saldo_num'])),
                                include_groups=False).to_dict())

    # Movimento
    c_mc = find_col(df_mov, ['material', 'cod', 'ca3'])
    c_mq = find_col(df_mov, ['quant'])
    c_mt = find_col(df_mov, ['tipo'])
    c_md = find_col(df_mov, ['data geracao','data mov','data','dt ger','dtger'],
                    forbidden=['almox','tipo','quant','material'])
    c_ma = find_col(df_mov, ['almox'])
    if not all([c_mc, c_mq, c_mt, c_md]):
        return pd.DataFrame()

    mov = filtrar_almoxarifados_excluidos(df_mov, c_ma) if c_ma else df_mov.copy()
    mov['dt_fmt'] = pd.to_datetime(mov[c_md], dayfirst=False, errors='coerce')
    mov_f = mov[(mov['dt_fmt'].dt.date >= data_ini) &
                (mov['dt_fmt'].dt.date <= data_fim) &
                (mov[c_mt].astype(str).str.upper() == 'RM')].copy()

    dias_ef = max(mov_f['dt_fmt'].dt.date.nunique(), 1)
    consumo = (mov_f.assign(qtd_num=lambda d: p_num_series(d[c_mq]),
                            key=lambda d: d[c_mc].apply(clean_key))
               .groupby('key')['qtd_num'].sum().reset_index())
    consumo['cmd'] = consumo['qtd_num'].apply(lambda x: calcular_cmd(x, dias_ef))
    consumo_map = consumo.set_index('key')['cmd'].to_dict()

    todos = sorted(set(est[est['almox_limpo'] == cod_farm]['key']) | set(consumo['key']))
    final = pd.DataFrame({'Código MV': todos})
    final['Material']             = final['Código MV'].map(mapa_prod).fillna('SEM DESCRIÇÃO')
    final['Farmácia']             = DIC_NOMES_FARMACIAS.get(cod_farm, f'Almox {cod_farm}')
    final['Cód. Farmácia']        = cod_farm
    final['Categoria']            = final['Código MV'].map(mapa_cat).fillna('OUTROS')
    mapa_antimicrobianos = mapa_antimicrobianos or {}
    final['Antimicrobianos']      = final['Código MV'].map(mapa_antimicrobianos).fillna('NÃO').apply(normalizar_antimicrobiano)
    final['Saldo Atual']          = final['Código MV'].map(est_farm).fillna(0)
    final['CMD']                  = final['Código MV'].map(consumo_map).fillna(0)
    final['Estoque Mínimo']       = final['Código MV'].map(est_min).fillna(0)
    final['Saldo Central']        = final['Código MV'].apply(
        lambda c: sum(dict_centrais.get(c, {}).values()))
    final['Necessidade']          = calcular_sugestao_vetorizado(
        final.rename(columns={'Saldo Atual': 'Saldo Atual Satélite',
                               'CMD': 'Consumo Médio Diário'}), dias_pedido)
    final['Cobertura (dias)']     = calcular_cobertura_vetorizado(
        final.rename(columns={'Saldo Atual': 'Saldo Atual Satélite',
                               'CMD': 'Consumo Médio Diário'}))

    # Parecer simplificado para consolidação, alinhado à lógica da aba Pedido.
    def _parecer(row):
        cmd = row['CMD'] if pd.notna(row['CMD']) else 0
        saldo = row['Saldo Atual'] if pd.notna(row['Saldo Atual']) else 0
        est_minimo = row['Estoque Mínimo'] if pd.notna(row['Estoque Mínimo']) else 0
        necessidade = row['Necessidade'] if pd.notna(row['Necessidade']) else 0
        saldo_central = row['Saldo Central'] if pd.notna(row['Saldo Central']) else 0

        if cmd == 0 and est_minimo <= 0 and saldo == 0:
            return 'Sem Consumo'
        if cmd > 0 and saldo > (cmd * 60):
            return 'Estoque Excessivo'
        if cmd == 0 and saldo > 0:
            return 'Estoque Parado'
        if necessidade <= 0:
            return 'Estoque Suficiente'
        if saldo_central >= necessidade:
            return 'Solicitar'
        if saldo_central > 0:
            return 'Estoque Crítico CAF'
        return 'Desabastecimento Crítico'

    final['Parecer'] = final.apply(_parecer, axis=1)
    return final

def padronizar_dataframe_categorias(df: pd.DataFrame, origem: str = "") -> tuple[pd.DataFrame | None, str]:
    """Padroniza a base de categorias vinda do Excel local ou Google Sheets.
    Mantém a coluna Antimicrobianos para uso futuro, sem interferir na lógica atual
    de categorias logísticas."""
    try:
        if df is None or df.empty:
            return None, f"A fonte de categorias {origem or 'informada'} está vazia."

        c_cod = find_col(
            df, ['codigo', 'cod', 'ca3'],
            forbidden=['material', 'descri', 'nome', 'produto']
        )
        c_mat = find_col(df, ['material', 'produto', 'insumo', 'descri', 'nome'])
        c_cat = find_col(df, ['categoria', 'grupo', 'classe', 'tipo'])
        c_atb = find_col(df, ['antimicrobiano', 'antimicrobianos', 'antibiotico', 'antibiótico', 'atb'])

        if not c_cod or not c_cat:
            return None, (
                f"A fonte de categorias {origem or 'informada'} foi encontrada, "
                "mas não foi possível identificar as colunas obrigatórias Código e Categoria."
            )

        df_clean = pd.DataFrame()
        df_clean["Código"] = df[c_cod].apply(clean_key)
        df_clean["Material"] = (
            df[c_mat].fillna("").astype(str).str.strip()
            if c_mat else ""
        )
        df_clean["Categoria"] = (
            df[c_cat]
            .fillna("OUTROS")
            .astype(str)
            .str.upper()
            .str.strip()
        )
        if c_atb:
            df_clean["Antimicrobianos"] = df[c_atb].apply(normalizar_antimicrobiano)
        else:
            df_clean["Antimicrobianos"] = "NÃO"

        df_clean = df_clean[df_clean["Código"] != ""].drop_duplicates("Código", keep="last")
        if df_clean.empty:
            return None, f"A fonte de categorias {origem or 'informada'} não possui códigos válidos."

        return df_clean.reset_index(drop=True), ""

    except Exception as e:
        return None, f"Erro ao padronizar categorias de {origem or 'fonte informada'}: {e}"


def carregar_categorias_google_sheets_publico() -> tuple[pd.DataFrame | None, str]:
    """Carrega categorias de uma planilha Google pública por exportação CSV.
    Essa etapa é apenas de leitura; o salvamento direto no Google Sheets exigirá
    autenticação própria em uma etapa futura."""
    try:
        df_raw = pd.read_csv(GOOGLE_SHEETS_CATEGORIAS_URL, dtype=str)
        df_clean, erro = padronizar_dataframe_categorias(df_raw, "Google Sheets público")
        if erro:
            return None, erro
        return df_clean, ""
    except Exception as e:
        return None, f"Erro ao carregar categorias do Google Sheets público: {e}"


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
                df_clean, erro = padronizar_dataframe_categorias(df, f"arquivo local / aba {aba}")
                if df_clean is not None and not df_clean.empty:
                    return df_clean.reset_index(drop=True)
        except Exception as e:
            st.session_state['categorias_erro_local'] = f"Erro ao carregar categorias locais: {e}"

    df_vazio = pd.DataFrame(columns=["Código", "Material", "Categoria", "Antimicrobianos"])
    # Evita tentativa de escrita repetida em ambientes com sistema de arquivos efêmero
    # (ex.: Streamlit Cloud). Só tenta criar o arquivo local se ele ainda não existir;
    # se já existe (ou já falhou antes), não há ganho em tentar gravar de novo a cada fallback.
    if not ARQUIVO_CATEGORIAS.exists():
        try:
            df_vazio.to_excel(ARQUIVO_CATEGORIAS, index=False)
        except Exception:
            pass
    return df_vazio


def carregar_categorias_preferencial() -> pd.DataFrame:
    """Fonte preferencial: Google Sheets público. Fallback: arquivo local."""
    st.session_state['categorias_erro_carga'] = ""
    st.session_state['categorias_erro_local'] = ""

    if USAR_GOOGLE_SHEETS_CATEGORIAS:
        df_gs, erro_gs = carregar_categorias_google_sheets_publico()
        if df_gs is not None and not df_gs.empty:
            st.session_state['categorias_fonte'] = 'Google Sheets público'
            st.session_state['categorias_ultima_carga'] = datetime.now().strftime('%d/%m/%Y %H:%M')
            st.session_state['categorias_erro_carga'] = ""
            return df_gs
        st.session_state['categorias_erro_carga'] = erro_gs or "Google Sheets público não retornou dados válidos."

    df_local = carregar_categorias_do_disco()
    st.session_state['categorias_fonte'] = 'Arquivo local Categorias_base.xlsx'
    st.session_state['categorias_ultima_carga'] = datetime.now().strftime('%d/%m/%Y %H:%M')
    return df_local


def salvar_categorias_no_disco(df: pd.DataFrame) -> bool:
    try:
        df_out = df.copy()
        for col in ["Código", "Material", "Categoria", "Antimicrobianos"]:
            if col not in df_out.columns:
                df_out[col] = ""
        cols_primeiras = ["Código", "Material", "Categoria", "Antimicrobianos"]
        outras_cols = [c for c in df_out.columns if c not in cols_primeiras]
        df_out = df_out[cols_primeiras + outras_cols]
        df_out.to_excel(ARQUIVO_CATEGORIAS, sheet_name="Planilha1", index=False)
        st.session_state['categorias_erro_salvamento'] = ""
        st.session_state['categorias_ultimo_salvamento_local'] = datetime.now().strftime('%d/%m/%Y %H:%M')
        return True
    except Exception as e:
        st.session_state['categorias_erro_salvamento'] = f"Erro ao salvar categorias no arquivo local: {e}"
        st.toast(
            "⚠️ Não foi possível gravar no disco (ambiente Cloud). "
            "Use o botão 'Exportar mapa atual' para salvar suas categorias "
            "e reimporte no próximo acesso.",
            icon="⚠️"
        )
        return False


def inicializar_categorias_session():
    if "df_categorias" not in st.session_state:
        st.session_state["df_categorias"] = carregar_categorias_preferencial()


def obter_mapa_categorias() -> dict:
    df = st.session_state.get("df_categorias", pd.DataFrame())
    if df.empty:
        return {}
    return dict(zip(df["Código"].astype(str), df["Categoria"].astype(str)))


def normalizar_antimicrobiano(v) -> str:
    """Padroniza a indicação de antimicrobiano para SIM/NÃO."""
    if pd.isna(v):
        return "NÃO"
    s = clean(str(v))
    if s in ("sim", "s", "yes", "y", "1", "true", "verdadeiro", "antimicrobiano", "antibiotico", "antibiotico sim"):
        return "SIM"
    if s in ("nao", "não", "n", "no", "0", "false", "falso", ""):
        return "NÃO"
    # Valores inesperados ficam preservados de forma conservadora como NÃO,
    # evitando superestimar rupturas de antimicrobianos.
    return "NÃO"


def eh_antimicrobiano(v) -> bool:
    return normalizar_antimicrobiano(v) == "SIM"


def obter_mapa_antimicrobianos() -> dict:
    df = st.session_state.get("df_categorias", pd.DataFrame())
    if df.empty or "Antimicrobianos" not in df.columns:
        return {}
    tmp = df.copy()
    tmp["Código"] = tmp["Código"].astype(str).apply(clean_key)
    tmp["Antimicrobianos"] = tmp["Antimicrobianos"].apply(normalizar_antimicrobiano)
    tmp = tmp[tmp["Código"] != ""].drop_duplicates("Código", keep="last")
    return dict(zip(tmp["Código"].astype(str), tmp["Antimicrobianos"].astype(str)))


# =============================================================================
# LÓGICA DE NEGÓCIO
# =============================================================================

def calcular_cmd(qtd_total: float, dias_com_movimento: int) -> float:
    cmd_bruto = qtd_total / max(dias_com_movimento, 1)
    return float(math.ceil(cmd_bruto)) if cmd_bruto > 0 else 0.0


def calcular_tendencia(mov_filtrado: pd.DataFrame, c_mov_cod: str,
                        c_mov_qtd: str,
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
        .assign(qtd_num=lambda df: p_num_series(df[c_mov_qtd]))
        .assign(key=lambda df: df[c_mov_cod].apply(clean_key))
        .groupby('key')['qtd_num'].sum()
        .reset_index()
    )

    # 4. Tratamos o resultado matematicamente (vetorizado)
    n_dias_reais = len(ultimos_n_dias)
    consumo_tend['cmd_tendencia'] = np.ceil(consumo_tend['qtd_num'] / n_dias_reais).astype(float)
    
    return consumo_tend[['key', 'cmd_tendencia']]


def calcular_sugestao(row: pd.Series, dias_pedido: int) -> int:
    cmd        = row['Consumo Médio Diário']
    est_atual  = row['Saldo Atual Satélite']
    est_minimo = row['Estoque Mínimo']

    if cmd == 0 and est_minimo > 0:
        return max(0, math.ceil(est_minimo - est_atual))
    if cmd == 0:
        return 0

    meta_final = max(cmd * dias_pedido, est_minimo)
    # ceil em vez de round: em ressuprimento hospitalar é mais seguro
    # arredondar para cima do que arriscar falta por arredondamento
    return max(0, math.ceil(meta_final - est_atual))


def calcular_sugestao_vetorizado(df: pd.DataFrame, dias_pedido: int) -> pd.Series:
    """Versão vetorizada de calcular_sugestao — mesma lógica, sem apply() linha a linha.
    Trata NaN e infinito nas três colunas de entrada para nunca quebrar no cast para int."""
    def _limpar(s):
        s = pd.to_numeric(s, errors='coerce')
        return s.replace([np.inf, -np.inf], 0).fillna(0)

    cmd        = _limpar(df['Consumo Médio Diário'])
    est_atual  = _limpar(df['Saldo Atual Satélite'])
    est_minimo = _limpar(df['Estoque Mínimo'])

    # caso cmd==0 e mínimo>0: sugestão = ceil(mínimo - atual), mínimo 0
    sug_sem_cmd = np.ceil(np.maximum(est_minimo - est_atual, 0)).astype(int)

    # caso cmd>0: meta = max(cmd*dias, mínimo); sugestão = ceil(meta - atual), mínimo 0
    meta = np.maximum(cmd * dias_pedido, est_minimo)
    sug_com_cmd = np.ceil(np.maximum(meta - est_atual, 0)).astype(int)

    return np.where(cmd == 0,
                    np.where(est_minimo > 0, sug_sem_cmd, 0),
                    sug_com_cmd)


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


def calcular_cobertura_vetorizado(df: pd.DataFrame) -> pd.Series:
    """Versão vetorizada de calcular_cobertura — sem apply() linha a linha.
    Trata NaN/inf antes de converter para int, para nunca quebrar com valores ausentes."""
    cmd   = pd.to_numeric(df['Consumo Médio Diário'], errors='coerce').replace([np.inf, -np.inf], 0).fillna(0)
    saldo = pd.to_numeric(df['Saldo Atual Satélite'], errors='coerce').replace([np.inf, -np.inf], 0).fillna(0)

    # Calcula dias só onde cmd > 0; nos demais casos usa 1 como placeholder seguro (evita /0)
    cmd_seguro = cmd.where(cmd > 0, 1)
    dias_float = saldo / cmd_seguro
    # Proteção extra: qualquer NaN/inf remanescente vira 0 antes do cast para int
    dias_float = dias_float.replace([np.inf, -np.inf], 0).fillna(0)
    dias_int   = dias_float.astype(int).astype(str)

    return np.where(cmd <= 0,
                    np.where(saldo <= 0, "Sem dado", "+∞"),
                    dias_int)


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
    """Define o parecer logístico e a ação sugerida para um item.

    Mantida como apply() linha a linha por necessidade de texto: a função monta
    frases com nomes de farmácias e quantidades específicas por item, o que não
    se vetoriza com segurança sem reescrever toda a lógica de mensagens.
    Os acessos a dicionário usam .get() com dict vazio padrão para evitar
    recriações repetidas e reduzir custo por chamada.
    """
    cod        = row['Código MV']
    # Blindagem contra NaN/None vindos de colunas com dados ausentes —
    # evita "Cannot convert non-finite values to integer" mais adiante
    sug        = row['Necessidade de Ressuprimento'] if pd.notna(row['Necessidade de Ressuprimento']) else 0
    cmd        = row['Consumo Médio Diário'] if pd.notna(row['Consumo Médio Diário']) else 0
    est_un     = row['Saldo Atual Satélite'] if pd.notna(row['Saldo Atual Satélite']) else 0
    est_minimo = row['Estoque Mínimo'] if pd.notna(row['Estoque Mínimo']) else 0

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

    saldos_parceiras   = dict_saldos_parceiras.get(cod) or {}
    consumos_parceiras = consumo_outras_total.get(cod) or {}
    farmacias_paradas  = [
        f"Cód {fid} ({DIC_NOMES_FARMACIAS.get(str(fid), 'Farmácia Satélite')} - {int(sf)} un.)"
        for fid, sf in saldos_parceiras.items()
        if sf > 0 and (
            consumos_parceiras.get(fid, 0) == 0 or
            sf > consumos_parceiras.get(fid, 0) * 3
        )
    ]
    locais_remanejo = " | ".join(farmacias_paradas)

    saldos_nas_centrais = dict_saldos_centrais.get(cod) or {}
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
# RENDERIZAÇÃO — PEDIDO DA FARMÁCIA ATIVA
# =============================================================================
def render_painel_pedido_completo() -> None:
    """Exibe a análise completa do pedido da farmácia ativa.

    A Central de Processamento deve ficar apenas com upload/processamento.
    Toda a leitura operacional do resultado fica concentrada nesta função,
    chamada na aba 📦 Pedido da Farmácia Ativa.
    """
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

        # ── BADGE RESUMO DE VALIDADES (itens deste pedido com validade crítica) ──
        if '⏰ Validade' in df_view.columns:
            n_venc_pedido = df_view['⏰ Validade'].str.contains('VENCIDO', na=False).sum()
            n_crit_pedido = df_view['⏰ Validade'].str.contains('Crítico', na=False).sum()
            if n_venc_pedido > 0 or n_crit_pedido > 0:
                partes = []
                if n_venc_pedido > 0:
                    partes.append(f"**{n_venc_pedido}** vencido(s)")
                if n_crit_pedido > 0:
                    partes.append(f"**{n_crit_pedido}** crítico(s) (≤30 dias)")
                st.error(
                    f"⏰ **Atenção às validades:** {' e '.join(partes)} entre os itens desta análise. "
                    f"Veja a coluna **⏰ Validade** na tabela abaixo ou confira o detalhe completo na "
                    f"aba **Controle de Validade**."
                )
            elif 'df_validades_mescladas' in st.session_state:
                st.success("⏰ **Validades:** nenhum item desta farmácia está vencido ou em estado crítico.")

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
            'Estoque Mínimo', 'Saldo Atual Satélite', '⏰ Validade',
            'Cobertura (dias)', 'CMD Últ. 3 dias', 'Consumo Médio Diário', COL_SUG,
            'Tendência', 'Δ% Tendência',
            'Saldo Almox. Centrais Unificado',
            'Parecer Logístico / Alerta', 'Ação Logística Sugerida',
        ]
        larguras_rel = {
            'Código MV': 12, 'Material': 45, 'Categoria': 16,
            'Estoque Mínimo': 16, 'Saldo Atual Satélite': 18,
            'Cobertura (dias)': 14, 'CMD Últ. 3 dias': 16,
            'Consumo Médio Diário': 18, COL_SUG: 26,
            f'PEDIDO ({dias_pedido} DIAS)': 20,
            'Tendência': 10, 'Δ% Tendência': 12, '⏰ Validade': 18,
            'Saldo Almox. Centrais Unificado': 28,
            'Parecer Logístico / Alerta': 26, 'Ação Logística Sugerida': 50,
        }

        def _preparar_df_card(df_raw: pd.DataFrame, incluir_antimicrobianos: bool = False) -> pd.DataFrame:
            df = df_raw.rename(columns={'Necessidade de Ressuprimento': COL_SUG})
            cols_base = list(ordem_cols)
            if incluir_antimicrobianos and 'Antimicrobianos' in df.columns:
                # No relatório de desabastecimento crítico, sinaliza se o item é antimicrobiano.
                if 'Antimicrobianos' not in cols_base:
                    pos = cols_base.index('Categoria') + 1 if 'Categoria' in cols_base else 0
                    cols_base.insert(pos, 'Antimicrobianos')
            cols_ok = [c for c in cols_base if c in df.columns]
            return df[cols_ok]

        ACAO_INATIVAR_ITEM = "Avaliar se é necessário inativar o item na farmácia."
        df_view_visual = df_view[
            (df_view['Parecer Logístico / Alerta'] != 'Sem Consumo') &
            (df_view['Ação Logística Sugerida'] != ACAO_INATIVAR_ITEM)
        ].copy()

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
            gerar_download_excel_sob_demanda(
                "📄 Gerar extração",
                _preparar_df_card(df_desabast, incluir_antimicrobianos=True),
                "Rupturas",
                f"Rupturas_{cod_farmacia_alvo}.xlsx",
                key=f"ex_c1_{cod_farmacia_alvo}",
                use_container_width=True,
            )
        with c2:
            st.metric("🔄 Remanejamento Potencial", f"{len(df_remanej)} itens")
            gerar_download_excel_sob_demanda(
                "📄 Gerar extração",
                _preparar_df_card(df_remanej),
                "Remanejamento",
                f"Remanejamento_{cod_farmacia_alvo}.xlsx",
                key=f"ex_c2_{cod_farmacia_alvo}",
                use_container_width=True,
            )
        with c3:
            st.metric("📦 Disponível no Almoxarifado", f"{len(df_caf)} itens")
            gerar_download_excel_sob_demanda(
                "📄 Gerar extração",
                _preparar_df_card(df_caf),
                "Disponiveis_CAF",
                f"Disponiveis_{cod_farmacia_alvo}.xlsx",
                key=f"ex_c3_{cod_farmacia_alvo}",
                use_container_width=True,
            )
        with c4:
            st.metric("⚠️ Excesso / Sem Giro", f"{len(df_excesso)} itens")
            gerar_download_excel_sob_demanda(
                "📄 Gerar extração",
                _preparar_df_card(df_excesso),
                "Overstock",
                f"Overstock_{cod_farmacia_alvo}.xlsx",
                key=f"ex_c4_{cod_farmacia_alvo}",
                use_container_width=True,
            )

        # ── GRÁFICOS ──────────────────────────────────────
        st.write("")
        g1, g2 = st.columns(2)

        with g1:
            st.markdown("**Saúde Geral do Estoque**")
            df_g1 = df_view_visual.copy()
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
                status_g1 = [s for s in MAPA_CORES_GRAFICO.keys() if s in set(df_g1_grp['Status'])]
                st.altair_chart(
                    alt.Chart(df_g1_grp).mark_arc(innerRadius=65, stroke='#fff').encode(
                        theta=alt.Theta('Quantidade:Q'),
                        color=alt.Color('Status:N', scale=alt.Scale(
                            domain=status_g1,
                            range=[MAPA_CORES_GRAFICO[s] for s in status_g1]
                        ), legend=None),
                        tooltip=['Status:N', 'Quantidade:Q']
                    ).properties(height=350),
                    use_container_width=True
                )

        with g2:
            st.markdown("**Matriz de Urgência por Categoria**")
            df_g2 = df_view_visual[df_view_visual['Categoria'] != 'OUTROS'].copy()
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

                status_g2 = [s for s in MAPA_CORES_GRAFICO.keys() if s in set(df_g2_grp['Parecer'])]
                heatmap = base.mark_rect(cornerRadius=6, stroke='white', strokeWidth=3).encode(
                    color=alt.Color('Parecer:N', scale=alt.Scale(
                        domain=status_g2,
                        range=[MAPA_CORES_GRAFICO[s] for s in status_g2]
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
        status_legenda = []
        if 'df_g1_grp' in locals() and not df_g1_grp.empty:
            status_legenda.extend(df_g1_grp['Status'].dropna().astype(str).tolist())
        if 'df_g2_grp' in locals() and not df_g2_grp.empty:
            status_legenda.extend(df_g2_grp['Parecer'].dropna().astype(str).tolist())
        status_legenda = [s for s in MAPA_CORES_GRAFICO.keys() if s in set(status_legenda)]
        if status_legenda:
            legend_html = "<div style='display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; padding: 10px; background-color: #F8FAFC; border-radius: 8px; border: 1px solid #E2E8F0;'>"
            for status in status_legenda:
                color = MAPA_CORES_GRAFICO[status]
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
            opcoes_alertas = sorted(df_view_visual['Parecer Logístico / Alerta'].unique().tolist())
            busca_alerta   = f2.multiselect("Filtrar por Parecer:", options=opcoes_alertas, key="busca_alerta")
            busca_cat      = f3.selectbox("Categoria:", ["TODAS"] + sorted(df_view_visual['Categoria'].unique().tolist()), key="busca_cat")

        df_filtrado = df_view_visual.copy()
        if busca_nome:
            df_filtrado = df_filtrado[
                df_filtrado['Material'].astype(str).str.contains(busca_nome, case=False, na=False, regex=False) |
                df_filtrado['Código MV'].astype(str).str.contains(busca_nome, case=False, na=False, regex=False)
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
                "⏰ Validade":              st.column_config.TextColumn("Validade", width="medium"),
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
            gerar_download_excel_sob_demanda(
                "📄 Gerar relatório completo — aba única (.xlsx)",
                df_export_geral,
                "Painel Geral",
                f"Painel_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                key=f"pedido_completo_{cod_farmacia_alvo}",
                use_container_width=True,
            )
        with b2:
            # Relatório operacional por categoria: sem coluna de parecer visível,
            # mas mantendo as cores por meio de uma coluna auxiliar oculta.
            col_pedido_label = f'PEDIDO ({dias_pedido} DIAS)'
            df_pedido_abas = df_view.copy()

            # Observação de validade: mostra somente itens da farmácia ativa que
            # ainda possuem saldo AGHU e vencem nos próximos 90 dias.
            df_pedido_abas['Observação'] = ""
            if 'df_validades_mescladas' in st.session_state:
                df_val_pedido = st.session_state['df_validades_mescladas'].copy()
                cod_farm_atual = st.session_state.get('cod_farmacia_alvo', '')

                if not df_val_pedido.empty and {'key', 'Farmácia', 'Dias até Vencer', 'Validade Fmt'}.issubset(df_val_pedido.columns):
                    df_val_pedido['Dias até Vencer'] = pd.to_numeric(
                        df_val_pedido['Dias até Vencer'], errors='coerce'
                    )
                    if 'Saldo AGHU' in df_val_pedido.columns:
                        df_val_pedido['Saldo AGHU'] = pd.to_numeric(
                            df_val_pedido['Saldo AGHU'], errors='coerce'
                        ).fillna(0)
                        mask_saldo = df_val_pedido['Saldo AGHU'] > 0
                    else:
                        mask_saldo = True

                    df_val_pedido = df_val_pedido[
                        (df_val_pedido['Farmácia'].astype(str) == str(cod_farm_atual)) &
                        (df_val_pedido['Dias até Vencer'].between(0, 90, inclusive='both')) &
                        mask_saldo
                    ].copy()

                    if not df_val_pedido.empty:
                        df_val_pedido = df_val_pedido.sort_values(['key', 'Dias até Vencer'])
                        df_val_pedido['Observação'] = df_val_pedido.apply(
                            lambda r: (
                                f"Confirmar se ainda tem a vencer em {int(r['Dias até Vencer'])} dia(s) "
                                f". Evitar solicitar grande quantidade"
                            ),
                            axis=1
                        )
                        mapa_obs_validade = (
                            df_val_pedido.drop_duplicates('key')
                            .set_index('key')['Observação']
                            .to_dict()
                        )
                        df_pedido_abas['Observação'] = (
                            df_pedido_abas['Código MV'].map(mapa_obs_validade).fillna("")
                        )

            # Coluna auxiliar para manter a coloração por status sem exibir o parecer no relatório.
            df_pedido_abas['Status para cor'] = df_pedido_abas['Parecer Logístico / Alerta']

            df_pedido_abas = df_pedido_abas.rename(columns={
                'Necessidade de Ressuprimento': col_pedido_label,
                'Saldo Atual Satélite': 'Saldo atual da Farmácia',
            })

            ordem_abas_final = [
                'Código MV', 'Material', 'Categoria', 'Estoque Mínimo',
                'Saldo atual da Farmácia', 'Cobertura (dias)', 'Δ% Tendência',
                col_pedido_label, 'Observação', 'Status para cor'
            ]

            larguras_pedido_abas = {
                **larguras_rel,
                'Saldo atual da Farmácia': 20,
                col_pedido_label: 20,
                'Observação': 34,
                'Status para cor': 26,
            }

            gerar_download_multi_aba_sob_demanda(
                "📄 Gerar pedido classificado por categoria (.xlsx)",
                df_pedido_abas,
                ordem_abas_final,
                col_categoria='Categoria',
                col_alerta='Status para cor',
                larguras=larguras_pedido_abas,
                excluir_acoes=["Avaliar se é necessário inativar o item na farmácia."],
                ocultar_colunas=['Status para cor'],
                ajustar_altura_linhas=False,
                file_name=f"Pedido_Abas_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                key=f"pedido_abas_{cod_farmacia_alvo}",
                use_container_width=True,
                orientacao_impressao="portrait",
            )
        with b3:
            df_filtrado_export = df_filtrado[cols_exibicao]
            gerar_download_excel_sob_demanda(
                "📄 Gerar resultado filtrado (.xlsx)",
                df_filtrado_export,
                "Filtro Atual",
                f"Filtro_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                key=f"pedido_filtro_{cod_farmacia_alvo}",
                use_container_width=True,
            )
    else:
        st.warning(
            "⚠️ Nenhum pedido processado nesta sessão. Carregue os arquivos obrigatórios "
            "na aba **📥 Central de Processamento** e clique em **Analisar os dados**."
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

    # Persistir parâmetros para uso nas abas Consolidação e Validade
    st.session_state['data_inicio_huufma']  = data_inicio
    st.session_state['data_fim_huufma']     = data_fim
    st.session_state['dias_pedido_huufma']  = dias_pedido

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

        ---
        ### Remanejamentos entre farmácias (aba 🔄 Remanejamentos)

        #### Remanejamento geral por contingência/equalização
        * **Quando aciona:** apenas quando, para o mesmo item, há simultaneamente:
          necessidade real por consumo em alguma farmácia, saldo insuficiente nos
          almoxarifados centrais para cobrir essa necessidade, e saldo disponível em
          outra farmácia satélite.
        * **Como distribui:** o saldo das farmácias é redistribuído de forma
          proporcional ao CMD de cada uma, buscando aproximar a cobertura (em dias)
          entre as farmácias que consomem o item. Farmácias sem consumo recente têm
          estoque-alvo zero e podem doar todo o saldo disponível.
        * **O que NÃO faz:** não equaliza estoques indiscriminadamente; só atua em
          cenário real de contingência, conforme os gatilhos acima.

        #### Remanejamento preventivo por validade (FEFO)
        * **Quando aciona:** identifica itens com validade próxima (não vencidos,
          dentro da janela de até 90 dias) em uma farmácia de origem com saldo AGHU
          confirmado, e localiza outra farmácia que consome o mesmo item e não possui
          alerta de validade para ele.
        * **Como calcula a quantidade:** limitada pelo menor valor entre o saldo em
          risco de vencer na origem e a capacidade estimada de consumo do destino até
          a data de validade, com base no CMD do destino.
        * **Itens vencidos:** não entram nesta análise como oportunidade de uso —
          a conduta esperada é segregação/retirada conforme rotina institucional.
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

tab_processamento, tab_pedido, tab_consolidacao, tab_remanejamentos, tab_validade, tab_categorias = st.tabs([
    "📥 Central de Processamento",
    "📦 Pedido da Farmácia Ativa",
    "🏥 Consolidação Multi-Farmácia",
    "🔄 Remanejamentos",
    "⏰ Controle de Validade",
    "🗂️ Categorias",
])

# Aliases mantêm a estrutura do código existente e permitem reorganizar o layout
# sem alterar a lógica de negócio já validada.
tab1 = tab_processamento
tab2 = tab_validade
tab3 = tab_consolidacao
tab4 = tab_categorias


# =============================================================================
# TAB — PEDIDO DA FARMÁCIA ATIVA
# =============================================================================
with tab_pedido:
    st.subheader("📦 Pedido da Farmácia Ativa")
    st.caption(
        "Análise operacional completa gerada a partir dos arquivos carregados na "
        "aba **📥 Central de Processamento**."
    )
    render_painel_pedido_completo()


# =============================================================================
# TAB — REMANEJAMENTOS
# =============================================================================
with tab_remanejamentos:
    st.subheader("🔄 Remanejamentos")
    st.caption(
        "Plano operacional para redistribuição entre farmácias. A análise usa a consolidação "
        "multi-farmácia já processada e mantém as exportações completas e filtradas."
    )

    st.info(
        "⏳ Os remanejamentos são calculados a partir da consolidação multi-farmácia. "
        "Se houver muitos itens, a avaliação pode levar alguns instantes. "
        "Use o painel de acompanhamento abaixo para visualizar a etapa em execução antes de consultar filtros ou exportar relatórios."
    )

    if 'df_consolidado' not in st.session_state:
        st.warning(
            "⚠️ Gere a consolidação multi-farmácia primeiro na aba **Consolidação Multi-Farmácia**. "
            "Os remanejamentos dependem dos saldos e CMDs consolidados por farmácia."
        )
    else:
        df_cons = st.session_state['df_consolidado'].copy()
        total_itens_rem = int(df_cons['Código MV'].nunique()) if 'Código MV' in df_cons.columns else len(df_cons)
        total_farm_rem = int(df_cons['Farmácia'].nunique()) if 'Farmácia' in df_cons.columns else 0
        tem_validade_rem = (
            'df_validades_mescladas' in st.session_state and
            isinstance(st.session_state.get('df_validades_mescladas'), pd.DataFrame) and
            not st.session_state.get('df_validades_mescladas').empty
        )
        cache_geral_pronto = 'df_remanejamento_geral_huufma' in st.session_state
        cache_validade_pronto = (not tem_validade_rem) or ('df_remanejamento_validade_huufma' in st.session_state)
        remanejamentos_prontos = cache_geral_pronto and cache_validade_pronto

        with st.container(border=True):
            st.markdown("##### 📊 Acompanhamento do processamento dos remanejamentos")
            pr1, pr2, pr3, pr4 = st.columns(4)
            pr1.metric("Itens avaliáveis", f"{total_itens_rem}")
            pr2.metric("Farmácias", f"{total_farm_rem}")
            pr3.metric(
                "Remanejamento geral",
                "✅ Pronto" if cache_geral_pronto else "⏳ Pendente",
                f"{len(st.session_state.get('df_remanejamento_geral_huufma', pd.DataFrame()))} sugestão(ões)" if cache_geral_pronto else None,
            )
            pr4.metric(
                "Validade/FEFO",
                "✅ Pronto" if cache_validade_pronto else "⏳ Pendente",
                f"{len(st.session_state.get('df_remanejamento_validade_huufma', pd.DataFrame()))} sugestão(ões)" if 'df_remanejamento_validade_huufma' in st.session_state else ("Sem validade carregada" if not tem_validade_rem else None),
            )

            ultima_geral = st.session_state.get('remanejamento_geral_ultima_atualizacao')
            ultima_val = st.session_state.get('remanejamento_validade_ultima_atualizacao')
            if ultima_geral or ultima_val:
                st.caption(
                    "Último cálculo — "
                    f"Geral: **{ultima_geral or 'não calculado'}** | "
                    f"Validade/FEFO: **{ultima_val or ('não aplicável' if not tem_validade_rem else 'não calculado')}**."
                )
            else:
                st.caption(
                    "A análise ainda não foi preparada nesta sessão. Clique no botão abaixo para calcular uma vez "
                    "e reaproveitar o resultado nos filtros e relatórios."
                )

            preparar_remanejamentos = st.button(
                "🚀 Gerar/atualizar análise de remanejamentos agora",
                key="btn_preparar_remanejamentos_com_progresso",
                type="primary",
                use_container_width=True,
                help="Calcula remanejamento geral e, quando houver validade carregada, o remanejamento preventivo por FEFO."
            )

        if preparar_remanejamentos:
            barra_rem = st.progress(0, text="Iniciando análise de remanejamentos...")
            status_rem = st.empty()

            status_rem.info("1/5 Conferindo dados consolidados por farmácia...")
            barra_rem.progress(10, text="1/5 Conferindo consolidação multi-farmácia...")

            df_val_rem_geral = st.session_state.get('df_validades_mescladas', pd.DataFrame())
            status_rem.info("2/5 Preparando dados de validade e saldos AGHU, quando disponíveis...")
            barra_rem.progress(25, text="2/5 Preparando dados auxiliares...")
            if isinstance(df_val_rem_geral, pd.DataFrame) and not df_val_rem_geral.empty and 'est_geral_raw' in st.session_state:
                df_val_rem_geral = aplicar_saldos_validades(df_val_rem_geral.copy(), st.session_state['est_geral_raw'])

            status_rem.info("3/5 Calculando oportunidades de remanejamento geral por contingência/equalização...")
            barra_rem.progress(45, text="3/5 Calculando remanejamento geral...")
            obter_remanejamento_geral_session_cache(
                df_cons,
                df_val_rem_geral,
                dias_cobertura=st.session_state.get('dias_pedido_huufma', 15),
                forcar_recalculo=True,
            )

            if tem_validade_rem:
                status_rem.info("4/5 Calculando remanejamento preventivo por validade (FEFO)...")
                barra_rem.progress(75, text="4/5 Calculando remanejamento por validade...")
                df_val_rem = st.session_state['df_validades_mescladas'].copy()
                if 'est_geral_raw' in st.session_state:
                    df_val_rem = aplicar_saldos_validades(df_val_rem, st.session_state['est_geral_raw'])
                obter_remanejamento_validade_session_cache(
                    df_val_rem,
                    df_cons,
                    forcar_recalculo=True,
                )
            else:
                st.session_state['df_remanejamento_validade_huufma'] = pd.DataFrame()
                st.session_state['remanejamento_validade_ultima_atualizacao'] = 'não aplicável — validade não carregada'
                barra_rem.progress(75, text="4/5 Remanejamento por validade não aplicável...")

            status_rem.info("5/5 Finalizando e atualizando a visualização...")
            barra_rem.progress(100, text="5/5 Remanejamentos prontos.")
            status_rem.success("✅ Análise de remanejamentos concluída. A tela será atualizada com os resultados.")
            st.rerun()

        if not remanejamentos_prontos:
            st.info(
                "Clique em **Gerar/atualizar análise de remanejamentos agora** para calcular os resultados. "
                "Depois disso, os filtros e relatórios usam o resultado salvo em sessão, sem recalcular a cada interação."
            )
        else:
            # ── OPORTUNIDADES DE REMANEJAMENTO ────────────────────────────────
            st.write("---")
            st.markdown("#### 🔄 Oportunidades de Remanejamento Entre Farmácias")
            st.caption(
                "Aciona remanejamento geral apenas em contingência: há farmácia precisando, "
                "os almoxarifados fornecedores estão zerados ou insuficientes, e existe saldo em outra farmácia. "
                "A redistribuição considera o CMD para aproximar a cobertura entre as unidades consumidoras."
            )

            df_opor = st.session_state.get('df_remanejamento_geral_huufma', pd.DataFrame()).copy()
            ultima_rem = st.session_state.get('remanejamento_geral_ultima_atualizacao')
            if ultima_rem:
                st.caption(f"Resultado de remanejamento geral calculado em {ultima_rem} e reaproveitado enquanto dados/parâmetros não mudarem.")

            if df_opor.empty:
                st.info("Nenhuma oportunidade de remanejamento geral foi identificada: não houve simultaneamente necessidade real por consumo, almoxarifado insuficiente e saldo disponível em outra farmácia.")
            else:
                st.success(
                    f"🔄 {len(df_opor)} sugestão(ões) de remanejamento geral por contingência/equalização identificada(s)."
                )

                # Filtros operacionais: permitem gerar solicitação por origem ou por destino específico.
                fr1, fr2 = st.columns(2)
                op_origens = ["TODAS"] + sorted(df_opor['Transferir DE'].dropna().astype(str).unique().tolist())
                op_destinos = ["TODAS"] + sorted(df_opor['Transferir PARA'].dropna().astype(str).unique().tolist())
                filtro_origem_rem = fr1.selectbox(
                    "Filtrar farmácia de origem:",
                    op_origens,
                    key="filtro_remanejamento_geral_origem",
                )
                filtro_destino_rem = fr2.selectbox(
                    "Filtrar farmácia de destino:",
                    op_destinos,
                    key="filtro_remanejamento_geral_destino",
                )

                df_opor_view = df_opor.copy()
                if filtro_origem_rem != "TODAS":
                    df_opor_view = df_opor_view[df_opor_view['Transferir DE'] == filtro_origem_rem]
                if filtro_destino_rem != "TODAS":
                    df_opor_view = df_opor_view[df_opor_view['Transferir PARA'] == filtro_destino_rem]

                st.caption(
                    f"Exibindo {len(df_opor_view)} de {len(df_opor)} sugestão(ões). "
                    "A exportação abaixo respeita os filtros aplicados."
                )

                if df_opor_view.empty:
                    st.info("Nenhuma sugestão permaneceu após os filtros selecionados.")
                else:
                    # A tela fica propositalmente mais enxuta; o Excel exportado continua
                    # levando todas as colunas técnicas calculadas para auditoria e conferência.
                    cols_tela_remanejamento = [
                        'Código MV', 'Material', 'Saldo almoxarifados fornecedores',
                        'Transferir DE', 'Transferir PARA', 'Quantidade sugerida remanejar',
                        'Cobertura alvo hospitalar',
                        'Cobertura estimada destino após remanejamento',
                        'Justificativa'
                    ]
                    cols_tela_remanejamento = [c for c in cols_tela_remanejamento if c in df_opor_view.columns]

                    df_opor_view_tela = (
                        df_opor_view
                        .sort_values('Material', kind='mergesort')
                        .reset_index(drop=True)
                    )
                    df_opor_view_export = df_opor_view_tela.copy()
                    df_opor_geral_export = (
                        df_opor
                        .sort_values('Material', kind='mergesort')
                        .reset_index(drop=True)
                    )

                    st.dataframe(
                        df_opor_view_tela[cols_tela_remanejamento],
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            'Código MV': st.column_config.TextColumn('Código MV', width='small'),
                            'Material': st.column_config.TextColumn('Material', width='large'),
                            'Saldo almoxarifados fornecedores': st.column_config.NumberColumn('Saldo almoxarifados', format='%d'),
                            'Transferir DE': st.column_config.TextColumn('Transferir DE', width='medium'),
                            'Transferir PARA': st.column_config.TextColumn('Transferir PARA', width='medium'),
                            'Quantidade sugerida remanejar': st.column_config.NumberColumn('Qtd sugerida', format='%d'),
                            'Cobertura alvo hospitalar': st.column_config.TextColumn('Cobertura alvo hospitalar', width='medium'),
                            'Cobertura estimada destino após remanejamento': st.column_config.TextColumn('Cobertura destino após', width='medium'),
                            'Justificativa': st.column_config.TextColumn('Justificativa', width='large'),
                        }
                    )

                    ex_rem1, ex_rem2 = st.columns(2)
                    with ex_rem1:
                        gerar_download_excel_sob_demanda(
                            "📄 Gerar remanejamento geral COMPLETO (.xlsx)",
                            df_opor_geral_export,
                            "Remanejamento_Geral",
                            f"Remanejamento_Geral_Completo_{datetime.now().strftime('%d%m%y')}.xlsx",
                            key="remanejamento_geral_completo",
                            use_container_width=True,
                        )
                    with ex_rem2:
                        gerar_download_excel_sob_demanda(
                            "📄 Gerar remanejamento geral FILTRADO (.xlsx)",
                            df_opor_view_export,
                            "Remanejamento_Filtrado",
                            f"Remanejamento_Geral_Filtrado_{datetime.now().strftime('%d%m%y')}.xlsx",
                            key="remanejamento_geral_filtrado",
                            use_container_width=True,
                        )


            # ── REMANEJAMENTO PREVENTIVO POR VALIDADE ─────────────────────────
            st.write("")
            st.markdown("##### ⏰ Remanejamento preventivo por validade (FEFO)")
            st.caption(
                "Identifica itens com validade próxima em uma farmácia e sugere envio para outra "
                "farmácia que consome o item e não possui alerta de validade para o mesmo código. "
                "Itens vencidos não entram como oportunidade de remanejamento para uso; devem ser segregados."
            )

            if 'df_validades_mescladas' not in st.session_state:
                st.info(
                    "Carregue o Controle de Validade na aba **📥 Central de Processamento** para ativar a análise de "
                    "remanejamento preventivo por vencimento."
                )
            else:
                df_val_rem = st.session_state['df_validades_mescladas'].copy()
                # Reaplica o saldo do AGHU, quando disponível, para garantir que o alerta seja operacional.
                if 'est_geral_raw' in st.session_state:
                    df_val_rem = aplicar_saldos_validades(df_val_rem, st.session_state['est_geral_raw'])

                df_rem_val = st.session_state.get('df_remanejamento_validade_huufma', pd.DataFrame()).copy()
                ultima_rem_val = st.session_state.get('remanejamento_validade_ultima_atualizacao')
                if ultima_rem_val:
                    st.caption(f"Resultado de remanejamento por validade calculado em {ultima_rem_val} e reaproveitado enquanto dados não mudarem.")

                n_vencidos_sem_remanejo = 0
                if not df_val_rem.empty and 'Dias até Vencer' in df_val_rem.columns:
                    n_vencidos_sem_remanejo = pd.to_numeric(
                        df_val_rem['Dias até Vencer'], errors='coerce'
                    ).lt(0).sum()

                if n_vencidos_sem_remanejo:
                    st.warning(
                        f"💀 {n_vencidos_sem_remanejo} registro(s) vencido(s) não foram considerados "
                        "para remanejamento preventivo. A conduta esperada é segregação/retirada conforme rotina institucional."
                    )

                if df_rem_val.empty:
                    st.success(
                        "✅ Nenhuma oportunidade segura de remanejamento preventivo por validade foi identificada."
                    )
                else:
                    st.success(
                        f"⏰ {len(df_rem_val)} oportunidade(s) de remanejamento preventivo por validade identificada(s)."
                    )

                    # Filtros operacionais: permitem gerar solicitação por origem ou por destino específico,
                    # mantendo a exportação completa para auditoria/conferência.
                    fv_rem1, fv_rem2 = st.columns(2)
                    op_origens_val = ["TODAS"] + sorted(
                        df_rem_val['Transferir DE'].dropna().astype(str).unique().tolist()
                    )
                    op_destinos_val = ["TODAS"] + sorted(
                        df_rem_val['Transferir PARA'].dropna().astype(str).unique().tolist()
                    )
                    filtro_origem_val = fv_rem1.selectbox(
                        "Filtrar farmácia de origem:",
                        op_origens_val,
                        key="filtro_remanejamento_validade_origem",
                    )
                    filtro_destino_val = fv_rem2.selectbox(
                        "Filtrar farmácia de destino:",
                        op_destinos_val,
                        key="filtro_remanejamento_validade_destino",
                    )

                    df_rem_val_view = df_rem_val.copy()
                    if filtro_origem_val != "TODAS":
                        df_rem_val_view = df_rem_val_view[df_rem_val_view['Transferir DE'] == filtro_origem_val]
                    if filtro_destino_val != "TODAS":
                        df_rem_val_view = df_rem_val_view[df_rem_val_view['Transferir PARA'] == filtro_destino_val]

                    st.caption(
                        f"Exibindo {len(df_rem_val_view)} de {len(df_rem_val)} sugestão(ões). "
                        "A exportação filtrada respeita os filtros aplicados; a exportação geral mantém todas as sugestões."
                    )

                    if df_rem_val_view.empty:
                        st.info("Nenhuma sugestão permaneceu após os filtros selecionados.")
                    else:
                        # A tela fica propositalmente mais enxuta; os Excel gerados continuam
                        # levando todas as colunas técnicas calculadas para auditoria e conferência.
                        cols_tela_validade = [
                            'Prioridade', 'Código MV', 'Material',
                            'Transferir DE', 'Transferir PARA', 'Dias até Vencer',
                            'CMD Destino', 'Consumo Possível no Destino até Validade',
                            'Qtd Sugerida Remanejar', 'Justificativa'
                        ]
                        cols_tela_validade = [c for c in cols_tela_validade if c in df_rem_val_view.columns]

                        df_rem_val_view_tela = (
                            df_rem_val_view
                            .sort_values(['Material', 'Dias até Vencer'], kind='mergesort')
                            .reset_index(drop=True)
                        )
                        df_rem_val_export_filtrado = df_rem_val_view_tela.copy()
                        df_rem_val_export_geral = (
                            df_rem_val
                            .sort_values(['Material', 'Dias até Vencer'], kind='mergesort')
                            .reset_index(drop=True)
                        )

                        st.dataframe(
                            df_rem_val_view_tela[cols_tela_validade],
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                'Prioridade': st.column_config.TextColumn('Prioridade', width='medium'),
                                'Código MV': st.column_config.TextColumn('Código', width='small'),
                                'Material': st.column_config.TextColumn('Material', width='large'),
                                'Transferir DE': st.column_config.TextColumn('Transferir DE', width='medium'),
                                'Transferir PARA': st.column_config.TextColumn('Transferir PARA', width='medium'),
                                'Dias até Vencer': st.column_config.NumberColumn('Dias para vencer', format='%d'),
                                'CMD Destino': st.column_config.NumberColumn('CMD destino', format='%d'),
                                'Consumo Possível no Destino até Validade': st.column_config.NumberColumn('Consumo possível até validade', format='%d'),
                                'Qtd Sugerida Remanejar': st.column_config.NumberColumn('Quantidade sugerida', format='%d'),
                                'Justificativa': st.column_config.TextColumn('Justificativa', width='large'),
                            }
                        )

                        ex_val1, ex_val2 = st.columns(2)
                        with ex_val1:
                            gerar_download_excel_sob_demanda(
                                "📄 Gerar remanejamento por validade COMPLETO (.xlsx)",
                                df_rem_val_export_geral,
                                "Remanejamento_Validade",
                                f"Remanejamento_Validade_Completo_{datetime.now().strftime('%d%m%y')}.xlsx",
                                key="remanejamento_validade_completo",
                                use_container_width=True,
                            )
                        with ex_val2:
                            gerar_download_excel_sob_demanda(
                                "📄 Gerar remanejamento por validade FILTRADO (.xlsx)",
                                df_rem_val_export_filtrado,
                                "Remanejamento_Val_Filtro",
                                f"Remanejamento_Validade_Filtrado_{datetime.now().strftime('%d%m%y')}.xlsx",
                                key="remanejamento_validade_filtrado",
                                use_container_width=True,
                            )



# =============================================================================
# TAB 2 — CONTROLE DE VALIDADE
# =============================================================================
with tab2:
    st.subheader("⏰ Controle de Validade — FEFO")

    # ── FONTE DE DADOS ─────────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("##### 📥 Fonte de dados de validade")
        ultima_carga_val = st.session_state.get('validade_ultima_carga')
        origem_val = st.session_state.get('validade_ultima_origem', 'não informada')

        if 'df_validades_mescladas' in st.session_state:
            st.success(
                f"✅ Validades carregadas pela **Central de Processamento** em "
                f"**{ultima_carga_val or 'horário não registrado'}**. Fonte: **{origem_val}**."
            )
        else:
            st.info(
                "⬆️ O Controle de Validade agora é carregado na aba **📥 Central de Processamento**, "
                "junto com o Estoque Geral e os arquivos de movimento. Esta aba fica reservada "
                "para consulta, filtros e exportações do painel FEFO."
            )

        n_descartados_farm = st.session_state.get('validade_descartados_farmacia_nao_mapeada', 0)
        if n_descartados_farm > 0:
            st.warning(
                f"⚠️ **{n_descartados_farm} registro(s)** da planilha de validade têm código e data "
                "válidos, mas a farmácia informada não pôde ser identificada (ex.: nome digitado "
                "de forma diferente do padrão, célula vazia, ou unidade fora das farmácias geridas "
                "pelo app, como 'Ambulatorial'). Esses registros não aparecem no painel FEFO. "
                "Revise a coluna de farmácia/unidade na planilha de validade, se necessário."
            )

        if st.button("🔁 Reaplicar filtro de saldo AGHU no painel de validades", use_container_width=True):
            if 'df_validades_mescladas' in st.session_state and 'est_geral_raw' in st.session_state:
                st.session_state['df_validades_mescladas'] = aplicar_saldo_atual_validades(
                    st.session_state['df_validades_mescladas'], st.session_state['est_geral_raw']
                )
                st.success("✅ Filtro de saldo AGHU reaplicado ao painel de validades.")
            elif 'df_validades_mescladas' not in st.session_state:
                st.warning("⚠️ Ainda não há controle de validade carregado nesta sessão.")
            else:
                st.warning("⚠️ Processe primeiro o Estoque Geral AGHU para aplicar o filtro de saldo atual.")

    # ── PAINEL DE VALIDADES ───────────────────────────────────────────────────
    if 'df_validades_mescladas' in st.session_state:
        df_val = st.session_state['df_validades_mescladas'].copy()

        # Guarda de segurança: se o Estoque Geral do AGHU já foi processado,
        # recalcula SEMPRE os saldos por código + farmácia antes de exibir o painel.
        # Isso evita manter em tela um DataFrame antigo da sessão, especialmente quando
        # o app foi atualizado de uma versão que ainda exibia "Saldo Atual/Fonte do Saldo"
        # ou quando as validades foram carregadas antes do estoque geral.
        est_geral_para_filtro = st.session_state.get('est_geral_raw')
        if est_geral_para_filtro is not None and not est_geral_para_filtro.empty:
            df_val_recalculado = aplicar_saldo_atual_validades(df_val, est_geral_para_filtro)
            st.session_state['df_validades_mescladas'] = df_val_recalculado
            df_val = df_val_recalculado.copy()

        hoje   = datetime.now().date()

        n_vencido  = (df_val['Situação'].str.contains('VENCIDO')).sum()
        n_critico  = (df_val['Situação'].str.contains('Crítico')).sum()
        n_atencao  = (df_val['Situação'].str.contains('Atenção')).sum()
        n_sem_data = (df_val['Situação'] == 'Sem data').sum()

        df_alerta_90 = df_val[df_val['Situação'].str.contains('VENCIDO|Crítico|Atenção', na=False)].copy()
        aghu_disponivel = bool(df_val.get('AGHU disponível para filtro', pd.Series([False])).fillna(False).any())
        saldo_aghu_num = pd.to_numeric(
            df_alerta_90.get('Saldo AGHU', pd.Series(dtype=float)), errors='coerce'
        ).fillna(0)
        saldo_planilha_num = pd.to_numeric(
            df_alerta_90.get('Saldo Planilha Validade', pd.Series(dtype=float)), errors='coerce'
        ).fillna(0)
        n_alerta_com_saldo_aghu = int((saldo_aghu_num > 0).sum()) if aghu_disponivel else 0
        n_alerta_com_saldo_planilha = int((saldo_planilha_num > 0).sum())

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("💀 Vencidos",           f"{n_vencido} itens",  delta_color="inverse",
                  delta="⚠️ Retirar imediatamente" if n_vencido else None)
        m2.metric("🔴 Críticos (≤30 dias)", f"{n_critico} itens")
        m3.metric("🟡 Atenção (31–90d)",    f"{n_atencao} itens")
        if aghu_disponivel:
            m4.metric("📦 Com saldo AGHU", f"{n_alerta_com_saldo_aghu} itens")
        else:
            m4.metric("📦 Saldo AGHU", "Não carregado")
        m5.metric("⚫ Sem data cadastrada", f"{n_sem_data} itens")

        if aghu_disponivel:
            st.caption(
                "📦 **Regra operacional ativa:** o painel e os relatórios mostram apenas itens "
                "vencidos ou a vencer que ainda possuem **Saldo AGHU > 0** na respectiva farmácia. "
                "A coluna **Saldo Planilha Validade** fica apenas como comparação histórica/manual."
            )
        else:
            st.warning(
                "⚠️ **Estoque Geral do AGHU ainda não foi processado nesta sessão.** "
                "O painel está usando apenas o Controle de Validade como referência provisória. "
                "Para ocultar automaticamente itens sem estoque atual, processe primeiro o arquivo "
                "de Estoque Geral na Aba 1."
            )
        st.write("---")

        # Filtros
        fv1, fv2, fv3 = st.columns(3)
        ops_farm = ["TODAS"] + sorted(df_val['Nome Farmácia'].unique().tolist())

        mapa_sem_val = {
            "TODAS": None,
            "💀 Vencidos": "💀",
            "🔴 Críticos (≤30 dias)": "🔴",
            "🟡 Atenção (31–90 dias)": "🟡",
            "🟢 OK (>90 dias)": "🟢",
            "⚫ Sem data": "⚫",
        }
        semaforos_presentes = set(df_val.get('🚦', pd.Series(dtype=str)).dropna().astype(str).tolist())
        ops_sit = ["TODAS"] + [rotulo for rotulo, emoji in mapa_sem_val.items()
                                if rotulo != "TODAS" and emoji in semaforos_presentes]

        ops_cat  = ["TODAS"] + sorted(df_val['Categoria'].unique().tolist())
        fil_farm = fv1.selectbox("Farmácia:", ops_farm, key="val_farm")
        fil_sit  = fv2.selectbox("Semáforo:", ops_sit,  key="val_sit")
        fil_cat  = fv3.selectbox("Categoria:", ops_cat, key="val_cat")

        df_vf = df_val.copy()
        if fil_farm != "TODAS":
            df_vf = df_vf[df_vf['Nome Farmácia'] == fil_farm]
        if fil_sit != "TODAS":
            emoji_sit = mapa_sem_val.get(fil_sit)
            if emoji_sit:
                df_vf = df_vf[df_vf['🚦'] == emoji_sit]
        if fil_cat != "TODAS":
            df_vf = df_vf[df_vf['Categoria'] == fil_cat]

        # A Categoria permanece disponível no filtro e nos relatórios, mas não polui a tabela da tela.
        cols_tela = ['🚦', 'Situação', 'key', 'Material', 'Lote',
                     'Validade Fmt', 'Saldo AGHU', 'Saldo Planilha Validade',
                     'Nome Farmácia', 'Fonte']
        cols_ok_tela = [c for c in cols_tela if c in df_vf.columns]

        cols_export = ['🚦', 'Situação', 'key', 'Material', 'Lote',
                       'Validade Fmt', 'Saldo AGHU', 'Saldo Planilha Validade',
                       'Nome Farmácia', 'Categoria', 'Fonte']
        cols_ok_export = [c for c in cols_export if c in df_vf.columns]

        st.dataframe(
            df_vf[cols_ok_tela].reset_index(drop=True),
            use_container_width=True, hide_index=True,
            column_config={
                'key':         st.column_config.TextColumn("Código MV", width="small"),
                'Material':    st.column_config.TextColumn("Material",   width="large"),
                'Lote':        st.column_config.TextColumn("Lote",       width="small"),
                'Validade Fmt':st.column_config.TextColumn("Validade",   width="small"),
                'Saldo AGHU':  st.column_config.NumberColumn("Saldo AGHU", width="small", format="%d"),
                'Saldo Planilha Validade': st.column_config.NumberColumn(
                    "Saldo Planilha Validade", width="medium", format="%d"
                ),
                'Nome Farmácia':st.column_config.TextColumn("Farmácia",  width="medium"),
                'Fonte':       st.column_config.TextColumn("Fonte Validade", width="small"),
            }
        )

        # Downloads
        dv1, dv2 = st.columns(2)
        df_criticos = df_val[df_val['🚦'].isin(['💀', '🔴', '🟡'])].copy()
        cols_ok_criticos = [c for c in cols_export if c in df_criticos.columns]
        with dv1:
            if not df_criticos.empty:
                gerar_download_excel_sob_demanda(
                    "📄 Gerar Vencidos + Até 90 dias (.xlsx)",
                    df_criticos[cols_ok_criticos].reset_index(drop=True),
                    "Ate_90_Dias",
                    f"Validades_Ate_90_Dias_{datetime.now().strftime('%d%m%y')}.xlsx",
                    key="validades_ate_90",
                    use_container_width=True,
                )
        with dv2:
            gerar_download_excel_sob_demanda(
                "📄 Gerar painel completo de validades (.xlsx)",
                df_vf[cols_ok_export].reset_index(drop=True),
                "Validades",
                f"Validades_Completo_{datetime.now().strftime('%d%m%y')}.xlsx",
                key="validades_completo",
                use_container_width=True,
            )
    else:
        st.info(
            "⬆️ Carregue o **Controle de Validade** na aba **📥 Central de Processamento**.\n\n"
            "💡 Dica: processe primeiro o Estoque Geral AGHU; assim o painel FEFO já "
            "oculta automaticamente os itens que não possuem saldo atual na farmácia correspondente."
        )


# =============================================================================
# TAB 3 — CONSOLIDAÇÃO MULTI-FARMÁCIA
# =============================================================================
with tab3:
    st.subheader("🏥 Consolidação Multi-Farmácia")
    st.info(
        "Usa o arquivo de estoque geral (AGDA2) e os arquivos de movimento já carregados "
        "na Aba 1 para gerar uma visão unificada de todas as farmácias."
    )

    msg_consolidacao = st.session_state.pop('consolidacao_msg_pos_rerun', '')
    if msg_consolidacao:
        st.success(msg_consolidacao)

    if 'est_geral_raw' not in st.session_state or 'mov_alvo_raw' not in st.session_state:
        st.warning(
            "⚠️ Processe os dados na **Aba 1** primeiro. "
            "A consolidação reutiliza os arquivos já carregados lá."
        )
    else:
        if st.button("🔄 GERAR CONSOLIDAÇÃO MULTI-FARMÁCIA", use_container_width=True, key="btn_consolida"):
            est_geral_raw = st.session_state['est_geral_raw']
            movs_raw      = st.session_state.get('movs_parceiras_raw', {})
            mov_alvo_raw  = st.session_state['mov_alvo_raw']
            cod_alvo      = st.session_state.get('cod_farmacia_alvo', '')
            data_ini      = st.session_state.get('data_inicio_huufma', datetime.now().date() - timedelta(days=7))
            data_fim      = st.session_state.get('data_fim_huufma', datetime.now().date() - timedelta(days=1))
            dias_ped      = st.session_state.get('dias_pedido_huufma', 15)
            mapa_cat      = obter_mapa_categorias()
            mapa_antimicrobianos = obter_mapa_antimicrobianos()

            # Montar mapa cod_farmacia → df_movimento
            movs_todos = {cod_alvo: mov_alvo_raw}
            movs_todos.update(movs_raw)

            resultados = []
            prog_c = st.progress(0, text="Analisando farmácias...")
            farms_com_mov = [c for c in CODIGOS_FARMACIAS if c in movs_todos]

            for i, cod in enumerate(farms_com_mov):
                prog_c.progress(
                    int((i / len(farms_com_mov)) * 100),
                    text=f"Analisando {DIC_NOMES_FARMACIAS.get(cod, cod)}..."
                )
                df_res = calcular_status_farmacia(
                    est_geral_raw, movs_todos[cod], cod,
                    mapa_cat, data_ini, data_fim, dias_ped,
                    mapa_antimicrobianos
                )
                if not df_res.empty:
                    resultados.append(df_res)

            if resultados:
                df_consolidado = pd.concat(resultados, ignore_index=True)
                st.session_state['df_consolidado'] = df_consolidado
                st.session_state['consolidacao_ultima_atualizacao'] = datetime.now().strftime('%d/%m/%Y %H:%M')
                prog_c.progress(100, text="✅ Consolidação concluída!")

                # Sincronização imediata com a aba Remanejamentos.
                # Como o Streamlit executa o script de cima para baixo, a aba Remanejamentos
                # é renderizada antes da Consolidação neste arquivo. Sem este bloco, a nova
                # consolidação só fica visível para os remanejamentos no rerun seguinte,
                # dando a impressão de que é necessário clicar em gerar consolidação várias vezes.
                invalidar_cache_remanejamentos()
                with st.spinner("🔄 Preparando remanejamentos com base na nova consolidação..."):
                    df_val_rem_sync = st.session_state.get('df_validades_mescladas', pd.DataFrame())
                    if (
                        isinstance(df_val_rem_sync, pd.DataFrame) and
                        not df_val_rem_sync.empty and
                        'est_geral_raw' in st.session_state
                    ):
                        df_val_rem_sync = aplicar_saldos_validades(
                            df_val_rem_sync.copy(), st.session_state['est_geral_raw']
                        )

                    obter_remanejamento_geral_session_cache(
                        df_consolidado,
                        df_val_rem_sync,
                        dias_cobertura=dias_ped,
                        forcar_recalculo=True,
                    )

                    if isinstance(df_val_rem_sync, pd.DataFrame) and not df_val_rem_sync.empty:
                        obter_remanejamento_validade_session_cache(
                            df_val_rem_sync,
                            df_consolidado,
                            forcar_recalculo=True,
                        )
                    else:
                        st.session_state['df_remanejamento_validade_huufma'] = pd.DataFrame()
                        st.session_state['remanejamento_validade_ultima_atualizacao'] = 'não aplicável — validade não carregada'

                st.session_state['consolidacao_msg_pos_rerun'] = (
                    "✅ Consolidação multi-farmácia concluída e remanejamentos sincronizados. "
                    "A aba Remanejamentos já pode ser consultada sem gerar a consolidação novamente."
                )
                st.rerun()
            else:
                st.error("❌ Nenhum dado pôde ser consolidado. Verifique os arquivos.")

        if 'df_consolidado' in st.session_state:
            df_cons = st.session_state['df_consolidado'].copy()
            if 'Antimicrobianos' not in df_cons.columns:
                mapa_antimicrobianos_cons = obter_mapa_antimicrobianos()
                df_cons['Antimicrobianos'] = df_cons['Código MV'].astype(str).apply(clean_key).map(mapa_antimicrobianos_cons).fillna('NÃO')
            df_cons['Antimicrobianos'] = df_cons['Antimicrobianos'].apply(normalizar_antimicrobiano)

            # Calcula os remanejamentos potenciais com a mesma inteligência da aba
            # Remanejamentos. O painel comparativo passa a contar oportunidades reais
            # por farmácia destino, em vez de depender apenas do parecer simplificado
            # da consolidação.
            rem_potencial_por_farm = {}
            df_rem_potencial_painel = pd.DataFrame()
            remanejamento_qtd_por_item_farm = {}
            try:
                df_val_rem_painel = st.session_state.get('df_validades_mescladas', pd.DataFrame())
                if (isinstance(df_val_rem_painel, pd.DataFrame) and
                        not df_val_rem_painel.empty and
                        'est_geral_raw' in st.session_state):
                    df_val_rem_painel = aplicar_saldos_validades(
                        df_val_rem_painel.copy(), st.session_state['est_geral_raw']
                    )

                with st.spinner("🔄 Atualizando indicador de remanejamento potencial..."):
                    df_rem_potencial_painel = obter_remanejamento_geral_session_cache(
                        df_cons,
                        df_val_rem_painel,
                        dias_cobertura=st.session_state.get('dias_pedido_huufma', 15),
                    )

                if (isinstance(df_rem_potencial_painel, pd.DataFrame) and
                        not df_rem_potencial_painel.empty and
                        {'Transferir PARA', 'Código MV'}.issubset(df_rem_potencial_painel.columns)):
                    rem_potencial_por_farm = (
                        df_rem_potencial_painel
                        .dropna(subset=['Transferir PARA'])
                        .groupby('Transferir PARA')['Código MV']
                        .nunique()
                        .to_dict()
                    )

                    # Mapa item + farmácia destino -> quantidade total remanejável sugerida.
                    # Esse mapa permite diferenciar necessidade crítica local de
                    # desabastecimento crítico real na consolidação. Se a necessidade
                    # da farmácia puder ser coberta por remanejamento interno suficiente,
                    # o item não deve permanecer no bloco de desabastecimento crítico real.
                    if 'Quantidade sugerida remanejar' in df_rem_potencial_painel.columns:
                        df_qtd_rem = df_rem_potencial_painel.copy()
                        df_qtd_rem['_key_rem_dest'] = df_qtd_rem['Código MV'].astype(str).apply(clean_key)
                        df_qtd_rem['_farm_dest'] = df_qtd_rem['Transferir PARA'].astype(str)
                        df_qtd_rem['_qtd_rem_dest'] = (
                            pd.to_numeric(df_qtd_rem['Quantidade sugerida remanejar'], errors='coerce')
                            .replace([np.inf, -np.inf], 0)
                            .fillna(0)
                            .clip(lower=0)
                        )
                        remanejamento_qtd_por_item_farm = (
                            df_qtd_rem[df_qtd_rem['_key_rem_dest'] != '']
                            .groupby(['_key_rem_dest', '_farm_dest'])['_qtd_rem_dest']
                            .sum()
                            .to_dict()
                        )
            except Exception as e:
                st.warning(
                    "⚠️ Não foi possível atualizar o indicador de remanejamento potencial "
                    f"no painel comparativo: {e}"
                )
                rem_potencial_por_farm = {}

            # Define desabastecimento crítico REAL na consolidação.
            # Regra: permanece crítico somente quando há necessidade na farmácia,
            # sem atendimento central suficiente e sem remanejamento interno suficiente.
            # A quantidade remanejável vem da mesma lógica da aba Remanejamentos.
            try:
                if 'Necessidade' not in df_cons.columns:
                    df_cons['Necessidade'] = 0
                df_cons['_key_critico_real'] = df_cons['Código MV'].astype(str).apply(clean_key)
                df_cons['_qtd_remanejamento_destino'] = df_cons.apply(
                    lambda r: float(remanejamento_qtd_por_item_farm.get(
                        (r['_key_critico_real'], str(r.get('Farmácia', ''))), 0
                    )),
                    axis=1
                )
                df_cons['_necessidade_num'] = (
                    pd.to_numeric(df_cons['Necessidade'], errors='coerce')
                    .replace([np.inf, -np.inf], 0)
                    .fillna(0)
                    .clip(lower=0)
                )
                df_cons['_Necessidade Residual Crítica'] = np.where(
                    df_cons['Parecer'].astype(str) == 'Desabastecimento Crítico',
                    np.maximum(df_cons['_necessidade_num'] - df_cons['_qtd_remanejamento_destino'], 0),
                    0
                )
                df_cons['_Desabastecimento Crítico Real'] = (
                    (df_cons['Parecer'].astype(str) == 'Desabastecimento Crítico') &
                    (df_cons['_Necessidade Residual Crítica'] > 0)
                )
            except Exception:
                df_cons['_qtd_remanejamento_destino'] = 0
                df_cons['_Necessidade Residual Crítica'] = 0
                df_cons['_Desabastecimento Crítico Real'] = df_cons['Parecer'].astype(str) == 'Desabastecimento Crítico'

            # ── KPIs POR FARMÁCIA ─────────────────────────────────────────────
            st.write("---")
            st.markdown("#### 📊 Painel Comparativo por Farmácia")

            painel_farms = []
            for cod, nome in DIC_NOMES_FARMACIAS.items():
                dff = df_cons[df_cons['Cód. Farmácia'] == cod]
                if dff.empty:
                    continue
                dff_vis = dff[dff['Parecer'] != 'Sem Consumo'].copy()
                dff_vis['_antimicrobiano_flag'] = dff_vis['Antimicrobianos'].apply(eh_antimicrobiano) if 'Antimicrobianos' in dff_vis.columns else False
                if '_Desabastecimento Crítico Real' in dff_vis.columns:
                    mask_desab = dff_vis['_Desabastecimento Crítico Real'].fillna(False).astype(bool)
                else:
                    mask_desab = dff_vis['Parecer'] == 'Desabastecimento Crítico'
                painel_farms.append({
                    'Farmácia': nome,
                    '📦 Itens': len(dff_vis),
                    '🚨 Desabastecimentos Críticos': mask_desab.sum(),
                    '🔴 Rupturas de antimicrobianos': (mask_desab & dff_vis['_antimicrobiano_flag']).sum(),
                    '🟠 Estoque Crítico no Almoxarifado': (dff_vis['Parecer'] == 'Estoque Crítico CAF').sum(),
                    '🔄 Remanejamento Potencial': int(rem_potencial_por_farm.get(nome, 0)),
                    '🩵 Excessos': (dff_vis['Parecer'] == 'Estoque Excessivo').sum(),
                    '🔘 Estoque Parado': (dff_vis['Parecer'] == 'Estoque Parado').sum(),
                })

            if painel_farms:
                df_painel = pd.DataFrame(painel_farms)
                st.dataframe(df_painel, use_container_width=True, hide_index=True)

            # ── DESABASTECIMENTOS CRÍTICOS ────────────────────────────────
            st.write("---")
            dias_ped_titulo = st.session_state.get('dias_pedido_huufma', 15)
            st.markdown(
                f"#### 🚨 Itens em Desabastecimento Crítico "
                f"(Considerando a necessidade dos próximos {dias_ped_titulo} dias)"
            )
            st.caption(
                "Lista apenas os itens com necessidade residual crítica: sem atendimento suficiente pelos "
                "almoxarifados centrais e sem remanejamento interno suficiente identificado. "
                "Itens resolvíveis por remanejamento saem deste bloco e permanecem na aba Remanejamentos."
            )

            if '_Desabastecimento Crítico Real' in df_cons.columns:
                total_critico_local = int((df_cons['Parecer'].astype(str) == 'Desabastecimento Crítico').sum())
                total_critico_real = int(df_cons['_Desabastecimento Crítico Real'].fillna(False).astype(bool).sum())
                total_resolvido_rem = max(total_critico_local - total_critico_real, 0)
                if total_resolvido_rem > 0:
                    st.info(
                        f"ℹ️ {total_resolvido_rem} ocorrência(s) de necessidade crítica local foram retiradas deste bloco "
                        "por possuírem remanejamento interno suficiente identificado."
                    )
                df_rupt_base = df_cons[df_cons['_Desabastecimento Crítico Real'].fillna(False).astype(bool)].copy()
            else:
                df_rupt_base = df_cons[df_cons['Parecer'] == 'Desabastecimento Crítico'].copy()
            if df_rupt_base.empty:
                st.success("✅ Nenhum item em desabastecimento crítico real nas farmácias analisadas.")
            else:
                df_rupt_base['Antimicrobianos'] = df_rupt_base.get('Antimicrobianos', 'NÃO')
                df_rupt_base['Antimicrobianos'] = df_rupt_base['Antimicrobianos'].apply(normalizar_antimicrobiano)

                # Estoque do item somado em TODOS os almoxarifados do arquivo de Estoque Geral.
                # Quando o arquivo bruto não estiver disponível ou não puder ser interpretado,
                # usa a soma dos saldos das farmácias consolidadas como fallback.
                estoque_todos_map = {}
                try:
                    est_all = st.session_state.get('est_geral_raw', pd.DataFrame()).copy()
                    c_all_cod = find_col(est_all, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod'])
                    c_all_qtd = find_col(est_all, ['qtde disp', 'disponivel'])
                    c_all_alm = find_col(est_all, ['almox'])
                    if c_all_alm:
                        est_all = filtrar_almoxarifados_excluidos(est_all, c_all_alm)
                    if not est_all.empty and c_all_cod and c_all_qtd:
                        est_all['_key_estoque_total'] = est_all[c_all_cod].apply(clean_key)
                        est_all['_saldo_estoque_total'] = p_num_series(est_all[c_all_qtd])
                        estoque_todos_map = (
                            est_all[est_all['_key_estoque_total'] != '']
                            .groupby('_key_estoque_total')['_saldo_estoque_total']
                            .sum()
                            .to_dict()
                        )
                except Exception:
                    estoque_todos_map = {}

                if not estoque_todos_map:
                    estoque_todos_map = (
                        df_cons.assign(_key_tmp=df_cons['Código MV'].astype(str).apply(clean_key))
                        .groupby('_key_tmp')['Saldo Atual']
                        .sum()
                        .to_dict()
                    )

                # CMD geral do item: soma do CMD das farmácias analisadas.
                # Não é média simples das farmácias; representa o consumo médio diário
                # hospitalar estimado para o item no conjunto de farmácias consolidadas.
                try:
                    if 'CMD' in df_cons.columns:
                        df_cmd_geral = df_cons.copy()
                        df_cmd_geral['_key_cmd_geral'] = df_cmd_geral['Código MV'].astype(str).apply(clean_key)
                        df_cmd_geral['_cmd_num'] = pd.to_numeric(df_cmd_geral['CMD'], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)
                        cmd_geral_map = (
                            df_cmd_geral[df_cmd_geral['_key_cmd_geral'] != '']
                            .groupby('_key_cmd_geral')['_cmd_num']
                            .sum()
                            .to_dict()
                        )
                    else:
                        cmd_geral_map = {}
                except Exception:
                    cmd_geral_map = {}

                # Filtros do tópico
                fcol_cat, fcol_farm, fcol_atb = st.columns([1.3, 1.5, 1.4])
                opcoes_cat_rupt = ["TODAS"] + sorted(df_rupt_base['Categoria'].dropna().astype(str).unique().tolist())
                opcoes_farm_rupt = ["TODAS"] + sorted(df_rupt_base['Farmácia'].dropna().astype(str).unique().tolist())

                f_cat_rupt = fcol_cat.selectbox(
                    "Categoria:",
                    opcoes_cat_rupt,
                    key="filtro_rupturas_categoria",
                )
                f_farm_rupt = fcol_farm.selectbox(
                    "Farmácia afetada:",
                    opcoes_farm_rupt,
                    key="filtro_rupturas_farmacia",
                )
                f_atb_rupt = fcol_atb.selectbox(
                    "Antimicrobianos:",
                    ["TODOS", "SOMENTE ANTIMICROBIANOS", "NÃO ANTIMICROBIANOS"],
                    key="filtro_rupturas_antimicrobianos",
                )

                df_rupt = df_rupt_base.copy()
                if f_cat_rupt != "TODAS":
                    df_rupt = df_rupt[df_rupt['Categoria'].astype(str) == f_cat_rupt]
                if f_farm_rupt != "TODAS":
                    df_rupt = df_rupt[df_rupt['Farmácia'].astype(str) == f_farm_rupt]
                if f_atb_rupt == "SOMENTE ANTIMICROBIANOS":
                    df_rupt = df_rupt[df_rupt['Antimicrobianos'].apply(eh_antimicrobiano)]
                elif f_atb_rupt == "NÃO ANTIMICROBIANOS":
                    df_rupt = df_rupt[~df_rupt['Antimicrobianos'].apply(eh_antimicrobiano)]

                if df_rupt.empty:
                    st.info("Nenhum item em desabastecimento crítico real para os filtros selecionados.")
                else:
                    tabela_rupt = df_rupt.groupby('Código MV').agg(
                        ITEM=('Material', 'first'),
                        CATEGORIA=('Categoria', 'first'),
                        ANTIMICROBIANO=('Antimicrobianos', 'first'),
                        **{'FARMÁCIA AFETADA': ('Farmácia', lambda x: ' | '.join(sorted(set(x.astype(str)))))}
                    ).reset_index().rename(columns={'Código MV': 'CÓDIGO'})

                    tabela_rupt['ESTOQUE EM TODOS OS ALMOXARIFADOS'] = (
                        tabela_rupt['CÓDIGO']
                        .astype(str)
                        .apply(clean_key)
                        .map(estoque_todos_map)
                        .fillna(0)
                        .round(0)
                        .astype(int)
                    )

                    tabela_rupt['CMD GERAL'] = (
                        tabela_rupt['CÓDIGO']
                        .astype(str)
                        .apply(clean_key)
                        .map(cmd_geral_map)
                        .fillna(0)
                        .astype(float)
                    )
                    tabela_rupt['COBERTURA GERAL ESTIMADA (DIAS)'] = np.where(
                        tabela_rupt['CMD GERAL'] > 0,
                        tabela_rupt['ESTOQUE EM TODOS OS ALMOXARIFADOS'] / tabela_rupt['CMD GERAL'],
                        np.nan
                    )
                    tabela_rupt['CMD GERAL'] = tabela_rupt['CMD GERAL'].round(2)
                    tabela_rupt['COBERTURA GERAL ESTIMADA (DIAS)'] = (
                        pd.to_numeric(tabela_rupt['COBERTURA GERAL ESTIMADA (DIAS)'], errors='coerce')
                        .replace([np.inf, -np.inf], np.nan)
                        .round(1)
                    )

                    tabela_rupt = tabela_rupt[
                        ['CÓDIGO', 'ITEM', 'CATEGORIA', 'ESTOQUE EM TODOS OS ALMOXARIFADOS',
                         'CMD GERAL', 'COBERTURA GERAL ESTIMADA (DIAS)',
                         'ANTIMICROBIANO', 'FARMÁCIA AFETADA']
                    ].sort_values(['ITEM', 'CÓDIGO']).reset_index(drop=True)

                    filtros_aplicados = any([
                        f_cat_rupt != "TODAS",
                        f_farm_rupt != "TODAS",
                        f_atb_rupt != "TODOS",
                    ])
                    st.error(
                        f"⚠️ {len(tabela_rupt)} item(ns) em desabastecimento crítico real "
                        f"{'nos filtros aplicados' if filtros_aplicados else 'nas farmácias analisadas'}."
                    )
                    st.dataframe(
                        tabela_rupt,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            'ESTOQUE EM TODOS OS ALMOXARIFADOS': st.column_config.NumberColumn(
                                'ESTOQUE EM TODOS OS ALMOXARIFADOS', format='%d'
                            ),
                            'CMD GERAL': st.column_config.NumberColumn(
                                'CMD GERAL',
                                help='Soma do CMD das farmácias analisadas para o item.',
                                format='%.2f'
                            ),
                            'COBERTURA GERAL ESTIMADA (DIAS)': st.column_config.NumberColumn(
                                'COBERTURA GERAL ESTIMADA (DIAS)',
                                help='Estoque em todos os almoxarifados dividido pelo CMD geral.',
                                format='%.1f'
                            ),
                        }
                    )

                    # Exportação: mantém OUTROS na visualização do aplicativo,
                    # mas exclui essa categoria do arquivo gerado, conforme regra operacional.
                    tabela_rupt_export = tabela_rupt[
                        tabela_rupt['CATEGORIA'].astype(str).str.upper().str.strip() != 'OUTROS'
                    ].copy()

                    if tabela_rupt_export.empty:
                        st.warning(
                            "⚠️ O resultado exibido possui apenas itens da categoria OUTROS. "
                            "Por regra, a categoria OUTROS não será incluída no relatório exportado."
                        )
                    else:
                        gerar_download_excel_sob_demanda(
                            "📄 Gerar desabastecimentos críticos reais exibidos (.xlsx)",
                            tabela_rupt_export,
                            "Desabastecimentos Críticos Reais",
                            f"Desabastecimentos_Criticos_Reais_{datetime.now().strftime('%d%m%y')}.xlsx",
                            key="desabastecimentos_criticos_exibidos",
                            use_container_width=True,
                            help="A visualização mantém a categoria OUTROS, mas o relatório exportado exclui esses itens.",
                        )

            # ── REMANEJAMENTOS ────────────────────────────────────────────
            st.write("---")
            st.info(
                "🔄 As análises de remanejamento geral e remanejamento preventivo por validade "
                "foram movidas para a aba **Remanejamentos**, mantendo a consolidação aqui como "
                "painel comparativo multi-farmácia."
            )

            # ── VALIDADES CRÍTICAS NA CONSOLIDAÇÃO ───────────────────────────
            if 'df_validades_mescladas' in st.session_state:
                st.write("---")
                st.markdown("#### ⏰ Validades por Farmácia — Semáforo FEFO")
                df_vc = st.session_state['df_validades_mescladas'].copy()
                if df_vc.empty:
                    st.info("Nenhum dado de validade carregado para compor o semáforo.")
                else:
                    def _classe_semaforo(situacao: str) -> str:
                        s = str(situacao)
                        if 'VENCIDO' in s:
                            return '💀 Vencidos'
                        if 'Crítico' in s:
                            return '🔴 Críticos ≤30 dias'
                        if 'Atenção' in s:
                            return '🟡 Atenção 31–90 dias'
                        if 'OK' in s:
                            return '🟢 OK >90 dias'
                        return '⚫ Sem data'

                    ordem_sem = [
                        '💀 Vencidos', '🔴 Críticos ≤30 dias', '🟡 Atenção 31–90 dias',
                        '🟢 OK >90 dias', '⚫ Sem data'
                    ]
                    df_vc['_Semáforo'] = df_vc['Situação'].apply(_classe_semaforo)
                    vc_farm = pd.pivot_table(
                        df_vc,
                        index='Nome Farmácia',
                        columns='_Semáforo',
                        values='key',
                        aggfunc='count',
                        fill_value=0,
                    ).reset_index().rename(columns={'Nome Farmácia': 'Farmácia'})
                    for col in ordem_sem:
                        if col not in vc_farm.columns:
                            vc_farm[col] = 0
                    vc_farm = vc_farm[['Farmácia'] + ordem_sem]
                    vc_farm = vc_farm.sort_values(
                        ['💀 Vencidos', '🔴 Críticos ≤30 dias', '🟡 Atenção 31–90 dias', 'Farmácia'],
                        ascending=[False, False, False, True]
                    )
                    st.dataframe(vc_farm, use_container_width=True, hide_index=True)

            # ── DOWNLOAD CONSOLIDAÇÃO ─────────────────────────────────────────
            st.write("---")
            gerar_download_multi_aba_sob_demanda(
                "📄 Gerar Consolidação Completa (.xlsx)",
                df_cons,
                ['Código MV', 'Material', 'Categoria', 'Antimicrobianos', 'Farmácia',
                 'Saldo Atual', 'CMD', 'Estoque Mínimo', 'Cobertura (dias)',
                 'Necessidade', 'Saldo Central', 'Parecer'],
                col_categoria='Farmácia',
                col_alerta='Parecer',
                larguras={'Código MV': 12, 'Material': 45, 'Categoria': 16, 'Antimicrobianos': 16,
                          'Farmácia': 28, 'Saldo Atual': 14, 'CMD': 12,
                          'Estoque Mínimo': 14, 'Cobertura (dias)': 14,
                          'Necessidade': 14, 'Saldo Central': 18, 'Parecer': 28},
                file_name=f"Consolidacao_{datetime.now().strftime('%d%m%y')}.xlsx",
                key="consolidacao_completa",
                use_container_width=True,
            )


# =============================================================================
# TAB 4 — GESTÃO DE CATEGORIAS
# =============================================================================
with tab4:
    st.subheader("🗂️ Mapeamento Global de Categorias")
    st.info(
        "A tabela serve como De/Para para classificar os itens no momento da análise. "
        "A edição oficial das categorias deve ser feita diretamente no Google Sheets."
        )

    with st.container(border=True):
        st.markdown("##### 🔗 Fonte das Categorias")
        fonte_cat = st.session_state.get("categorias_fonte", "Não identificada")
        ultima_carga_cat = st.session_state.get("categorias_ultima_carga", "—")
        erro_carga_cat = st.session_state.get("categorias_erro_carga", "")
        erro_local_cat = st.session_state.get("categorias_erro_local", "")
        msg_cat = st.session_state.pop("categorias_msg_salvamento", "")

        if fonte_cat == "Google Sheets público":
            st.success(f"✅ Categorias carregadas do **Google Sheets público** em **{ultima_carga_cat}**.")
        else:
            st.warning(f"⚠️ Categorias carregadas de: **{fonte_cat}**. Última carga: **{ultima_carga_cat}**.")

        if erro_carga_cat:
            st.error(f"Erro ao carregar do Google Sheets: {erro_carga_cat}")
        if erro_local_cat:
            st.error(erro_local_cat)
        if msg_cat:
            st.info(msg_cat)

        st.caption(
            "A edição oficial das categorias deve ser feita diretamente no Google Sheets. "
            "Depois de alterar a planilha, clique em 'Carregar categorias do Google Sheets' "
            "para atualizar os dados usados pelo aplicativo. Se o Google Sheets falhar, "
            "o app usa Categorias_base.xlsx apenas como fallback de leitura."
        )

        gs1, gs2, gs3 = st.columns([1, 1, 1])
        if gs1.button("🔄 Carregar categorias do Google Sheets", use_container_width=True):
            df_gs, erro_gs = carregar_categorias_google_sheets_publico()
            if df_gs is not None and not df_gs.empty:
                st.session_state["df_categorias"] = df_gs
                st.session_state["categorias_fonte"] = "Google Sheets público"
                st.session_state["categorias_ultima_carga"] = datetime.now().strftime('%d/%m/%Y %H:%M')
                st.session_state["categorias_erro_carga"] = ""
                st.session_state["categorias_msg_salvamento"] = "✅ Categorias carregadas novamente do Google Sheets público."
                st.rerun()
            else:
                st.session_state["categorias_erro_carga"] = erro_gs or "Google Sheets público não retornou dados válidos."
                st.rerun()

        if gs2.button("🧪 Testar conexão", use_container_width=True):
            df_gs, erro_gs = carregar_categorias_google_sheets_publico()
            if df_gs is not None:
                st.success(f"✅ Conexão OK. {len(df_gs)} categoria(s) encontradas no Google Sheets.")
            else:
                st.error(erro_gs or "Falha não especificada ao testar o Google Sheets.")

        gs3.markdown(f"[🔗 Abrir planilha Google]({GOOGLE_SHEETS_CATEGORIAS_LINK})")

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
                df_filtrado_cat["Código"].astype(str).str.contains(filtro_termo_cat, case=False, na=False, regex=False) |
                df_filtrado_cat["Material"].astype(str).str.contains(filtro_termo_cat, case=False, na=False, regex=False)
            ]
        if filtro_sel_cat != "TODOS":
            df_filtrado_cat = df_filtrado_cat[df_filtrado_cat["Categoria"] == filtro_sel_cat]

        st.markdown(f"##### 📋 Itens Encontrados ({len(df_filtrado_cat)} registros)")
        st.caption(
            "Visualização somente leitura. Para corrigir Código, Material, Categoria ou Antimicrobianos, "
            "edite diretamente a planilha Google e depois clique em 'Carregar categorias do Google Sheets'."
        )

        cols_cat_exibir = [c for c in ["Código", "Material", "Categoria", "Antimicrobianos"] if c in df_filtrado_cat.columns]
        st.dataframe(
            df_filtrado_cat[cols_cat_exibir].reset_index(drop=True),
            use_container_width=True, hide_index=True,
            column_config={
                "Código": st.column_config.TextColumn("Código MV", width="small"),
                "Material": st.column_config.TextColumn("Descrição do Insumo", width="large"),
                "Categoria": st.column_config.TextColumn("Categoria Logística", width="medium"),
                "Antimicrobianos": st.column_config.TextColumn(
                    "Antimicrobianos", width="small",
                    help="Campo preservado para futuras análises de antimicrobianos. Sugestão: SIM ou NÃO."
                ),
            }
        )

        # Download de itens OUTROS para classificação em lote / revisão na planilha Google
        df_outros = df_cat_atual[df_cat_atual["Categoria"] == "OUTROS"]
        if not df_outros.empty:
            gerar_download_excel_sob_demanda(
                "📄 Gerar itens OUTROS para revisão no Google Sheets",
                df_outros,
                "Para Revisar",
                f"Itens_OUTROS_{datetime.now().strftime('%d%m%y')}.xlsx",
                key="categorias_outros_revisao",
                use_container_width=True,
                help="Use este arquivo como apoio para revisar a planilha Google. A edição oficial deve ser feita no Sheets.",
            )
    else:
        st.warning("🔍 Utilize os filtros acima para visualizar os registros.")
        st.info("Para alterar categorias, abra a planilha Google, edite os dados e recarregue as categorias no aplicativo.")

    # Botão exportar mapa completo — sempre disponível, fora do if/else
    st.write("")
    df_completo_export = st.session_state["df_categorias"].copy()
    if not df_completo_export.empty:
        gerar_download_excel_sob_demanda(
            "📄 Gerar mapa COMPLETO de categorias (.xlsx)",
            df_completo_export,
            "Categorias",
            f"Categorias_base_{datetime.now().strftime('%d%m%y')}.xlsx",
            key="categorias_mapa_completo",
            use_container_width=True,
            help="Backup da base atualmente carregada. A edição oficial deve ser realizada no Google Sheets.",
        )


# =============================================================================
# TAB 1 — CENTRAL DE PROCESSAMENTO
# =============================================================================
with tab1:
    with st.container(border=True):
        st.markdown("##### 📥 Central de Processamento das Fontes de Dados")
        col1, col2 = st.columns(2)
        file_mov_alvo  = col1.file_uploader("1. Movimento da Farmácia Alvo (.csv)", type=["csv"])
        file_est_geral = col2.file_uploader("2. Estoque Geral de todos os Almoxarifados (.csv)", type=["csv"])
        st.write("")
        files_mov_parceiras = st.file_uploader(
            "3. Movimentos das outras Farmácias — Ativa a análise da viabilidade de remanejamentos (opcional) (Múltiplos.csv)",
            type=["csv"], accept_multiple_files=True,
        )

    with st.container(border=True):
        st.markdown("##### ⏰ Controle de Validade — fonte complementar recomendada")
        st.caption(
            "Carregue aqui a planilha de validade para que o pedido, o painel FEFO e os "
            "remanejamentos preventivos já sejam processados de forma integrada."
        )

        status_val = st.session_state.get('validade_ultima_carga')
        origem_val = st.session_state.get('validade_ultima_origem', '')
        if 'df_validades_mescladas' in st.session_state:
            st.success(
                f"✅ Validades carregadas em **{status_val or 'horário não registrado'}**"
                + (f" — Fonte: **{origem_val}**" if origem_val else "")
            )
        else:
            st.info("Controle de validade ainda não carregado nesta sessão.")

        cv1, cv2 = st.columns([2, 1])
        file_val_manual_central = cv1.file_uploader(
            "4. Controle de Validade (.xlsx) — opcional, mas recomendado",
            type=['xlsx', 'xls'],
            key='upload_validades_central',
            help="Se este arquivo for enviado, ele será usado como fonte de validade. Se não for enviado, o botão tentará carregar via SharePoint."
        )
        carregar_val_central = cv2.button(
            "🔄 Carregar / Atualizar Validades",
            use_container_width=True,
            key='btn_carregar_validades_central'
        )

        if carregar_val_central:
            df_sp_raw_central = None
            origem_central = ""

            if file_val_manual_central is not None:
                try:
                    file_val_manual_central.seek(0)
                    df_sp_raw_central = carregar_planilha_validades_excel(
                        io.BytesIO(file_val_manual_central.read())
                    )
                    origem_central = f"upload manual ({file_val_manual_central.name})"
                except Exception as e:
                    st.error(f"❌ Erro ao ler o arquivo manual de validade: {e}")
            else:
                with st.spinner("🌐 Tentando carregar validades do SharePoint..."):
                    df_sp_raw_central, erro_sp_central = carregar_validades_sharepoint()
                origem_central = "SharePoint"
                if df_sp_raw_central is None:
                    st.warning(
                        f"⚠️ Não foi possível acessar o SharePoint: {erro_sp_central}\n\n"
                        "Carregue o arquivo manualmente no campo acima e clique novamente em atualizar."
                    )

            if df_sp_raw_central is not None:
                ok_val, msg_val = processar_validades_para_sessao(df_sp_raw_central, origem_central)
                if ok_val:
                    st.success(f"✅ {msg_val}")
                else:
                    st.error(f"❌ {msg_val}")

    st.write("")

    if file_mov_alvo and file_est_geral:
        if st.button("🚀 ANALISAR OS DADOS COM INTELIGÊNCIA LOGÍSTICA", use_container_width=True):
            _limpar_resultados_derivados()
            st.session_state['disparar_processamento_huufma'] = True

        if st.session_state.get('disparar_processamento_huufma', False):
            progress = st.progress(0, text="📂 Lendo arquivos...")

            try:
                file_mov_alvo.seek(0)
                file_est_geral.seek(0)
                mov       = ler_csv_cached(file_mov_alvo.read(), file_mov_alvo.name)
                est_geral = ler_csv_cached(file_est_geral.read(), file_est_geral.name)

                # Salvar raws para uso nas abas Validade e Consolidação
                st.session_state['est_geral_raw']  = est_geral
                st.session_state['mov_alvo_raw']   = mov

                # Controle de validade integrado à Central de Processamento:
                # se o usuário anexou o arquivo de validade na Central, ele é processado
                # junto com o Estoque Geral, já aplicando o filtro Saldo AGHU > 0.
                if 'file_val_manual_central' in locals() and file_val_manual_central is not None:
                    try:
                        file_val_manual_central.seek(0)
                        df_val_upload_proc = carregar_planilha_validades_excel(
                            io.BytesIO(file_val_manual_central.read())
                        )
                        ok_val_proc, msg_val_proc = processar_validades_para_sessao(
                            df_val_upload_proc,
                            f"upload manual ({file_val_manual_central.name})"
                        )
                        if ok_val_proc:
                            st.toast("✅ Controle de validade integrado ao processamento.", icon="⏰")
                        else:
                            st.warning(f"⚠️ Controle de validade não processado: {msg_val_proc}")
                    except Exception as e:
                        st.warning(f"⚠️ Controle de validade ignorado por erro na leitura: {e}")
                elif 'df_validades_mescladas' in st.session_state:
                    # Se as validades já foram carregadas antes do Estoque Geral,
                    # recalcula os saldos de validade e aplica o filtro operacional: Saldo AGHU > 0.
                    st.session_state['df_validades_mescladas'] = aplicar_saldo_atual_validades(
                        st.session_state['df_validades_mescladas'], est_geral
                    )

                progress.progress(10, text="🔍 Identificando colunas...")

                cols_mov = {
                    'código':       find_col(mov, ['material', 'cod', 'ca3']),
                    'quantidade':   find_col(mov, ['quant']),
                    'tipo':         find_col(mov, ['tipo']),
                    'data':         find_col(mov, ['data geracao', 'data mov', 'data', 'dt ger', 'dtger'],
                                            forbidden=['almox', 'tipo', 'quant', 'material']),
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

                # Regra global: excluir o almoxarifado 45 antes de qualquer cálculo.
                # Ele contém estoque vinculado a pacientes e não deve influenciar pedido,
                # consumo, validade, consolidação, remanejamentos ou relatórios.
                mov = filtrar_almoxarifados_excluidos(mov, c_mov_almox)
                est_geral = filtrar_almoxarifados_excluidos(est_geral, c_est_almox)
                st.session_state['est_geral_raw'] = est_geral
                st.session_state['mov_alvo_raw'] = mov

                if mov.empty:
                    st.error(
                        "❌ O arquivo de movimento ficou vazio após excluir almoxarifados fora da gestão do app "
                        "(ex.: Almoxarifado 45 - estoque de pacientes). Verifique se o arquivo enviado é de uma farmácia gerenciada."
                    )
                    st.session_state['disparar_processamento_huufma'] = False
                    st.stop()

                # --- TRAVA INTELIGENTE: DETECÇÃO AUTOMÁTICA DA FARMÁCIA ALVO ---
                codigos_almox = mov[c_mov_almox].dropna().astype(str).apply(clean_key)
                modas_almox = codigos_almox[codigos_almox != ""].mode()
                if modas_almox.empty:
                    st.error(
                        "❌ Não foi possível detectar a farmácia alvo: a coluna de "
                        "almoxarifado do arquivo de movimento está vazia ou inválida."
                    )
                    st.session_state['disparar_processamento_huufma'] = False
                    st.stop()
                cod_farmacia_alvo = modas_almox[0]
                if almoxarifado_excluido(cod_farmacia_alvo):
                    st.error(
                        "❌ O almoxarifado detectado é o 45, que contém estoque de pacientes e foi excluído da lógica do aplicativo. "
                        "Envie o movimento de uma farmácia/almoxarifado gerenciado pela UDIS."
                    )
                    st.session_state['disparar_processamento_huufma'] = False
                    st.stop()
                
                st.session_state['cod_farmacia_alvo'] = cod_farmacia_alvo
                st.session_state['nome_farmacia_alvo'] = DIC_NOMES_FARMACIAS.get(cod_farmacia_alvo, f"Almoxarifado (Cód. {cod_farmacia_alvo})")

                progress.progress(20, text="🏗️ Processando estoque geral...")

                est_geral = est_geral.copy()
                est_geral['key']         = est_geral[c_est_cod].apply(clean_key)
                est_geral['almox_limpo'] = est_geral[c_est_almox].apply(clean_key)
                est_geral['saldo_num']   = p_num_series(est_geral[c_est_qtd])
                est_geral['min_num']     = p_num_series(est_geral[c_est_min]) if c_est_min else 0.0
                est_geral = est_geral[~est_geral['almox_limpo'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()

                # Saldo: soma das linhas. Mínimo: parâmetro único por item — usar max
                # evita duplicar o mínimo quando o item aparece em mais de uma linha.
                _est_alvo = (
                    est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo]
                    .groupby('key')
                    .agg(saldo_num=('saldo_num', 'sum'), min_num=('min_num', 'max'))
                )
                est_farmacia_alvo = _est_alvo['saldo_num'].to_dict()
                est_min_alvo      = _est_alvo['min_num'].to_dict()

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
                # Formato AGHUx: AAAA-MM-DD HH:MM — dayfirst=False (padrão ISO)
                mov['dt_formatada'] = pd.to_datetime(mov[c_mov_data], dayfirst=False, errors='coerce')

                # Auditoria de formato: se muitas datas viraram NaT, o arquivo
                # provavelmente está em DD/MM/AAAA e o CMD sairia errado.
                pct_nat = mov['dt_formatada'].isna().mean() if len(mov) else 0.0
                if pct_nat > 0.2:
                    st.warning(
                        f"⚠️ **Auditoria de datas:** {pct_nat:.0%} das datas do movimento "
                        "não foram reconhecidas (formato esperado: AAAA-MM-DD). "
                        "Verifique o formato de exportação do AGHU antes de confiar no CMD."
                    )
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
                    .assign(qtd_num=lambda df: p_num_series(df[c_mov_qtd]))
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
                    mov_filtrado, c_mov_cod, c_mov_qtd,
                    n_dias=JANELA_TENDENCIA_DIAS
                )

                progress.progress(55, text="🔄 Cruzando dados entre farmácias...")

                consumo_outras_total = {}
                if files_mov_parceiras:
                    for f_parc in files_mov_parceiras:
                        try:
                            f_parc.seek(0)
                            df_p = ler_csv_cached(f_parc.read(), f_parc.name)
                            c_p_cod   = find_col(df_p, ['material', 'cod', 'ca3'])
                            c_p_qtd   = find_col(df_p, ['quant'])
                            c_p_tipo  = find_col(df_p, ['tipo'])
                            c_p_almox = find_col(df_p, ['almox'])
                            c_p_data  = find_col(df_p, ['data geracao', 'data mov', 'data', 'dt ger', 'dtger'],
                                                        forbidden=['almox', 'tipo', 'quant', 'material'])

                            if not all([c_p_cod, c_p_qtd, c_p_tipo, c_p_almox, c_p_data]):
                                st.warning(f"⚠️ '{f_parc.name}' ignorado: colunas não identificadas.")
                                continue

                            df_p = df_p.copy()
                            # Formato AGHUx: AAAA-MM-DD HH:MM — dayfirst=False (padrão ISO)
                            df_p['dt_formatada'] = pd.to_datetime(
                                df_p[c_p_data], dayfirst=False, errors='coerce'
                            )
                            df_p_filt = df_p[
                                (df_p['dt_formatada'].dt.date >= data_inicio) &
                                (df_p['dt_formatada'].dt.date <= data_fim) &
                                (df_p[c_p_tipo].astype(str).str.upper() == 'RM')
                            ].assign(
                                key=lambda d: d[c_p_cod].apply(clean_key),
                                almox_limpo=lambda d: d[c_p_almox].apply(clean_key),
                                qtd_num=lambda d: p_num_series(d[c_p_qtd]),
                            )

                            df_p_filt = df_p_filt[~df_p_filt['almox_limpo'].isin(CODIGOS_ALMOXARIFADOS_EXCLUIR)].copy()

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

                # Salvar movimentos das parceiras para uso na consolidação
                if files_mov_parceiras:
                    movs_parceiras_raw = {}
                    for f_parc in files_mov_parceiras:
                        try:
                            f_parc.seek(0)
                            df_tmp = ler_csv_cached(f_parc.read(), f_parc.name)
                            c_tmp_almox = find_col(df_tmp, ['almox'])
                            if c_tmp_almox:
                                cod_tmp = df_tmp[c_tmp_almox].dropna().astype(str).apply(clean_key).mode()
                                if not cod_tmp.empty:
                                    cod_detectado_tmp = cod_tmp[0]
                                    if almoxarifado_excluido(cod_detectado_tmp):
                                        continue
                                    df_tmp = filtrar_almoxarifados_excluidos(df_tmp, c_tmp_almox)
                                    movs_parceiras_raw[cod_detectado_tmp] = df_tmp
                        except Exception:
                            pass
                    st.session_state['movs_parceiras_raw'] = movs_parceiras_raw

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
                final['Necessidade de Ressuprimento'] = calcular_sugestao_vetorizado(final, dias_pedido)

                final['Cobertura (dias)'] = calcular_cobertura_vetorizado(final)

                final['CMD Últ. 3 dias'] = final['Código MV'].map(tendencia_map).fillna(0)
                # result_type='expand' evita desempacotar tuplas em dois loops extras
                final[['Tendência', 'Δ% Tendência']] = final.apply(
                    calcular_delta_tendencia, axis=1, result_type='expand'
                )

                mapa_cat = obter_mapa_categorias()
                mapa_antimicrobianos = obter_mapa_antimicrobianos()
                final['Categoria'] = final['Código MV'].map(mapa_cat).fillna('OUTROS')
                final['Antimicrobianos'] = final['Código MV'].map(mapa_antimicrobianos).fillna('NÃO').apply(normalizar_antimicrobiano)

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

                # ── INTEGRAÇÃO COM VALIDADES (se já carregadas na Aba 2) ────────
                final['⏰ Validade'] = "—"  # placeholder padrão quando não há dado de validade
                if 'df_validades_mescladas' in st.session_state:
                    df_val_atual = st.session_state['df_validades_mescladas']
                    cod_farm_atual = st.session_state.get('cod_farmacia_alvo', '')
                    # Filtra validades da farmácia ativa e pega a situação mais crítica por item
                    df_val_farm = df_val_atual[df_val_atual['Farmácia'] == cod_farm_atual]
                    if not df_val_farm.empty:
                        ordem_sit = {'VENCIDO': 0, 'Crítico': 1, 'Atenção': 2, 'OK': 3, 'Sem data': 4}
                        df_val_farm = df_val_farm.copy()
                        df_val_farm['_ord'] = df_val_farm['Situação'].apply(
                            lambda s: next((v for k, v in ordem_sit.items() if k in s), 5)
                        )
                        # Mantém só a situação mais crítica por código quando há múltiplos lotes
                        mapa_val = (df_val_farm.sort_values('_ord')
                                    .drop_duplicates('key')
                                    .set_index('key')['🚦'] + ' ' +
                                    df_val_farm.sort_values('_ord').drop_duplicates('key')
                                    .set_index('key')['Situação'])
                        final['⏰ Validade'] = final['Código MV'].map(mapa_val).fillna("—")

                st.session_state['df_final_huufma']        = final
                st.session_state['n_dias_efetivos_huufma'] = n_dias_efetivos
                st.session_state['disparar_processamento_huufma'] = False
                progress.progress(100, text="✅ Processamento concluído!")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Erro crítico no processamento: {e}")
                st.session_state['disparar_processamento_huufma'] = False

        # ── ATUALIZAÇÃO AUTOMÁTICA DE VALIDADES NO RESULTADO EXISTENTE ──────────
        # Se as validades foram carregadas DEPOIS do processamento, atualiza a coluna
        # sem precisar reprocessar tudo novamente
        if ('df_final_huufma' in st.session_state and
                'df_validades_mescladas' in st.session_state):
            df_atual = st.session_state['df_final_huufma']
            df_val_ss = st.session_state['df_validades_mescladas']
            cod_farm_atual = st.session_state.get('cod_farmacia_alvo', '')
            # Só atualiza se ainda tiver a coluna placeholder '—' para todos
            if '⏰ Validade' not in df_atual.columns or df_atual['⏰ Validade'].eq('—').all():
                df_val_farm = df_val_ss[df_val_ss['Farmácia'] == cod_farm_atual]
                if not df_val_farm.empty:
                    ordem_sit = {'VENCIDO': 0, 'Crítico': 1, 'Atenção': 2, 'OK': 3, 'Sem data': 4}
                    df_val_farm = df_val_farm.copy()
                    df_val_farm['_ord'] = df_val_farm['Situação'].apply(
                        lambda s: next((v for k, v in ordem_sit.items() if k in s), 5)
                    )
                    df_sorted = df_val_farm.sort_values('_ord').drop_duplicates('key')
                    mapa_val = df_sorted.set_index('key')['🚦'] + ' ' + df_sorted.set_index('key')['Situação']
                    df_atual['⏰ Validade'] = df_atual['Código MV'].map(mapa_val).fillna('—')
                    st.session_state['df_final_huufma'] = df_atual
        if 'df_final_huufma' in st.session_state:
            st.success(
                '✅ Dados processados com sucesso. A análise operacional completa está disponível na aba **📦 Pedido da Farmácia Ativa**.'
            )
            st.info(
                'Esta aba permanece como central de entrada dos arquivos. Use as abas de resultado para visualizar, filtrar e exportar as análises.'
            )

    else:
        st.session_state['disparar_processamento_huufma'] = False
        if 'df_final_huufma' in st.session_state:
            st.success(
                '✅ Dados já processados nesta sessão. A análise operacional completa está disponível na aba **📦 Pedido da Farmácia Ativa**.'
            )
            st.info(
                'Para processar uma nova extração do AGHU, carregue novamente o movimento da farmácia alvo e o estoque geral. '
                'Recarregar categorias do Google Sheets não apaga a análise já processada.'
            )
        elif 'est_geral_raw' in st.session_state and 'mov_alvo_raw' in st.session_state:
            st.info(
                'ℹ️ Arquivos obrigatórios já foram lidos nesta sessão, mas a análise final ainda não está disponível. '
                'Clique em **ANALISAR OS DADOS COM INTELIGÊNCIA LOGÍSTICA** se os arquivos ainda estiverem anexados, '
                'ou carregue-os novamente para reprocessar.'
            )
        else:
            st.warning("⚠️ Aguardando upload dos arquivos obrigatórios para iniciar o processamento.")