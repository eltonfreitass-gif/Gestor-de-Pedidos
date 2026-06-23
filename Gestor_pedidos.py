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

# Faixas de alerta de validade (dias até o vencimento)
VALIDADE_CRITICO_DIAS  = 30   # 🔴 vence em até 30 dias
VALIDADE_ATENCAO_DIAS  = 90   # 🟡 vence em até 90 dias

# Chave session_state para o link SharePoint de validades
SS_SHAREPOINT_URL = 'sharepoint_validades_url'

# Farmácias satélites (códigos de almoxarifado)
CODIGOS_FARMACIAS = ['7', '13', '31', '34', '39']


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
                         ocultar_colunas: list | None = None) -> None:
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


def exportar_excel_padronizado(df_dados: pd.DataFrame, nome_aba: str = "Dados") -> bytes:
    buf = io.BytesIO()
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
                            fmt_h, fmt_b, fmt_t, fmt_status_cache, fmt_critico)
    return buf.getvalue()


def exportar_excel_multi_aba(df_total: pd.DataFrame, ordem_cols: list,
                              col_categoria: str, col_alerta: str,
                              larguras: dict, excluir_acoes: list = None,
                              ocultar_colunas: list | None = None,
                              ajustar_altura_linhas: bool = True) -> bytes:
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

            nome_aba = str(cat)[:31]
            df_exp.to_excel(writer, sheet_name=nome_aba, index=False)
            ws = writer.sheets[nome_aba]

            _aplicar_estilo_aba(ws, wb, df_exp, col_alerta, larguras,
                                fmt_h, fmt_b, fmt_t, fmt_status_cache, fmt_critico,
                                ajustar_altura_linhas=ajustar_altura_linhas,
                                ocultar_colunas=ocultar_colunas)
    return buf.getvalue()


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
    """
    if df_val is None or df_val.empty:
        return df_val

    df = df_val.copy()

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


# =============================================================================
# CONSOLIDAÇÃO MULTI-FARMÁCIA — FUNÇÕES
# =============================================================================

def calcular_status_farmacia(df_estoque: pd.DataFrame, df_mov: pd.DataFrame,
                              cod_farm: str, mapa_cat: dict,
                              data_ini: date, data_fim: date,
                              dias_pedido: int) -> pd.DataFrame:
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
    if not all([c_mc, c_mq, c_mt, c_md]):
        return pd.DataFrame()

    mov = df_mov.copy()
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

    # Parecer simplificado para consolidação
    def _parecer(row):
        if row['CMD'] == 0 and row['Saldo Atual'] == 0:
            return 'Sem Consumo'
        if row['CMD'] > 0 and row['Saldo Atual'] > row['CMD'] * 60:
            return 'Estoque Excessivo'
        if row['Necessidade'] <= 0:
            return 'Estoque Suficiente'
        if row['Saldo Central'] >= row['Necessidade']:
            return 'Solicitar'
        if row['Saldo Central'] > 0:
            return 'Estoque Crítico CAF'
        return 'Desabastecimento Crítico'

    final['Parecer'] = final.apply(_parecer, axis=1)
    return final

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
    except Exception:
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

tab1, tab2, tab3, tab4 = st.tabs([
    "⚡ Processar Pedido com IA Logística",
    "⏰ Controle de Validade",
    "🏥 Consolidação Multi-Farmácia",
    "🗂️ Gestão de Categorias de Insumos",
])




# =============================================================================
# TAB 2 — CONTROLE DE VALIDADE
# =============================================================================
with tab2:
    st.subheader("⏰ Controle de Validade — FEFO")

    # ── CONFIGURAÇÃO DO SHAREPOINT ────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("##### 🔗 Fonte de Dados de Validade")

        # Status da última carga automática
        ultima_carga = st.session_state.get('sharepoint_ultima_carga', None)
        if ultima_carga:
            st.success(f"✅ SharePoint carregado com sucesso em: **{ultima_carga}**")
        else:
            st.info(
                "🔗 A planilha de validades será carregada automaticamente do SharePoint "
                "(conta farmaciahuufma@outlook.com). Clique no botão abaixo para carregar."
            )

        c_sp1, c_sp2 = st.columns([2, 1])
        file_val_manual = c_sp1.file_uploader(
            "📎 Carregar manualmente (usado se o SharePoint falhar):",
            type=['xlsx', 'xls'], key='upload_validades'
        )
        carregar_val = c_sp2.button(
            "🔄 Carregar / Atualizar Validades", use_container_width=True
        )

    # ── CARREGAMENTO ──────────────────────────────────────────────────────────
    if carregar_val or 'df_validades_mescladas' not in st.session_state:
        df_sp_raw = None
        erro_sp   = ''

        # 1. Tentar SharePoint autenticado (conta convidada)
        with st.spinner("🌐 Conectando ao SharePoint..."):
            df_sp_raw, erro_sp = carregar_validades_sharepoint()

        # 2. Fallback para upload manual
        if df_sp_raw is None:
            if erro_sp:
                st.warning(
                    f"⚠️ **Não foi possível acessar o SharePoint:** {erro_sp}\n\n"
                    "Carregue o arquivo manualmente pelo campo acima para continuar."
                )
            if file_val_manual:
                file_val_manual.seek(0)
                df_sp_raw = carregar_planilha_validades_excel(io.BytesIO(file_val_manual.read()))
                st.info("📎 Usando arquivo carregado manualmente.")
        
        if df_sp_raw is None:
            st.error(
                "❌ Não foi possível carregar as validades automaticamente e nenhum "
                "arquivo foi carregado manualmente. Verifique a conexão ou carregue o arquivo."
            )
        else:
            df_sp_norm = normalizar_planilha_validades(df_sp_raw)
            if df_sp_norm.empty:
                st.error(
                    "❌ A planilha de validades não pôde ser interpretada. "
                    "Verifique se tem colunas de Código MV e Validade identificáveis."
                )
            else:
                # 3. Extrair validades do AGHU (se estoque já foi processado)
                df_aghu_val = pd.DataFrame()
                if 'est_geral_raw' in st.session_state:
                    df_aghu_val = extrair_validades_aghu(
                        st.session_state['est_geral_raw']
                    )

                df_val_final = mesclar_validades(df_aghu_val, df_sp_norm)
                mapa_cat_val = obter_mapa_categorias()
                est_geral_para_saldo = st.session_state.get('est_geral_raw')
                st.session_state['df_validades_mescladas'] = preparar_painel_validades(
                    df_val_final, mapa_cat_val, est_geral_para_saldo
                )
                n_aghu = len(df_aghu_val)
                n_sp   = len(df_sp_norm)
                n_painel = len(st.session_state['df_validades_mescladas'])
                tem_estoque_aghu = est_geral_para_saldo is not None and not est_geral_para_saldo.empty
                complemento_filtro = (
                    f" Após filtro por **Saldo AGHU > 0**, permanecem **{n_painel}** registros no painel."
                    if tem_estoque_aghu else
                    f" **{n_painel}** registros foram mantidos no painel provisório, sem filtro por AGHU."
                )
                st.success(
                    f"✅ Validades carregadas: **{n_aghu}** registros do AGHU + "
                    f"**{n_sp}** da planilha da equipe → "
                    f"**{len(df_val_final)}** registros candidatos após mesclagem."
                    + complemento_filtro
                )

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
        ops_sit  = ["TODAS"] + sorted(df_val['Situação'].unique().tolist())
        ops_cat  = ["TODAS"] + sorted(df_val['Categoria'].unique().tolist())
        fil_farm = fv1.selectbox("Farmácia:", ops_farm, key="val_farm")
        fil_sit  = fv2.selectbox("Situação:", ops_sit,  key="val_sit")
        fil_cat  = fv3.selectbox("Categoria:", ops_cat, key="val_cat")

        df_vf = df_val.copy()
        if fil_farm != "TODAS": df_vf = df_vf[df_vf['Nome Farmácia'] == fil_farm]
        if fil_sit  != "TODAS": df_vf = df_vf[df_vf['Situação'] == fil_sit]
        if fil_cat  != "TODAS": df_vf = df_vf[df_vf['Categoria'] == fil_cat]

        cols_exib = ['🚦', 'Situação', 'key', 'Material', 'Lote',
                     'Validade Fmt', 'Saldo AGHU', 'Saldo Planilha Validade',
                     'Nome Farmácia', 'Categoria', 'Fonte']
        cols_ok   = [c for c in cols_exib if c in df_vf.columns]

        st.dataframe(
            df_vf[cols_ok].reset_index(drop=True),
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
        df_criticos = df_val[df_val['Situação'].str.contains('VENCIDO|Crítico|Atenção', na=False)]
        with dv1:
            if not df_criticos.empty:
                st.download_button(
                    "📥 Exportar Vencidos + Até 90 dias (.xlsx)",
                    data=exportar_excel_padronizado(
                        df_criticos[cols_ok].reset_index(drop=True), "Ate_90_Dias"
                    ),
                    file_name=f"Validades_Ate_90_Dias_{datetime.now().strftime('%d%m%y')}.xlsx",
                    use_container_width=True,
                )
        with dv2:
            st.download_button(
                "📥 Exportar Painel Completo de Validades (.xlsx)",
                data=exportar_excel_padronizado(
                    df_vf[cols_ok].reset_index(drop=True), "Validades"
                ),
                file_name=f"Validades_Completo_{datetime.now().strftime('%d%m%y')}.xlsx",
                use_container_width=True,
            )
    else:
        st.info(
            "⬆️ Configure a URL do SharePoint ou carregue o arquivo de validades acima "
            "e clique em **Carregar / Atualizar Validades**.\n\n"
            "💡 Dica: processe o pedido na **Aba 1** antes — isso permite que o app "
            "já extraia as validades presentes no arquivo de estoque do AGHU automaticamente."
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
                    mapa_cat, data_ini, data_fim, dias_ped
                )
                if not df_res.empty:
                    resultados.append(df_res)

            if resultados:
                df_consolidado = pd.concat(resultados, ignore_index=True)
                st.session_state['df_consolidado'] = df_consolidado
                prog_c.progress(100, text="✅ Consolidação concluída!")
            else:
                st.error("❌ Nenhum dado pôde ser consolidado. Verifique os arquivos.")

        if 'df_consolidado' in st.session_state:
            df_cons = st.session_state['df_consolidado'].copy()

            # ── KPIs POR FARMÁCIA ─────────────────────────────────────────────
            st.write("---")
            st.markdown("#### 📊 Painel Comparativo por Farmácia")

            painel_farms = []
            for cod, nome in DIC_NOMES_FARMACIAS.items():
                dff = df_cons[df_cons['Cód. Farmácia'] == cod]
                if dff.empty:
                    continue
                painel_farms.append({
                    'Farmácia': nome,
                    '📦 Itens': len(dff),
                    '🔴 Rupturas': (dff['Parecer'] == 'Desabastecimento Crítico').sum(),
                    '🔵 Solicitar': (dff['Parecer'] == 'Solicitar').sum(),
                    '🟠 CAF Crítico': (dff['Parecer'] == 'Estoque Crítico CAF').sum(),
                    '🟡 Remanejar': (dff['Parecer'] == 'Remanejar').sum(),
                    '🩵 Excessos': (dff['Parecer'] == 'Estoque Excessivo').sum(),
                    '⚫ Sem Consumo': (dff['Parecer'] == 'Sem Consumo').sum(),
                })

            if painel_farms:
                df_painel = pd.DataFrame(painel_farms)
                st.dataframe(df_painel, use_container_width=True, hide_index=True)

            # ── RUPTURAS SIMULTÂNEAS ──────────────────────────────────────────
            st.write("---")
            st.markdown("#### 🚨 Itens em Desabastecimento Crítico em Múltiplas Farmácias")
            st.caption("Sinal de desabastecimento sistêmico — não localizado numa única unidade.")

            df_rupt = df_cons[df_cons['Parecer'] == 'Desabastecimento Crítico']
            conta_rupt = df_rupt.groupby('Código MV').agg(
                Material=('Material', 'first'),
                Categoria=('Categoria', 'first'),
                N_Farmacias=('Farmácia', 'nunique'),
                Farmacias=('Farmácia', lambda x: ' | '.join(sorted(x)))
            ).reset_index().sort_values('N_Farmacias', ascending=False)

            df_simult = conta_rupt[conta_rupt['N_Farmacias'] > 1]
            if df_simult.empty:
                st.success("✅ Nenhum item em ruptura simultânea em mais de uma farmácia.")
            else:
                st.error(f"⚠️ {len(df_simult)} item(ns) em ruptura em 2+ farmácias simultaneamente.")
                st.dataframe(df_simult.rename(columns={
                    'N_Farmacias': 'Nº Farmácias', 'Farmacias': 'Farmácias Afetadas'
                }), use_container_width=True, hide_index=True)

            # ── OPORTUNIDADES DE REMANEJAMENTO ────────────────────────────────
            st.write("---")
            st.markdown("#### 🔄 Oportunidades de Remanejamento Entre Farmácias")
            st.caption("Item com excesso em uma farmácia e necessidade em outra.")

            df_exc = df_cons[df_cons['Parecer'].isin(['Estoque Excessivo', 'Estoque Parado'])][
                ['Código MV', 'Material', 'Categoria', 'Farmácia', 'Saldo Atual', 'CMD']
            ].copy()
            df_nec = df_cons[df_cons['Parecer'] == 'Desabastecimento Crítico'][
                ['Código MV', 'Farmácia', 'Necessidade']
            ].copy()

            if df_exc.empty or df_nec.empty:
                st.info("Nenhuma oportunidade de remanejamento identificada no momento.")
            else:
                df_opor = df_exc.merge(
                    df_nec, on='Código MV', suffixes=('_origem', '_destino')
                )
                df_opor = df_opor[
                    df_opor['Farmácia_origem'] != df_opor['Farmácia_destino']
                ].rename(columns={
                    'Farmácia_origem': 'Transferir DE',
                    'Farmácia_destino': 'Transferir PARA',
                    'Saldo Atual': 'Saldo Origem',
                    'Necessidade': 'Qtd Necessária',
                })[['Código MV', 'Material', 'Categoria',
                    'Transferir DE', 'Saldo Origem', 'Transferir PARA', 'Qtd Necessária']]

                if df_opor.empty:
                    st.info("Nenhuma oportunidade de remanejamento identificada.")
                else:
                    st.success(f"🔄 {len(df_opor)} oportunidade(s) de remanejamento identificada(s).")
                    st.dataframe(df_opor, use_container_width=True, hide_index=True)


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
                    "Carregue o Controle de Validade na Aba 2 para ativar a análise de "
                    "remanejamento preventivo por vencimento."
                )
            else:
                df_val_rem = st.session_state['df_validades_mescladas'].copy()
                # Reaplica o saldo do AGHU, quando disponível, para garantir que o alerta seja operacional.
                if 'est_geral_raw' in st.session_state:
                    df_val_rem = aplicar_saldos_validades(df_val_rem, st.session_state['est_geral_raw'])

                df_rem_val = calcular_remanejamento_preventivo_validade(df_val_rem, df_cons)

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
                    st.dataframe(
                        df_rem_val,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            'Prioridade': st.column_config.TextColumn('Prioridade', width='medium'),
                            'Código MV': st.column_config.TextColumn('Código MV', width='small'),
                            'Material': st.column_config.TextColumn('Material', width='large'),
                            'Transferir DE': st.column_config.TextColumn('Transferir DE', width='medium'),
                            'Transferir PARA': st.column_config.TextColumn('Transferir PARA', width='medium'),
                            'Dias até Vencer': st.column_config.NumberColumn('Dias até vencer', format='%d'),
                            'Saldo AGHU': st.column_config.NumberColumn('Saldo origem AGHU', format='%d'),
                            'CMD Origem': st.column_config.NumberColumn('CMD origem', format='%d'),
                            'Qtd em Risco na Origem': st.column_config.NumberColumn('Qtd em risco', format='%d'),
                            'CMD Destino': st.column_config.NumberColumn('CMD destino', format='%d'),
                            'Saldo Destino': st.column_config.NumberColumn('Saldo destino', format='%d'),
                            'Necessidade Destino': st.column_config.NumberColumn('Necessidade destino', format='%d'),
                            'Consumo Possível no Destino até Validade': st.column_config.NumberColumn('Consumo possível até validade', format='%d'),
                            'Qtd Sugerida Remanejar': st.column_config.NumberColumn('Qtd sugerida', format='%d'),
                            'Justificativa': st.column_config.TextColumn('Justificativa', width='large'),
                        }
                    )
                    st.download_button(
                        "📥 Exportar remanejamento preventivo por validade (.xlsx)",
                        data=exportar_excel_padronizado(df_rem_val, "Remanejamento_Validade"),
                        file_name=f"Remanejamento_Validade_{datetime.now().strftime('%d%m%y')}.xlsx",
                        use_container_width=True,
                    )

            # ── VALIDADES CRÍTICAS NA CONSOLIDAÇÃO ───────────────────────────
            if 'df_validades_mescladas' in st.session_state:
                st.write("---")
                st.markdown("#### ⏰ Validades Críticas por Farmácia")
                df_vc = st.session_state['df_validades_mescladas']
                df_vc_crit = df_vc[df_vc['Situação'].str.contains('VENCIDO|Crítico', na=False)]
                if not df_vc_crit.empty:
                    vc_farm = df_vc_crit.groupby('Nome Farmácia')['key'].count().reset_index()
                    vc_farm.columns = ['Farmácia', 'Itens com Validade Crítica']
                    st.dataframe(vc_farm, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Nenhum item com validade crítica nas farmácias.")

            # ── DOWNLOAD CONSOLIDAÇÃO ─────────────────────────────────────────
            st.write("---")
            st.download_button(
                "📥 Exportar Consolidação Completa (.xlsx)",
                data=exportar_excel_multi_aba(
                    df_cons,
                    ['Código MV', 'Material', 'Categoria', 'Farmácia',
                     'Saldo Atual', 'CMD', 'Estoque Mínimo', 'Cobertura (dias)',
                     'Necessidade', 'Saldo Central', 'Parecer'],
                    col_categoria='Farmácia',
                    col_alerta='Parecer',
                    larguras={'Código MV': 12, 'Material': 45, 'Categoria': 16,
                              'Farmácia': 28, 'Saldo Atual': 14, 'CMD': 12,
                              'Estoque Mínimo': 14, 'Cobertura (dias)': 14,
                              'Necessidade': 14, 'Saldo Central': 18, 'Parecer': 28},
                ),
                file_name=f"Consolidacao_{datetime.now().strftime('%d%m%y')}.xlsx",
                use_container_width=True,
            )


# =============================================================================
# TAB 4 — GESTÃO DE CATEGORIAS
# =============================================================================
with tab4:
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
                df_filtrado_cat["Código"].astype(str).str.contains(filtro_termo_cat, case=False, na=False, regex=False) |
                df_filtrado_cat["Material"].astype(str).str.contains(filtro_termo_cat, case=False, na=False, regex=False)
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
    else:
        st.warning("🔍 Utilize os filtros acima para visualizar e editar os registros.")

    # Botão exportar mapa completo — sempre disponível, fora do if/else
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
                file_mov_alvo.seek(0)
                file_est_geral.seek(0)
                mov       = ler_csv_cached(file_mov_alvo.read(), file_mov_alvo.name)
                est_geral = ler_csv_cached(file_est_geral.read(), file_est_geral.name)

                # Salvar raws para uso nas abas Validade e Consolidação
                st.session_state['est_geral_raw']  = est_geral
                st.session_state['mov_alvo_raw']   = mov

                # Se as validades já foram carregadas antes do Estoque Geral,
                # recalcula os saldos de validade e aplica o filtro operacional: Saldo AGHU > 0.
                if 'df_validades_mescladas' in st.session_state:
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
                
                st.session_state['cod_farmacia_alvo'] = cod_farmacia_alvo
                st.session_state['nome_farmacia_alvo'] = DIC_NOMES_FARMACIAS.get(cod_farmacia_alvo, f"Almoxarifado (Cód. {cod_farmacia_alvo})")

                progress.progress(20, text="🏗️ Processando estoque geral...")

                est_geral = est_geral.copy()
                est_geral['key']         = est_geral[c_est_cod].apply(clean_key)
                est_geral['almox_limpo'] = est_geral[c_est_almox].apply(clean_key)
                est_geral['saldo_num']   = p_num_series(est_geral[c_est_qtd])
                est_geral['min_num']     = p_num_series(est_geral[c_est_min]) if c_est_min else 0.0

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
                                    movs_parceiras_raw[cod_tmp[0]] = df_tmp
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
                'Estoque Mínimo', 'Saldo Atual Satélite', 'Cobertura (dias)',
                'CMD Últ. 3 dias', 'Consumo Médio Diário', COL_SUG,
                'Tendência', 'Δ% Tendência', '⏰ Validade',
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
                st.download_button(
                    "📥 BAIXAR RELATÓRIO COMPLETO — ABA ÚNICA (.XLSX)",
                    data=exportar_excel_padronizado(df_export_geral, "Painel Geral"),
                    file_name=f"Painel_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
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
                                    f"Confirmar se vence em {int(r['Dias até Vencer'])} dia(s) "
                                    f"e não pedir até zerar o estoque"
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

                st.download_button(
                    "📥 BAIXAR PEDIDO - CLASSIFICADO POR CATEGORIA (.XLSX)",
                    data=exportar_excel_multi_aba(
                        df_pedido_abas,
                        ordem_abas_final,
                        col_categoria='Categoria',
                        col_alerta='Status para cor',
                        larguras=larguras_pedido_abas,
                        excluir_acoes=["Avaliar se é necessário inativar o item na farmácia."],
                        ocultar_colunas=['Status para cor'],
                        ajustar_altura_linhas=False,
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