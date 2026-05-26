import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import os
import math
import altair as alt
from datetime import datetime, timedelta

# Configuração da Página - Mantendo o layout wide
st.set_page_config(page_title="Gestor HUUFMA PRO", layout="wide", page_icon="🏥")

# --- ESTILIZAÇÃO CSS CUSTOMIZADA (UI/UX) ---
st.markdown("""
    <style>
    /* Estilização Geral do Fundo e Títulos */
    .main .block-container { padding-top: 2rem; }
    h1 { color: #1E3A8A; font-weight: 700; margin-bottom: 0.5rem; }
    h2, h3 { color: #2C3E50; font-weight: 600; }
    
    /* Customização dos Cards de Métrica */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 1rem 1.25rem;
        border-radius: 0.75rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
    }
    
    /* Botão Principal Destacado */
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
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES ORIGINAIS DE APOIO ---
def clean(t):
    if pd.isna(t): return ""
    nfkd = unicodedata.normalize('NFKD', str(t))
    t = "".join([c for c in nfkd if not unicodedata.combining(c)])
    t = re.sub(r'[^a-zA-Z0-9\s]', ' ', t)
    return " ".join(t.lower().split())

def clean_key(v):
    return re.sub(r'[^0-9]', '', str(v)).lstrip('0')

def p_num(v):
    try:
        if pd.isna(v) or str(v).strip() == "": return 0.0
        l = re.sub(r'[^0-9,.]', '', str(v))
        if "," in l and "." in l: l = l.replace(".", "").replace(",", ".")
        elif "," in l: l = l.replace(",", ".")
        return float(l) if l else 0.0
    except: return 0.0

def find_col(df, terms, forbidden=[]):
    for col in df.columns:
        if any(t in clean(col) for t in terms) and not any(f in clean(col) for f in forbidden): return col
    return None

# --- GESTÃO DESCENTRALIZADA DE CATEGORIAS POR FARMÁCIA ---
PASTA_RAIZ_CAT = "Categorias_HUUFMA"

def obter_pasta_farmacia(cod_farmacia):
    caminho_subpasta = os.path.join(PASTA_RAIZ_CAT, str(cod_farmacia))
    if not os.path.exists(caminho_subpasta):
        os.makedirs(caminho_subpasta)
    return caminho_subpasta

def carregar_mapa_categorias(cod_farmacia):
    mapa = {}
    pasta_farmacia = obter_pasta_farmacia(cod_farmacia)
    if os.path.exists(pasta_farmacia):
        for arquivo in os.listdir(pasta_farmacia):
            caminho = os.path.join(pasta_farmacia, arquivo)
            try:
                df_c = pd.read_csv(caminho, sep=None, engine='python') if arquivo.endswith('.csv') else pd.read_excel(caminho)
                col_c = find_col(df_c, ['cod', 'material', 'ca3'])
                if col_c:
                    for cod in df_c[col_c].apply(clean_key):
                        if cod: mapa[cod] = arquivo.split('.')[0].upper()
            except: continue
    return mapa

# --- PAINEL LATERAL (SIDEBAR) CONTROLES GERAIS ---
with st.sidebar:
    st.markdown("### 🏥 Unidade & Parâmetros")
    farmacias_opcoes = {
        "Farmácia UMI (Cód. 13)": "13",
        "Farmácia Dutra (Cód. 31)": "31",
        "Farmácia Centro Cirúrgico (Cód. 7)": "7",
        "Farmácia Oftalmologia (Cód. 39)": "39",
        "Farmácia UTI (Cód. 34)": "34"
    }
    farmacia_selecionada = st.selectbox("Unidade Hospitalar de Trabalho:", list(farmacias_opcoes.keys()))
    cod_farmacia_alvo = farmacias_opcoes[farmacia_selecionada]
    
    st.write("---")
    st.markdown("### ⚙️ Parâmetros do Pedido")
    dias_pedido = st.number_input("Sugestão de reabastecimento (dias):", value=15, min_value=1)
    
    hoje = datetime.now()
    data_padrao_inicio = hoje - timedelta(days=5)
    data_inicio = st.date_input("Início do Período de Consumo:", value=data_padrao_inicio)
    data_fim = st.date_input("Fim do Período de Consumo:", value=hoje)

pasta_atual_categorias = obter_pasta_farmacia(cod_farmacia_alvo)

# --- CORPO PRINCIPAL ---
st.title("🏥 Gestor de Pedidos Logístico Avançado")
st.markdown(f"**Unidade Ativa:** `{farmacia_selecionada}` | **Período de Análise:** `{data_inicio.strftime('%d/%m/%Y')}` até `{data_fim.strftime('%d/%m/%Y')}`")
st.write("")

tab1, tab2 = st.tabs(["⚡ Processar Pedido com IA Logística", "📂 Gestão de Categorias da Unidade"])

with tab2:
    st.subheader("Gerenciador de Categorias Vinculadas")
    st.info("O mapeamento de categorias permite classificar os materiais dinamicamente (Ex: Medicamentos vs MMH).")
    
    uploaded_file = st.file_uploader("Adicionar nova Planilha de Categoria (CSV/XLSX)", type=["csv", "xlsx"])
    if uploaded_file:
        caminho_salvar = os.path.join(pasta_atual_categorias, uploaded_file.name)
        with open(caminho_salvar, "wb") as f: 
            f.write(uploaded_file.getbuffer())
        st.success(f"Categoria '{uploaded_file.name}' vinculada com sucesso à {farmacia_selecionada}!")
    
    st.write("")
    st.markdown("##### Arquivos Ativos nesta Unidade")
    arquivos_na_pasta = os.listdir(pasta_atual_categorias)
    
    if arquivos_na_pasta:
        for arq in arquivos_na_pasta:
            col_arq, col_btn = st.columns([4, 1])
            col_arq.markdown(f"🔹 **{arq}**")
            if col_btn.button("Remover", key=f"del_{cod_farmacia_alvo}_{arq}", use_container_width=True):
                os.remove(os.path.join(pasta_atual_categorias, arq))
                st.warning(f"Arquivo {arq} removido.")
                st.rerun()
    else:
        st.info("Nenhuma categoria específica cadastrada. Itens padrão agrupados em 'OUTROS'.")

with tab1:
    with st.container(border=True):
        st.markdown("##### 📥 Upload das Fontes de Dados Obrigatórias")
        col1, col2 = st.columns(2)
        file_mov_alvo = col1.file_uploader("1. Movimento da Farmácia Alvo (.csv)", type=["csv"], help="Arquivo contendo o histórico de RM da unidade.")
        file_est_geral = col2.file_uploader("2. Estoque Geral de todos os Almoxarifados (.csv)", type=["csv"], help="Espelho de estoque atual consolidado.")
        
        st.write("")
        files_mov_parceiras = st.file_uploader("3. Movimentos das Outras Farmácias (Opcional - Múltiplos .csv)", type=["csv"], accept_multiple_files=True, help="Utilizado para cruzar e rastrear estoque parado em outras unidades.")

    st.write("")
    
    # Execução do Processamento
    if file_mov_alvo and file_est_geral:
        st.write("")
        if st.button("🚀 GERAR PEDIDO COM INTELIGÊNCIA LOGÍSTICA", use_container_width=True):
            with st.spinner("Cruzando estoques institucionais e calculando CMD..."):
                try:
                    # 1. Leitura dos Arquivos Principais
                    mov = pd.read_csv(file_mov_alvo, sep=None, engine='python', encoding='latin1', index_col=False)
                    est_geral = pd.read_csv(file_est_geral, sep=None, engine='python', encoding='latin1', index_col=False)

                    # Busca de Colunas Dinâmicas
                    c_mov_cod = find_col(mov, ['material', 'cod', 'ca3'])
                    c_mov_qtd = find_col(mov, ['quant'])
                    c_mov_tipo = find_col(mov, ['tipo'])
                    c_mov_data = find_col(mov, ['data', 'ger'])
                    c_mov_almox = find_col(mov, ['almox'])
                    
                    c_est_cod = find_col(est_geral, ['cod', 'ca3', 'ident'], forbidden=['material', 'prod'])
                    c_est_qtd = find_col(est_geral, ['qtde disp', 'disponivel'])
                    c_est_prod = find_col(est_geral, ['material', 'produto', 'descri'])
                    c_est_almox = find_col(est_geral, ['almox'])
                    c_est_min = find_col(est_geral, ['qtde estq min', 'estoque minimo', 'minimo'])

                    # Padronização de chaves do Estoque Geral
                    est_geral['key'] = est_geral[c_est_cod].apply(clean_key)
                    est_geral['almox_limpo'] = est_geral[c_est_almox].apply(clean_key)
                    est_geral['saldo_num'] = est_geral[c_est_qtd].apply(p_num)
                    est_geral['min_num'] = est_geral[c_est_min].apply(p_num) if c_est_min else 0.0

                    # Separar estoques estratégicos do Estoque Geral
                    est_farmacia_alvo = est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo].groupby('key')['saldo_num'].sum().to_dict()
                    
                    # --- MAPEAR DIFERENTES ALMOXARIFADOS CENTRAIS ---
                    est_central_6 = est_geral[est_geral['almox_limpo'] == '6'].groupby('key')['saldo_num'].sum().to_dict()
                    est_central_mmh = est_geral[est_geral['almox_limpo'].isin(['1', '43'])].groupby('key')['saldo_num'].sum().to_dict()
                    
                    # Mapeamento do Estoque Mínimo cadastrado especificamente para a farmácia alvo
                    est_min_alvo = est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo].groupby('key')['min_num'].sum().to_dict()
                    
                    # Mapeamento de quais farmácias têm saldo de cada item
                    cod_parceiras = [c for c in ['7', '13', '31', '34', '39'] if c != cod_farmacia_alvo]
                    est_outras_farmacias = est_geral[est_geral['almox_limpo'].isin(cod_parceiras)]
                    
                    dict_saldos_parceiras = {}
                    for _, row in est_outras_farmacias.dropna(subset=['key']).iterrows():
                        k = row['key']
                        alm = row['almox_limpo']
                        sal = row['saldo_num']
                        if k not in dict_saldos_parceiras: dict_saldos_parceiras[k] = {}
                        dict_saldos_parceiras[k][alm] = dict_saldos_parceiras[k].get(alm, 0) + sal

                    # 2. Processar consumo das OUTRAS farmácias para achar "Estoque Parado"
                    consumo_outras_total = {}
                    if files_mov_parceiras:
                        for f_parc in files_mov_parceiras:
                            try:
                                df_p = pd.read_csv(f_parc, sep=None, engine='python', encoding='latin1', index_col=False)
                                c_p_cod = find_col(df_p, ['material', 'cod', 'ca3'])
                                c_p_qtd = find_col(df_p, ['quant'])
                                c_p_tipo = find_col(df_p, ['tipo'])
                                c_p_almox = find_col(df_p, ['almox'])
                                c_p_data = find_col(df_p, ['data', 'ger'])
                                
                                df_p['dt_formatada'] = pd.to_datetime(df_p[c_p_data], dayfirst=True, errors='coerce')
                                df_p_filt = df_p[
                                    (df_p['dt_formatada'].dt.date >= data_inicio) & 
                                    (df_p['dt_formatada'].dt.date <= data_fim) & 
                                    (df_p[c_p_tipo].astype(str).str.upper() == 'RM')
                                ].copy()
                                
                                if c_p_cod and c_p_qtd and c_p_almox:
                                    df_p_filt['key'] = df_p_filt[c_p_cod].apply(clean_key)
                                    df_p_filt['almox_limpo'] = df_p_filt[c_p_almox].apply(clean_key)
                                    for _, r_p in df_p_filt.iterrows():
                                        k_p = r_p['key']
                                        alm_p = r_p['almox_limpo']
                                        qtd_p = p_num(r_p[c_p_qtd])
                                        if k_p not in consumo_outras_total: consumo_outras_total[k_p] = {}
                                        consumo_outras_total[k_p][alm_p] = consumo_outras_total[k_p].get(alm_p, 0) + qtd_p
                            except: continue

                    # 3. Calcular Consumo e Sugestão da Farmácia Alvo
                    mov['dt_formatada'] = pd.to_datetime(mov[c_mov_data], dayfirst=True, errors='coerce')
                    mov_filtrado = mov[
                        (mov['dt_formatada'].dt.date >= data_inicio) & 
                        (mov['dt_formatada'].dt.date <= data_fim) & 
                        (mov[c_mov_tipo].astype(str).str.upper() == 'RM')
                    ].copy()
                    
                    dias_considerados = max((data_fim - data_inicio).days + 1, 1)
                    consumo = mov_filtrado.groupby(c_mov_cod)[c_mov_qtd].apply(lambda x: sum(p_num(v) for v in x)).reset_index()
                    
                    def aplicar_teto_logistico(qtd_bruta):
                        cmd_bruto = qtd_bruta / dias_considerados
                        return float(math.ceil(cmd_bruto)) if cmd_bruto > 0 else 0.0

                    consumo['cmd'] = consumo[c_mov_qtd].apply(aplicar_teto_logistico)
                    consumo['key'] = consumo[c_mov_cod].apply(clean_key)

                    # Montagem do Painel Final de Decisões
                    mapa_produtos = est_geral.drop_duplicates(subset=['key']).set_index('key')[c_est_prod].to_dict()
                    todos_codigos = sorted(list(set(est_geral[est_geral['almox_limpo'] == cod_farmacia_alvo]['key'].unique()) | set(consumo['key'])))
                    
                    final = pd.DataFrame({'Código': todos_codigos})
                    final['Material'] = final['Código'].map(mapa_produtos).fillna('PRODUTO SEM DESCRIÇÃO NO ESTOQUE')
                    final['Estoque Atual Unidade'] = final['Código'].map(est_farmacia_alvo).fillna(0)
                    final['Consumo Médio Diário'] = final['Código'].map(consumo.set_index('key')['cmd']).fillna(0)
                    final['Estoque Mínimo'] = final['Código'].map(est_min_alvo).fillna(0)
                    
                    def calcular_sugestao(row):
                        cmd = row['Consumo Médio Diário']
                        est_atual = row['Estoque Atual Unidade']
                        est_minimo = row['Estoque Mínimo']
                        meta_consumo = cmd * dias_pedido
                        meta_final = max(meta_consumo, est_minimo)
                        return max(0, round(meta_final - est_atual))

                    final['Sugestão de Pedido'] = final.apply(calcular_sugestao, axis=1)
                    
                    mapa_cat = carregar_mapa_categorias(cod_farmacia_alvo)
                    final['Categoria'] = final['Código'].map(mapa_cat).fillna('OUTROS')
                    
                    def definir_saldo_central(row):
                        cod = row['Código']
                        cat = row['Categoria']
                        if str(cat).upper() == "MMH": return est_central_mmh.get(cod, 0)
                        return est_central_6.get(cod, 0)

                    final['Saldo Almox. Central'] = final.apply(definir_saldo_central, axis=1)
                    
                    def inteligencia_remanejamento(row):
                        cod = row['Código']
                        v_sug_val = row['Sugestão de Pedido']
                        saldo_central = row['Saldo Almox. Central']
                        cmd = row['Consumo Médio Diário']
                        est_un = row['Estoque Atual Unidade']
                        est_minimo = row['Estoque Mínimo']
                        cat = row['Categoria']
                        
                        nome_almox = "Almox 1/43 (MMH)" if str(cat).upper() == "MMH" else "Almox 6 (Med)"
                        
                        if cmd > 0 and est_un > (cmd * 60):
                            return "Estoque Excessivo", "Estoque cobre mais de 60 dias. Avaliar devolução."
                        if cmd == 0 and est_minimo <= 0 and est_un > 0:
                            return "Estoque Parado", "Sem consumo registrado no período."
                        if v_sug_val <= 0:
                            return "Estoque Suficiente", "Estoque dentro da cobertura ideal."
                        if saldo_central >= v_sug_val:
                            return "Solicitar CAF", f"Solicitar {int(v_sug_val)} un. ao {nome_almox}."
                        if 0 < saldo_central < v_sug_val:
                            return "Estoque Crítico CAF", f"Pegar {int(saldo_central)} un. na Central e remanejar o restante."
                        
                        saldos_parceiras = dict_saldos_parceiras.get(cod, {})
                        consumos_parceiras = consumo_outras_total.get(cod, {})
                        
                        farmacias_com_estoque_parado = []
                        for farm_id, saldo_f in saldos_parceiras.items():
                            if saldo_f > 0 and consumos_parceiras.get(farm_id, 0) == 0:
                                farmacias_com_estoque_parado.append(f"Cód {farm_id} ({int(saldo_f)} un.)")
                        
                        if farmacias_com_estoque_parado:
                            locais_remanejamento = ", ".join(farmacias_com_estoque_parado)
                            return "Remanejar", f"Central Zerada! Transferir de: {locais_remanejamento}."
                        
                        return "Desabastecimento Crítico", "Sem saldo na central e sem estoque parado em parceiras."

                    resultados = final.apply(inteligencia_remanejamento, axis=1)
                    final['ALERTAS'] = [r[0] for r in resultados]
                    final['Ação Logística Sugerida'] = [r[1] for r in resultados]
                    
                    # --- DASHBOARD VISUAL DE MÉTRICAS (PRO) ---
                    st.write("---")
                    st.subheader(f"📊 Resumo Executivo Logístico")
                    
                    df_desabast = final[final['ALERTAS'] == "Desabastecimento Crítico"].sort_values(by='Material')
                    df_remanej = final[final['ALERTAS'] == "Remanejar"].sort_values(by='Material')
                    df_caf_disp = final[final['ALERTAS'].isin(["Solicitar CAF", "Estoque Crítico CAF"])].sort_values(by='Material')
                    df_excesso_parados = final[final['ALERTAS'].isin(["Estoque Excessivo", "Estoque Parado"])].sort_values(by='Material')
                    
                    def gerar_excel_memorizado(df_dados):
                        out = io.BytesIO()
                        with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                            df_dados.to_excel(wr, index=False, sheet_name="Filtro")
                        return out.getvalue()

                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        st.metric("🚨 Desabastecimento Crítico", f"{len(df_desabast)} itens")
                        st.download_button(label="📥 Baixar Lista", data=gerar_excel_memorizado(df_desabast), file_name=f"Desabastecimento_{cod_farmacia_alvo}.xlsx", key="btn_rel_c1", use_container_width=True)
                    with c2:
                        st.metric("🔄 Oportunidades Remanejar", f"{len(df_remanej)} itens")
                        st.download_button(label="📥 Baixar Lista", data=gerar_excel_memorizado(df_remanej), file_name=f"Remanejamento_{cod_farmacia_alvo}.xlsx", key="btn_rel_c2", use_container_width=True)
                    with c3:
                        st.metric("📦 Disponíveis na CAF", f"{len(df_caf_disp)} itens")
                        st.download_button(label="📥 Baixar Lista", data=gerar_excel_memorizado(df_caf_disp), file_name=f"Disponiveis_CAF_{cod_farmacia_alvo}.xlsx", key="btn_rel_c3", use_container_width=True)
                    with c4:
                        st.metric("⚠️ Excesso / Parados", f"{len(df_excesso_parados)} itens")
                        st.download_button(label="📥 Baixar Lista", data=gerar_excel_memorizado(df_excesso_parados), file_name=f"Risco_Vencimento_{cod_farmacia_alvo}.xlsx", key="btn_rel_c4", use_container_width=True)
                    
                    # --- SEÇÃO DE GRÁFICOS INTERATIVOS OTIMIZADA ---
                    st.write("")
                    g1, g2 = st.columns(2)
                    
                    with g1:
                        st.markdown("**Distribuição por Status Logístico**")
                        df_tree = final.groupby(['ALERTAS', 'Categoria'])['Código'].count().reset_index()
                        df_tree.columns = ['Status', 'Categoria', 'Quantidade']
                        
                        grafico_tree = alt.Chart(df_tree).mark_bar().encode(
                            x=alt.X('Quantidade:Q', title="Quantidade de Itens"),
                            y=alt.Y('Status:N', title=None, sort='-x'),
                            color=alt.Color('Status:N', scale=alt.Scale(
                                domain=["Estoque Suficiente", "Solicitar CAF", "Estoque Crítico CAF", "Remanejar", "Desabastecimento Crítico", "Estoque Excessivo", "Estoque Parado"],
                                range=["#A2E8A2", "#A6C8FF", "#FFC499", "#FFEAA6", "#FFA6A6", "#B2EBF2", "#CFD8DC"]
                            ), legend=None),
                            tooltip=['Status', 'Categoria', 'Quantidade']
                        ).properties(height=240)
                        st.altair_chart(grafico_tree, use_container_width=True)
                    
                    with g2:
                        st.markdown("**Necessidade de Reordenamento por Categoria**")
                        df_sug_cat = final[final['Categoria'] != 'OUTROS'].groupby('Categoria')['Sugestão de Pedido'].count().reset_index()
                        df_sug_cat.columns = ['Categoria', 'Quantidade de Itens']
                        
                        grafico_barras = alt.Chart(df_sug_cat).mark_bar(color="#1E3A8A").encode(
                            x=alt.X('Quantidade de Itens:Q', title="Qtd Itens Solicitados"),
                            y=alt.Y('Categoria:N', title=None, sort='-x'),
                            tooltip=["Categoria", "Quantidade de Itens"]
                        ).properties(height=240)
                        st.altair_chart(grafico_barras, use_container_width=True)
                    
                    # --- DATAFRAME COM FORMATAÇÃO AVANÇADA (BADGES NA TELA) ---
                    st.write("---")
                    st.markdown("#### 📋 Painel de Análise Inteligente Integrado")
                    
                    # Preparando dados para exibição na tela
                    final_exibicao = final[['Código', 'Material', 'Categoria', 'Estoque Atual Unidade', 'Sugestão de Pedido', 'ALERTAS', 'Ação Logística Sugerida']].copy()
                    
                    # Configurando colunas interativas e badges estilizadas usando a api st.column_config
                    st.dataframe(
                        final_exibicao.sort_values(by='Material'),
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Código": st.column_config.TextColumn("Código ID", width="small"),
                            "Material": st.column_config.TextColumn("Descrição do Material", width="large"),
                            "Estoque Atual Unidade": st.column_config.NumberColumn("Estoque Físico", format="%d"),
                            "Sugestão de Pedido": st.column_config.NumberColumn("Sugestão Pedido", format="%d"),
                            "ALERTAS": st.column_config.SelectboxColumn(
                                "Status Alerta",
                                width="medium",
                                options=["Estoque Suficiente", "Solicitar CAF", "Estoque Crítico CAF", "Remanejar", "Desabastecimento Crítico", "Estoque Excessivo", "Estoque Parado"]
                            )
                        }
                    )
                    
                    # --- GERAÇÃO COMPLETA DO RELATÓRIO EXCEL MULTI-ABA EM BUFFER ---
                    final_salvar = final.rename(columns={
                        'Sugestão de Pedido': f'Sugestão de Pedido ({dias_pedido} dias)',
                        'Saldo Almox. Central': 'Saldo Almox. Central (1/43 ou 6)'
                    })
                    
                    ordem_colunas = [
                        'Código', 'Material', 'Categoria', 'Estoque Atual Unidade', 
                        'Consumo Médio Diário', 'Estoque Mínimo', f'Sugestão de Pedido ({dias_pedido} dias)',
                        'Saldo Almox. Central (1/43 ou 6)', 'ALERTAS', 'Ação Logística Sugerida'
                    ]
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        for cat, df_cat in final_salvar.groupby('Categoria'):
                            df_exportar = df_cat[ordem_colunas].sort_values(by='Material')
                            df_exportar.to_excel(writer, sheet_name=str(cat)[:31], index=False)
                            
                            workbook = writer.book
                            worksheet = writer.sheets[str(cat)[:31]]
                            
                            # Formatação de larguras automáticas
                            for idx, col in enumerate(df_exportar.columns):
                                series = df_exportar[col]
                                max_len = max(series.astype(str).map(len).max(), len(str(col))) + 3
                                worksheet.set_column(idx, idx, max_len)
                            
                            # Regras de Cores Condicionais no Excel Gerado
                            f_verde = workbook.add_format({'bg_color': '#E2EFDA', 'font_color': '#375623'}) 
                            f_azul = workbook.add_format({'bg_color': '#DDEBF7', 'font_color': '#1F4E78'}) 
                            f_laranja = workbook.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911'}) 
                            f_amarelo = workbook.add_format({'bg_color': '#FFF2CC', 'font_color': '#7F6000'}) 
                            f_vermelho = workbook.add_format({'bg_color': '#F8CBAD', 'font_color': '#C00000'})
                            f_ciano = workbook.add_format({'bg_color': '#E5F1F4', 'font_color': '#006666'}) 
                            f_cinza = workbook.add_format({'bg_color': '#F2F2F2', 'font_color': '#595959'}) 
                            
                            total_l = len(df_exportar)
                            worksheet.conditional_format(1, 8, total_l, 8, {'type': 'cell', 'criteria': 'equal to', 'value': '"Estoque Suficiente"', 'format': f_verde})
                            worksheet.conditional_format(1, 8, total_l, 8, {'type': 'cell', 'criteria': 'equal to', 'value': '"Solicitar CAF"', 'format': f_azul})
                            worksheet.conditional_format(1, 8, total_l, 8, {'type': 'cell', 'criteria': 'equal to', 'value': '"Estoque Crítico CAF"', 'format': f_laranja})
                            worksheet.conditional_format(1, 8, total_l, 8, {'type': 'cell', 'criteria': 'equal to', 'value': '"Remanejar"', 'format': f_amarelo})
                            worksheet.conditional_format(1, 8, total_l, 8, {'type': 'cell', 'criteria': 'equal to', 'value': '"Desabastecimento Crítico"', 'format': f_vermelho})
                            worksheet.conditional_format(1, 8, total_l, 8, {'type': 'cell', 'criteria': 'equal to', 'value': '"Estoque Excessivo"', 'format': f_ciano})
                            worksheet.conditional_format(1, 8, total_l, 8, {'type': 'cell', 'criteria': 'equal to', 'value': '"Estoque Parado"', 'format': f_cinza})

                    st.write("")
                    st.download_button(
                        label="📥 BAIXAR MAPA DE PEDIDO COMPLETO (TODAS AS ABAS)",
                        data=buffer.getvalue(),
                        file_name=f"Pedido_Inteligente_{cod_farmacia_alvo}_{datetime.now().strftime('%d%m%y')}.xlsx",
                        use_container_width=True
                    )
                    
                except Exception as e: 
                    st.error(f"Ocorreu um erro ao cruzar as colunas das tabelas anexadas. Certifique-se de que os cabeçalhos estão legíveis. Erro técnico: {e}")
    else:
        st.warning("⚠️ Aguardando o upload dos arquivos obrigatórios (Movimento e Estoque Geral) para iniciar o cruzamento de dados.")