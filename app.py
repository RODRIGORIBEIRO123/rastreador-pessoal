import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re
import io
import requests
import plotly.express as px
import json
import os
import sqlite3
import hashlib

# ==========================================
# 1. CONFIGURAÇÃO E FORMATADORES DE DADOS
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")

# Formatadores Sênior (Padrão Financeiro Brasileiro)
def f_brl(x): return f"R$ {x:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
def f_brl_4(x): return f"R$ {x:,.4f}".replace(",", "v").replace(".", ",").replace("v", ".")
def f_pct(x): return f"{x:,.2f}%".replace(",", "v").replace(".", ",").replace("v", ".")

MAPEAMENTO_TICKERS = {
    "GALG11": "GARE11", "SOMA3": "ALOS3", "ARZZ3": "ALOS3", 
    "VVAR3": "BHIA3", "VIIA3": "BHIA3", "BRML3": "ALSO3", 
    "BBRK11": "BRCR11", "HCTR11": "TRXD11", "TORD11": "TRXD11"
}

UNITS_ACOES = ['SANB11', 'TAEE11', 'KLBN11', 'BPAC11', 'ALUP11', 'ENGI11', 'BIDI11', 'CPLE11', 'SAPR11', 'RNEW11']

# Inicialização de Variáveis de Sessão
if 'df_base' not in st.session_state: st.session_state.df_base = pd.DataFrame()
if 'dados_mercado' not in st.session_state: st.session_state.dados_mercado = {}
if 'df_simul' not in st.session_state: st.session_state.df_simul = pd.DataFrame()
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'username' not in st.session_state: st.session_state.username = ""

# ==========================================
# 2. MOTOR DE BANCO DE DADOS (SQLite3)
# ==========================================
DB_FILE = "terminal_cnpi.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (username TEXT PRIMARY KEY, password TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS carteiras 
                 (username TEXT, Ativo TEXT, Quantidade REAL, Preco_Medio REAL, Data_Media TEXT)''')
    conn.commit()
    conn.close()

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def registrar_usuario(username, password):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    try:
        c.execute("INSERT INTO usuarios (username, password) VALUES (?, ?)", (username, hash_password(password)))
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        conn.close()
        return False

def autenticar_usuario(username, password):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM usuarios WHERE username=? AND password=?", (username, hash_password(password)))
    user = c.fetchone()
    conn.close()
    return user is not None

def salvar_carteira_db(username, df):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Limpa a carteira antiga do usuário para sobrescrever com a atualizada
    c.execute("DELETE FROM carteiras WHERE username=?", (username,))
    if not df.empty:
        for _, row in df.iterrows():
            c.execute("INSERT INTO carteiras (username, Ativo, Quantidade, Preco_Medio, Data_Media) VALUES (?, ?, ?, ?, ?)",
                      (username, row['Ativo'], row['Quantidade'], row['Preço Médio'], str(row['Data Média'])))
    conn.commit()
    conn.close()

def carregar_carteira_db(username):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT Ativo, Quantidade, Preco_Medio as 'Preço Médio', Data_Media as 'Data Média' FROM carteiras WHERE username=?", conn, params=(username,))
    conn.close()
    if not df.empty:
        df['Data Média'] = pd.to_datetime(df['Data Média']).dt.date
    return df

# Inicializa o DB ao carregar a página
init_db()

# ==========================================
# 3. TELA DE AUTENTICAÇÃO (GATEKEEPER)
# ==========================================
if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center;'>🔐 Terminal de Gestão Profissional</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Acesso restrito. Identifique-se para carregar seu portfólio.</p>", unsafe_allow_html=True)
    
    col_log1, col_log2, col_log3 = st.columns([1, 1, 1])
    with col_log2:
        tab_login, tab_register = st.tabs(["Acesso", "Novo Registro"])
        
        with tab_login:
            login_user = st.text_input("Usuário", key="log_user")
            login_pass = st.text_input("Senha", type="password", key="log_pass")
            if st.button("Entrar", use_container_width=True):
                if autenticar_usuario(login_user, login_pass):
                    st.session_state.logged_in = True
                    st.session_state.username = login_user
                    st.session_state.df_base = carregar_carteira_db(login_user) # Puxa do DB
                    st.rerun()
                else:
                    st.error("Credenciais inválidas.")
                    
        with tab_register:
            reg_user = st.text_input("Novo Usuário", key="reg_user")
            reg_pass = st.text_input("Nova Senha", type="password", key="reg_pass")
            if st.button("Registrar", use_container_width=True):
                if reg_user and reg_pass:
                    if registrar_usuario(reg_user, reg_pass):
                        st.success("Conta criada! Pode fazer o login na aba ao lado.")
                    else:
                        st.error("Nome de usuário já existe.")
                else:
                    st.warning("Preencha ambos os campos.")
    st.stop() # Bloqueia o carregamento do restante da página se não estiver logado

# ==========================================
# 4. APP PRINCIPAL (USUÁRIO AUTENTICADO)
# ==========================================
data_hoje = pd.Timestamp.now().strftime('%d/%m/%Y')
st.title(f"📊 Terminal de Gestão - Analista: {st.session_state.username.upper()}")

# Sistema de Persistência do Chat isolado por usuário
ARQUIVO_CHAT = f"historico_ia_{st.session_state.username}.json"
MENSAGEM_INICIAL = [{"role": "assistant", "content": f"Saudações, {st.session_state.username}. O terminal está mapeado em tempo real. Pode fazer a sua pergunta."}]

if 'historico_chat' not in st.session_state:
    if os.path.exists(ARQUIVO_CHAT):
        try:
            with open(ARQUIVO_CHAT, "r", encoding="utf-8") as f: st.session_state.historico_chat = json.load(f)
        except: st.session_state.historico_chat = MENSAGEM_INICIAL.copy()
    else:
        st.session_state.historico_chat = MENSAGEM_INICIAL.copy()

def salvar_chat():
    with open(ARQUIVO_CHAT, "w", encoding="utf-8") as f: json.dump(st.session_state.historico_chat, f, ensure_ascii=False, indent=4)

# ==========================================
# FUNÇÕES MACRO E FUNDAMENTOS (MANTIDAS)
# ==========================================
@st.cache_data(ttl=86400)
def carregar_macro():
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'], macro['IPCA'] = macro['CDI'] / 100, macro['IPCA'] / 100
        return macro
    except: return pd.DataFrame()

@st.cache_data(ttl=86400)
def obter_fundamentos_brasil():
    try:
        url = 'https://www.fundamentus.com.br/resultado.php'
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        res = requests.get(url, headers=headers, timeout=10)
        df = pd.read_html(io.StringIO(res.text), decimal=',', thousands='.')[0]
        fundamentos = {}
        for _, row in df.iterrows():
            t = str(row['Papel']).strip().upper()
            c = float(row['Cotação'])
            pl, pvp = float(row['P/L']), float(row['P/VP'])
            vpa = c / pvp if pvp > 0 else 0.0
            lpa = c / pl if pl > 0 else 0.0
            fundamentos[t] = {'vpa': vpa, 'lpa': lpa}
        return fundamentos
    except: return {}

@st.cache_data(ttl=86400)
def obter_macro_atual():
    selic_atual, ipca_12m = 10.50, 4.00
    try:
        res = requests.get("https://brasilapi.com.br/api/taxas/v1", timeout=5)
        if res.status_code == 200:
            for taxa in res.json():
                if taxa['nome'] == 'Selic': selic_atual = float(taxa['valor'])
    except: pass
    try:
        ipca_df = sgs.get({'IPCA_12M': 13522}, last=1)
        if not ipca_df.empty: ipca_12m = float(ipca_df['IPCA_12M'].iloc[-1])
    except: pass
    return selic_atual, ipca_12m

@st.cache_data(ttl=86400)
def obter_projecoes_focus():
    ano_atual = pd.Timestamp.now().year
    selic_atual, _ = obter_macro_atual()
    fallback = {f"IPCA_{ano_atual}": 3.80, f"Selic_{ano_atual}": selic_atual, f"IPCA_{ano_atual+1}": 3.70, f"Selic_{ano_atual+1}": selic_atual - 1.0, f"IPCA_{ano_atual+2}": 3.50, f"Selic_{ano_atual+2}": selic_atual - 1.5}
    try:
        url = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$top=300&$filter=Indicador%20eq%20'IPCA'%20or%20Indicador%20eq%20'Selic'&$orderby=Data%20desc&$format=json"
        res = requests.get(url, timeout=8).json()
        if 'value' in res and len(res['value']) > 0:
            df = pd.DataFrame(res['value'])
            data_recente = df['Data'].max()
            df = df[df['Data'] == data_recente]
            for ano_offset in [0, 1, 2]:
                ano_alvo = str(ano_atual + ano_offset)
                df_ano = df[df['DataReferencia'] == ano_alvo]
                if not df_ano[df_ano['Indicador'] == 'IPCA'].empty: fallback[f"IPCA_{ano_alvo}"] = float(df_ano[df_ano['Indicador'] == 'IPCA']['Mediana'].values[0])
                if not df_ano[df_ano['Indicador'] == 'Selic'].empty: fallback[f"Selic_{ano_alvo}"] = float(df_ano[df_ano['Indicador'] == 'Selic']['Mediana'].values[0])
        return fallback, ano_atual
    except: return fallback, ano_atual

def calcular_macro_acumulado(df_macro, data_inicio, data_fim=None):
    if df_macro is None or df_macro.empty or pd.isna(data_inicio): return 0.0, 0.0
    try:
        filtro = df_macro.loc[data_inicio:data_fim] if data_fim else df_macro.loc[data_inicio:]
        return ((1 + filtro['CDI'].dropna()).prod() - 1) * 100, ((1 + filtro['IPCA'].dropna()).prod() - 1) * 100
    except: return 0.0, 0.0

def calcular_meses(data_inicio):
    if pd.isna(data_inicio): return 0
    hoje = pd.Timestamp.now()
    return int((hoje.year - data_inicio.year) * 12 + (hoje.month - data_inicio.month))

def ignorar_ativo(ticker):
    if pd.isna(ticker): return True
    t = str(ticker).strip().upper()
    if not t or t == 'NAN' or t.startswith(('WIN', 'WDO', 'IND', 'DOL')): return True
    t_limpo = t[:-1] if t.endswith('F') else t
    if re.match(r'^[A-Z]{4}[A-Z]\d+', t_limpo) and not t_limpo.endswith(('11', '34', '39')) and len(t_limpo) >= 6: return True
    return False

def limpar_numero(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float, np.number)): return float(x)
    x = str(x).replace('R$', '').replace('.', '').replace(',', '.').strip()
    try: return float(x)
    except: return 0.0

def traduzir_setor(setor_en):
    mapa = {"Banks": "Bancos", "Utilities - Regulated Electric": "Energia", "Real Estate - Retail": "Shoppings/Varejo", "REIT - Retail": "Shoppings/Varejo", "Real Estate - Industrial": "Logística", "REIT - Industrial": "Logística", "REIT - Office": "Lajes Corporativas", "REIT - Diversified": "Fundo Híbrido", "Financial Data & Stock Exchanges": "Bolsa de Valores", "Insurance": "Seguradoras", "Oil & Gas Integrated": "Petróleo e Gás"}
    return mapa.get(setor_en, "Outros Setores")

def consolidar_carteira(df):
    if df.empty: return df
    df['Ativo'] = df['Ativo'].astype(str).str.strip().str.upper()
    df['Ativo'] = df['Ativo'].apply(lambda x: MAPEAMENTO_TICKERS.get(x, x))
    linhas = []
    for ativo, group in df.groupby('Ativo'):
        qtd = float(group['Quantidade'].sum())
        if qtd <= 0: continue
        valor_total = (group['Quantidade'] * group['Preço Médio']).sum()
        pm = valor_total / qtd if qtd > 0 else 0
        soma_tempo = sum((pd.Timestamp(row['Data Média']).timestamp() * row['Quantidade']) for _, row in group.iterrows() if pd.notna(row['Data Média']))
        dt_media = pd.to_datetime(soma_tempo / qtd, unit='s').date() if qtd > 0 else pd.Timestamp.now().date()
        linhas.append({"Ativo": ativo, "Quantidade": qtd, "Preço Médio": float(pm), "Data Média": dt_media})
    return pd.DataFrame(linhas)

def ler_arquivo_universal(arquivo_upload):
    texto = arquivo_upload.getvalue().decode('utf-8-sig', errors='ignore') if arquivo_upload.name.endswith('.csv') else None
    return pd.read_csv(io.StringIO(texto), sep=';' if ';' in texto.split('\n')[0] else ',') if texto else pd.read_excel(arquivo_upload)

def processar_planilha_b3(df):
    if df.empty: return pd.DataFrame()
    df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
    df['Quantidade'] = df['Quantidade'].apply(limpar_numero)
    df['Valor'] = df['Valor'].apply(limpar_numero)
    df = df.sort_values('Data do Negócio')
    posicoes = {}
    for _, row in df.iterrows():
        if pd.isna(row['Código de Negociação']): continue
        ticker = str(row['Código de Negociação']).strip().upper()
        ticker = ticker[:-1] if ticker.endswith('F') else ticker
        ticker = MAPEAMENTO_TICKERS.get(ticker, ticker)
        if ignorar_ativo(ticker): continue
        qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio']
        if pd.isna(data): data = pd.Timestamp.now()
        if ticker not in posicoes: posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'ts_medio': 0.0}
        
        if row['Tipo de Movimentação'] == 'Compra':
            q_ant, ts_ant = posicoes[ticker]['qtd'], posicoes[ticker]['ts_medio']
            ts_novo = pd.Timestamp(data).timestamp()
            posicoes[ticker]['ts_medio'] = ts_novo if q_ant == 0 else ((ts_ant * q_ant) + (ts_novo * qtd)) / (q_ant + qtd)
            posicoes[ticker]['qtd'] += qtd
            posicoes[ticker]['valor'] += valor
        elif row['Tipo de Movimentação'] == 'Venda':
            if qtd >= (posicoes[ticker]['qtd'] - 0.001): posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'ts_medio': 0.0}
            else:
                pm = posicoes[ticker]['valor'] / posicoes[ticker]['qtd'] if posicoes[ticker]['qtd'] > 0 else 0
                posicoes[ticker]['qtd'] -= qtd
                posicoes[ticker]['valor'] -= (qtd * pm)
    ativos_limpos = [{"Ativo": t, "Quantidade": d['qtd'], "Preço Médio": d['valor']/d['qtd'] if d['qtd'] > 0 else 0, "Data Média": pd.to_datetime(d['ts_medio'], unit='s').date()} for t, d in posicoes.items() if d['qtd'] > 0]
    return consolidar_carteira(pd.DataFrame(ativos_limpos))

# ==========================================
# SIDEBAR: UPLOAD & LOGOUT & BANCO DE DADOS
# ==========================================
st.sidebar.markdown(f"👤 Logado como: **{st.session_state.username}**")
if st.sidebar.button("🚪 Sair (Logout)", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.df_base = pd.DataFrame()
    st.rerun()

st.sidebar.divider()

st.sidebar.markdown("### 💾 Banco de Dados Nuvem")
if not st.session_state.df_base.empty:
    if st.sidebar.button("Salvar Estado Atual no DB", type="primary", use_container_width=True):
        salvar_carteira_db(st.session_state.username, st.session_state.df_base)
        st.sidebar.success("Carteira sincronizada com sucesso!")

st.sidebar.divider()

st.sidebar.header("1. Subir Novas Operações")
arquivo_principal = st.sidebar.file_uploader("Substituir Carteira Completa", type=["xlsx", "csv"])
arquivo_novo = st.sidebar.file_uploader("Apenas Novas Negociações B3", type=["xlsx", "csv"])
data_corte = None
if arquivo_novo: data_corte = st.sidebar.date_input("Filtrar Novas a partir de:", pd.Timestamp.now().date() - pd.Timedelta(days=15))

if st.sidebar.button("🚀 Processar Arquivos", use_container_width=True):
    base_atual = st.session_state.df_base.copy() # Base atual (pode ser a do DB)
    
    if arquivo_principal: # Se mandou principal, sobrescreve tudo
        with st.spinner("Lendo planilha principal..."):
            df_prin = ler_arquivo_universal(arquivo_principal)
            if 'Data Média' in df_prin.columns:
                df_prin['Data Média'] = pd.to_datetime(df_prin['Data Média'], errors='coerce').dt.date
                base_atual = consolidar_carteira(df_prin)
            else:
                base_atual = processar_planilha_b3(df_prin)
                
    if arquivo_novo and not base_atual.empty: # Mescla com novas negociações
        with st.spinner("Cruzando novas operações..."):
            df_novos = ler_arquivo_universal(arquivo_novo)
            df_novos['Data do Negócio'] = pd.to_datetime(df_novos['Data do Negócio'], dayfirst=True, errors='coerce')
            df_novos = df_novos[df_novos['Data do Negócio'].dt.date >= data_corte]
            
            linhas_base = []
            for _, row in base_atual.iterrows():
                linhas_base.append({"Código de Negociação": row['Ativo'], "Tipo de Movimentação": "Compra", "Data do Negócio": pd.to_datetime(row['Data Média']), "Quantidade": row['Quantidade'], "Valor": row['Quantidade'] * row['Preço Médio']})
            df_historico_falso = pd.DataFrame(linhas_base)
            df_mesclado = pd.concat([df_historico_falso, df_novos], ignore_index=True)
            base_atual = processar_planilha_b3(df_mesclado)
            
    st.session_state.df_base = base_atual
    st.sidebar.warning("Memória atualizada. Lembre-se de clicar em 'Salvar Estado Atual no DB' para tornar permanente.")
    st.rerun()

# ==========================================
# PAINEL MACRO (SEMPRE VISÍVEL)
# ==========================================
proj_focus, ano_atual = obter_projecoes_focus()
selic_hoje, ipca_12m_hoje = obter_macro_atual()

st.markdown("### 🇧🇷 Conjuntura Macroeconômica")
c_mac1, c_mac2 = st.columns([1, 2])
c_mac1.success(f"🎯 **Cenário Atual (Vigente)**\n\nSelic Atual: **{f_pct(selic_hoje)} a.a.**\n\nIPCA 12 meses: **{f_pct(ipca_12m_hoje)}**")
c_mac2.info(
    f"🔮 **Projeções do Mercado (Focus)**\n\n"
    f"**Selic:** {ano_atual}: **{f_pct(proj_focus.get(f'Selic_{ano_atual}', 0))}**  |  "
    f"{ano_atual+1}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+1}', 0))}**  |  "
    f"{ano_atual+2}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+2}', 0))}**\n\n"
    f"**IPCA:** {ano_atual}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual}', 0))}**  |  "
    f"{ano_atual+1}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+1}', 0))}**  |  "
    f"{ano_atual+2}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+2}', 0))}**"
)
st.write("---")

# ==========================================
# CONTROLE DE CARTEIRA E CONEXÃO B3
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Controle do Banco de Dados")
    c_a, c_b = st.columns([1, 2])
    with c_a:
        tdel = st.selectbox("Excluir Ativo:", [""] + sorted(st.session_state.df_base["Ativo"].tolist()))
        if st.button("Remover", use_container_width=True) and tdel != "":
            st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != tdel]
            st.rerun()
    with c_b:
        nt = st.text_input("Nova Compra Manual (Ticker)")
        ncq, ncp = st.columns(2)
        nq = ncq.number_input("Qtd", min_value=1)
        np_val = ncp.number_input("PM (R$)", min_value=0.01)
        if st.button("Adicionar à Memória", use_container_width=True) and nt != "":
            nl = pd.DataFrame([{"Ativo": nt.upper(), "Quantidade": float(nq), "Preço Médio": float(np_val), "Data Média": pd.Timestamp.now().date()}])
            st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, nl], ignore_index=True))
            st.rerun()

    df_editado = st.data_editor(st.session_state.df_base, use_container_width=True, hide_index=True, column_config={"Ativo": st.column_config.TextColumn("Ativo", disabled=False), "Quantidade": st.column_config.NumberColumn(min_value=0.0), "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0), "Data Média": st.column_config.DateColumn("Data Média", format="DD/MM/YYYY")})

    if st.button("🚀 Conectar ao Mercado Vivo", type="primary"):
        st.session_state.df_base = consolidar_carteira(df_editado) 
        df_macro, fundamentos_br = carregar_macro(), obter_fundamentos_brasil()
        progresso, total = st.progress(0), len(st.session_state.df_base)
        dados_mercado, linhas_simul_iniciais = {}, []

        for i, row in st.session_state.df_base.iterrows():
            ticker = str(row['Ativo']).strip().upper()
            data_compra = pd.to_datetime(row['Data Média']) if pd.notna(row['Data Média']) else pd.Timestamp.now()
            preco_atual, divs_total, divs_12m, lpa, vpa, setor = float(row['Preço Médio']), 0.0, 0.0, 0.0, 0.0, "Desconhecido"
            tipo_ativo = "FII" if ticker.endswith('11') and ticker not in UNITS_ACOES else "Ação"
            
            try:
                acao = yf.Ticker(f"{ticker}.SA")
                try: 
                    hist = acao.history(period="1d")
                    if not hist.empty: preco_atual = float(hist['Close'].iloc[-1])
                except: pass
                try: 
                    divs = acao.dividends
                    if not divs.empty:
                        if divs.index.tz is not None: divs.index = divs.index.tz_localize(None)
                        divs_total = float(divs[divs.index >= data_compra].sum() * row['Quantidade'])
                        divs_12m = float(divs[divs.index >= (pd.Timestamp.now() - pd.DateOffset(years=1))].sum())
                except: pass
                try: setor = traduzir_setor(acao.info.get('industry', ''))
                except: pass
            except: pass

            if ticker in fundamentos_br: vpa, lpa = fundamentos_br[ticker]['vpa'], fundamentos_br[ticker]['lpa']
            cdi, ipca = calcular_macro_acumulado(df_macro, data_compra)
            
            dados_mercado[ticker] = {"Qtd": float(row['Quantidade']), "PM": float(row['Preço Médio']), "Data": data_compra, "Preço Atual": preco_atual, "Div_Total": divs_total, "CDI": cdi, "IPCA": ipca, "Setor": setor, "Tipo": tipo_ativo}
            linhas_simul_iniciais.append({"Ativo": ticker, "Cotação Atual": preco_atual, "VPA (Contábil)": vpa, "LPA Projetado": lpa, "Div. Projetado (R$)": divs_12m})
            progresso.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dados_mercado
        st.session_state.df_simul = pd.DataFrame(linhas_simul_iniciais)
        st.success("Mercado Sincronizado!")

# ==========================================
# PAINEL DE RELATÓRIOS E DASHBOARD
# ==========================================
    if st.session_state.dados_mercado:
        linhas_perf = []
        for t, dm in st.session_state.dados_mercado.items():
            investido = dm['Qtd'] * dm['PM']
            saldo = dm['Qtd'] * dm['Preço Atual']
            resultado = saldo - investido
            var_c_div = (((saldo + dm['Div_Total']) / investido) - 1) * 100 if investido > 0 else 0
            yoc = (dm['Div_Total'] / investido) * 100 if investido > 0 else 0
            
            linhas_perf.append({
                "Ativo": t, "Tipo": dm["Tipo"], "Setor": dm["Setor"], "Qtd": int(dm['Qtd']), "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
                "Total Investido": investido, "Saldo Atual": saldo, "Resultado (R$)": resultado,
                "Data Média": dm['Data'].strftime('%d/%m/%Y'), "Total Div. (R$)": dm['Div_Total'], "DY on Cost (%)": yoc, "Evolução c/ Div (%)": var_c_div
            })
        df_perf_final = pd.DataFrame(linhas_perf)

        st.markdown("### 🏆 Visão Global do Portfólio")
        df_acoes = df_perf_final[df_perf_final['Tipo'] == 'Ação']
        df_fiis = df_perf_final[df_perf_final['Tipo'] == 'FII']
        evolucao_acoes = (df_acoes['Saldo Atual'].sum() / df_acoes['Total Investido'].sum() - 1) * 100 if df_acoes['Total Investido'].sum() > 0 else 0
        evolucao_fiis = (df_fiis['Saldo Atual'].sum() / df_fiis['Total Investido'].sum() - 1) * 100 if df_fiis['Total Investido'].sum() > 0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📈 Patrimônio (Ações)", f_brl(df_acoes['Saldo Atual'].sum()), f_pct(evolucao_acoes))
        m2.metric("🏢 Patrimônio (FIIs)", f_brl(df_fiis['Saldo Atual'].sum()), f_pct(evolucao_fiis))
        m3.metric("💸 Renda Acumulada (Ações)", f_brl(df_acoes['Total Div. (R$)'].sum()))
        m4.metric("💸 Renda Acumulada (FIIs)", f_brl(df_fiis['Total Div. (R$)'].sum()))

        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Visão Geral", "💰 Valuation (Bazin & Graham)", "⚖️ Pesos e Setores", "📈 Gráficos", "💸 Proventos Mensais"])
        
        with tab1:
            styled_perf = df_perf_final.drop(columns=['Tipo', 'Setor']).style.format({
                "Preço Médio": f_brl, "Preço Atual": f_brl, "Total Investido": f_brl, "Saldo Atual": f_brl, 
                "Resultado (R$)": f_brl, "Total Div. (R$)": f_brl, "DY on Cost (%)": f_pct, "Evolução c/ Div (%)": f_pct
            })
            st.dataframe(styled_perf, use_container_width=True, hide_index=True)

        with tab2:
            st.markdown("#### 1. Método Bazin (Foco em Renda)")
            yield_desejado = st.number_input("Taxa de Risco Exigida (%):", value=6.0, step=0.5) / 100.0
            df_bazin_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)"]].copy()
            df_bazin_editado = st.data_editor(df_bazin_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"], key="edit_bazin")
            st.session_state.df_simul["Div. Projetado (R$)"] = df_bazin_editado.iloc[:, 2]
            linhas_bazin = []
            for _, row in df_bazin_editado.iterrows():
                bazin = (row.iloc[2] / yield_desejado) if (row.iloc[2] > 0 and yield_desejado > 0) else 0.0
                margem_b = (((bazin / row['Cotação Atual']) - 1) * 100) if (bazin > 0 and row['Cotação Atual'] > 0) else 0.0
                linhas_bazin.append({"Ativo": row['Ativo'], "Preço Teto (Bazin)": bazin, "Margem Segurança (%)": margem_b})
            st.dataframe(pd.DataFrame(linhas_bazin).style.format({"Preço Teto (Bazin)": f_brl, "Margem Segurança (%)": f_pct}), use_container_width=True, hide_index=True)
            
            st.markdown("#### 2. Método Graham (Foco em Valor)")
            df_graham_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "VPA (Contábil)", "LPA Projetado"]].copy()
            df_graham_editado = st.data_editor(df_graham_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"], key="edit_graham")
            st.session_state.df_simul["VPA (Contábil)"] = df_graham_editado["VPA (Contábil)"]
            st.session_state.df_simul["LPA Projetado"] = df_graham_editado["LPA Projetado"]
            linhas_graham = []
            for _, row in df_graham_editado.iterrows():
                graham = (22.5 * row['LPA Projetado'] * row['VPA (Contábil)']) ** 0.5 if (row['LPA Projetado'] > 0 and row['VPA (Contábil)'] > 0) else 0.0
                margem_g = (((graham / row['Cotação Atual']) - 1) * 100) if (graham > 0 and row['Cotação Atual'] > 0) else 0.0
                linhas_graham.append({"Ativo": row['Ativo'], "Preço Justo (Graham)": graham, "Margem Segurança (%)": margem_g})
            st.dataframe(pd.DataFrame(linhas_graham).style.format({"Preço Justo (Graham)": f_brl, "Margem Segurança (%)": f_pct}), use_container_width=True, hide_index=True)

        with tab3:
            c_g1, c_g2, c_g3 = st.columns(3)
            c_g1.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Tipo', hole=0.4, title="Por Tipo"), use_container_width=True)
            c_g2.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Ativo', title="Por Ativo"), use_container_width=True)
            c_g3.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Setor', title="Por Setor"), use_container_width=True)

        with tab4:
            todos_ativos = df_perf_final['Ativo'].tolist()
            c_sel, c_ind = st.columns([2, 1])
            with c_sel: ativos_selecionados = st.multiselect("Ativos:", todos_ativos, default=todos_ativos[:6], key="ms_g")
            with c_ind: ind_selecionados = st.multiselect("Indicadores:", ["Evolução c/ Div (%)", "DY on Cost (%)"], default=["Evolução c/ Div (%)"], key="ind_g")
            if ativos_selecionados and ind_selecionados:
                df_grafico = df_perf_final[df_perf_final['Ativo'].isin(ativos_selecionados)].melt(id_vars=["Ativo"], value_vars=ind_selecionados, var_name="Indicador", value_name="Rentabilidade")
                fig1 = px.bar(df_grafico, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group")
                fig1.update_layout(yaxis_ticksuffix=" %")
                st.plotly_chart(fig1, use_container_width=True)

        with tab5:
            st.markdown("### 💸 Filtro e Histórico de Proventos")
            c_f1, c_f2, c_btn = st.columns([2, 2, 2])
            meses_map = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
            mes_atual = pd.Timestamp.now().month
            ano_atual_data = pd.Timestamp.now().year
            
            mes_selecionado = c_f1.selectbox("Mês de Referência:", options=list(meses_map.keys()), format_func=lambda x: meses_map[x], index=mes_atual-1)
            ano_selecionado = c_f2.selectbox("Ano de Referência:", options=[ano_atual_data, ano_atual_data-1, ano_atual_data-2, ano_atual_data-3, ano_atual_data-4])
            
            if 'divs_calculados' not in st.session_state: st.session_state.divs_calculados = None
            if c_btn.button("🔄 Processar Dividendos do Período", use_container_width=True) or st.session_state.divs_calculados is None:
                st.session_state.divs_mes = mes_selecionado
                st.session_state.divs_ano = ano_selecionado
                with st.spinner(f"Varrendo histórico de {meses_map[mes_selecionado]}/{ano_selecionado}..."):
                    linhas_div = []
                    for t, dm in st.session_state.dados_mercado.items():
                        try:
                            divs = yf.Ticker(f"{t}.SA").dividends
                            if not divs.empty:
                                if divs.index.tz is not None: divs.index = divs.index.tz_localize(None)
                                divs_mes = divs[(divs.index.month == st.session_state.divs_mes) & (divs.index.year == st.session_state.divs_ano)].sum()
                                if divs_mes > 0:
                                    yoc = ((divs_mes * dm['Qtd']) / (dm['Qtd'] * dm['PM'])) * 100 if dm['PM'] > 0 else 0
                                    dy_spot = (divs_mes / dm['Preço Atual']) * 100 if dm['Preço Atual'] > 0 else 0
                                    linhas_div.append({"Ativo": t, "Provento Unit. (R$)": float(divs_mes), "Qtd Detida": int(dm['Qtd']), "Total Recebido (R$)": float(divs_mes * dm['Qtd']), "Yield on Cost (%)": float(yoc), "DY Atual (%)": float(dy_spot)})
                        except: continue
                    st.session_state.divs_calculados = pd.DataFrame(linhas_div)
            
            df_divs = st.session_state.divs_calculados
            if df_divs is not None and not df_divs.empty:
                st.dataframe(df_divs.style.format({"Provento Unit. (R$)": f_brl_4, "Total Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct, "DY Atual (%)": f_pct}), use_container_width=True, hide_index=True)
                st.success(f"**Total Recebido em {meses_map[st.session_state.divs_mes]}/{st.session_state.divs_ano}:** {f_brl(df_divs['Total Recebido (R$)'].sum())}")
            elif df_divs is not None and df_divs.empty:
                st.info(f"Sem proventos para {meses_map[st.session_state.divs_mes]}/{st.session_state.divs_ano}.")

# ==========================================
# 8. COMITÊ DE IA E CHAT (SEMPRE VISÍVEL)
# ==========================================
st.write("---")
c_chat1, c_chat2 = st.columns([10, 2])
with c_chat1: st.markdown("### 💬 Comitê de Alocação IA")
with c_chat2:
    if st.button("🗑️ Limpar Chat"):
        st.session_state.historico_chat = MENSAGEM_INICIAL.copy()
        if os.path.exists(ARQUIVO_CHAT): os.remove(ARQUIVO_CHAT)
        st.rerun()

api_key = ""
try: api_key = st.secrets["GEMINI_API_KEY"]
except: api_key = st.text_input("Chave API do Google Gemini:", type="password")

for msg in st.session_state.historico_chat:
    with st.chat_message(msg["role"]): st.write(msg["content"])
    
if prompt := st.chat_input("Ex: Com a Selic atual, devo aportar em Renda Fixa ou FIIs?"):
    st.session_state.historico_chat.append({"role": "user", "content": prompt})
    salvar_chat()
    with st.chat_message("user"): st.write(prompt)
    
    ctx_carteira = "Nenhuma carteira carregada."
    if not st.session_state.df_base.empty and 'df_perf_final' in locals():
        ctx_carteira = df_perf_final[['Ativo', 'Qtd', 'Preço Médio', 'Preço Atual', 'Saldo Atual', 'Evolução c/ Div (%)']].to_csv(index=False, sep='|')
        
    ctx_macro = (f"Selic Hoje: {f_pct(selic_hoje)} | IPCA 12m: {f_pct(ipca_12m_hoje)}\n"
                 f"Proj Selic: {ano_atual}: {f_pct(proj_focus.get(f'Selic_{ano_atual}'))} | {ano_atual+1}: {f_pct(proj_focus.get(f'Selic_{ano_atual+1}'))}\n"
                 f"Proj IPCA: {ano_atual}: {f_pct(proj_focus.get(f'IPCA_{ano_atual}'))} | {ano_atual+1}: {f_pct(proj_focus.get(f'IPCA_{ano_atual+1}'))}")
    
    sys_prompt = f"Gestora Sênior CNPI. Responda tecnicamente cruzando a base. Dados: {ctx_carteira} Macro: {ctx_macro}."
    
    if api_key:
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            modelos_para_tentar = ['gemini-3.5-flash', 'gemini-3.1-flash-lite', 'gemini-2.5-flash']
            resposta_sucesso = False
            for nome_modelo in modelos_para_tentar:
                try:
                    model = genai.GenerativeModel(nome_modelo)
                    resposta = model.generate_content([sys_prompt, prompt]).text
                    resposta_sucesso = True
                    break 
                except: continue
            if not resposta_sucesso: resposta = "⚠️ Falha de comunicação com a IA."
        except Exception as e: resposta = f"⚠️ Erro estrutural: {str(e)}"
    else: resposta = "⚠️ Operação bloqueada. Insira a chave."
    
    st.session_state.historico_chat.append({"role": "assistant", "content": resposta})
    salvar_chat()
    st.rerun()
