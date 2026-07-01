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
# 1. CONFIGURAÇÃO E FORMATADORES
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")

def f_brl(x): return f"R$ {float(x):,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
def f_brl_4(x): return f"R$ {float(x):,.4f}".replace(",", "v").replace(".", ",").replace("v", ".")
def f_pct(x): return f"{float(x):,.2f}%".replace(",", "v").replace(".", ",").replace("v", ".")

MAPEAMENTO_TICKERS = {"GALG11": "GARE11", "SOMA3": "ALOS3", "ARZZ3": "ALOS3", "VVAR3": "BHIA3", "VIIA3": "BHIA3", "BRML3": "ALSO3", "BBRK11": "BRCR11", "HCTR11": "TRXD11", "TORD11": "TRXD11"}
UNITS_ACOES = ['SANB11', 'TAEE11', 'KLBN11', 'BPAC11', 'ALUP11', 'ENGI11', 'BIDI11', 'CPLE11', 'SAPR11', 'RNEW11']

if 'df_base' not in st.session_state: st.session_state.df_base = pd.DataFrame()
if 'dados_mercado' not in st.session_state: st.session_state.dados_mercado = {}
if 'df_simul' not in st.session_state: st.session_state.df_simul = pd.DataFrame()
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'username' not in st.session_state: st.session_state.username = ""

# ==========================================
# 2. MOTOR DE BANCO DE DADOS (SQLite)
# ==========================================
DB_FILE = "terminal_cnpi.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (username TEXT PRIMARY KEY, password TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS carteiras (username TEXT, Ativo TEXT, Quantidade REAL, Preco_Medio REAL, Data_Media TEXT)''')
    conn.commit()
    conn.close()

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()

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
    if not df.empty: df['Data Média'] = pd.to_datetime(df['Data Média']).dt.date
    return df

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
                    st.session_state.df_base = carregar_carteira_db(login_user)
                    st.rerun()
                else: st.error("Credenciais inválidas.")
        with tab_register:
            reg_user = st.text_input("Novo Usuário", key="reg_user")
            reg_pass = st.text_input("Nova Senha", type="password", key="reg_pass")
            if st.button("Registrar", use_container_width=True):
                if reg_user and reg_pass:
                    if registrar_usuario(reg_user, reg_pass): st.success("Conta criada! Pode fazer o login.")
                    else: st.error("Nome de usuário já existe.")
                else: st.warning("Preencha ambos os campos.")
    st.stop()

# ==========================================
# 4. APP PRINCIPAL: FUNÇÕES DE DADOS E IA
# ==========================================
st.title(f"📊 Terminal de Gestão - Analista: {st.session_state.username.upper()} ({pd.Timestamp.now().strftime('%d/%m/%Y')})")

ARQUIVO_CHAT = f"historico_ia_{st.session_state.username}.json"
MENSAGEM_INICIAL = [{"role": "assistant", "content": f"Saudações, {st.session_state.username}. O terminal está mapeado em tempo real."}]

if 'historico_chat' not in st.session_state:
    if os.path.exists(ARQUIVO_CHAT):
        try:
            with open(ARQUIVO_CHAT, "r", encoding="utf-8") as f: st.session_state.historico_chat = json.load(f)
        except: st.session_state.historico_chat = MENSAGEM_INICIAL.copy()
    else: st.session_state.historico_chat = MENSAGEM_INICIAL.copy()

def salvar_chat():
    with open(ARQUIVO_CHAT, "w", encoding="utf-8") as f: json.dump(st.session_state.historico_chat, f, ensure_ascii=False, indent=4)

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
        df = pd.read_html(io.StringIO(requests.get(url, headers=headers, timeout=10).text), decimal=',', thousands='.')[0]
        fundamentos = {}
        for _, row in df.iterrows():
            t, c, pl, pvp = str(row['Papel']).strip().upper(), float(row['Cotação']), float(row['P/L']), float(row['P/VP'])
            fundamentos[t] = {'vpa': c/pvp if pvp>0 else 0.0, 'lpa': c/pl if pl>0 else 0.0}
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
    fallback = {f"IPCA_{ano_atual}": 3.80, f"Selic_{ano_atual}": selic_atual, f"IPCA_{ano_atual+1}": 3.70, f"Selic_{ano_atual+1}": selic_atual-1.0, f"IPCA_{ano_atual+2}": 3.50, f"Selic_{ano_atual+2}": selic_atual-1.5}
    try:
        url = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$top=300&$filter=Indicador%20eq%20'IPCA'%20or%20Indicador%20eq%20'Selic'&$orderby=Data%20desc&$format=json"
        res = requests.get(url, timeout=8).json()
        if 'value' in res and len(res['value']) > 0:
            df = pd.DataFrame(res['value'])
            df = df[df['Data'] == df['Data'].max()]
            for ano_offset in [0, 1, 2]:
                ano_alvo = str(ano_atual + ano_offset)
                df_ano = df[df['DataReferencia'] == ano_alvo]
                if not df_ano[df_ano['Indicador'] == 'IPCA'].empty: fallback[f"IPCA_{ano_alvo}"] = float(df_ano[df_ano['Indicador'] == 'IPCA']['Mediana'].values[0])
                if not df_ano[df_ano['Indicador'] == 'Selic'].empty: fallback[f"Selic_{ano_alvo}"] = float(df_ano[df_ano['Indicador'] == 'Selic']['Mediana'].values[0])
    except: pass
    return fallback, ano_atual

def calcular_macro_acumulado(df_macro, data_inicio):
    if df_macro is None or df_macro.empty or pd.isna(data_inicio): return 0.0, 0.0
    try:
        filtro = df_macro.loc[data_inicio:]
        return ((1 + filtro['CDI'].dropna()).prod() - 1) * 100, ((1 + filtro['IPCA'].dropna()).prod() - 1) * 100
    except: return 0.0, 0.0

def limpar_numero(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float, np.number)): return float(x)
    try: return float(str(x).replace('R$', '').replace('.', '').replace(',', '.').strip())
    except: return 0.0

def traduzir_setor(setor_en):
    return {"Banks": "Bancos", "Utilities - Regulated Electric": "Energia", "Real Estate - Retail": "Shoppings/Varejo", "REIT - Retail": "Shoppings/Varejo", "Real Estate - Industrial": "Logística", "REIT - Industrial": "Logística", "REIT - Office": "Lajes Corporativas", "REIT - Diversified": "Fundo Híbrido", "Financial Data & Stock Exchanges": "Bolsa de Valores", "Insurance": "Seguradoras", "Oil & Gas Integrated": "Petróleo e Gás"}.get(setor_en, "Outros Setores")

def consolidar_carteira(df):
    if df.empty: return df
    df['Ativo'] = df['Ativo'].astype(str).str.strip().str.upper()
    df['Ativo'] = df['Ativo'].apply(lambda x: MAPEAMENTO_TICKERS.get(x, x))
    linhas = []
    for ativo, group in df.groupby('Ativo'):
        qtd = float(group['Quantidade'].sum())
        if qtd <= 0: continue
        pm = (group['Quantidade'] * group['Preço Médio']).sum() / qtd
        soma_tempo = sum((pd.Timestamp(row['Data Média']).timestamp() * row['Quantidade']) for _, row in group.iterrows() if pd.notna(row['Data Média']))
        linhas.append({"Ativo": ativo, "Quantidade": qtd, "Preço Médio": float(pm), "Data Média": pd.to_datetime(soma_tempo/qtd, unit='s').date() if qtd>0 else pd.Timestamp.now().date()})
    return pd.DataFrame(linhas)

def processar_planilha_b3(df):
    if df.empty: return pd.DataFrame()
    df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
    df['Quantidade'], df['Valor'] = df['Quantidade'].apply(limpar_numero), df['Valor'].apply(limpar_numero)
    df = df.sort_values('Data do Negócio')
    posicoes = {}
    for _, row in df.iterrows():
        if pd.isna(row['Código de Negociação']): continue
        ticker = str(row['Código de Negociação']).strip().upper()
        ticker = MAPEAMENTO_TICKERS.get(ticker[:-1] if ticker.endswith('F') else ticker, ticker[:-1] if ticker.endswith('F') else ticker)
        if re.match(r'^[A-Z]{4}[A-Z]\d+', ticker) and not ticker.endswith(('11','34','39')) and len(ticker)>=6: continue
        qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio'] if pd.notna(row['Data do Negócio']) else pd.Timestamp.now()
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
    ativos = [{"Ativo": t, "Quantidade": d['qtd'], "Preço Médio": d['valor']/d['qtd'] if d['qtd']>0 else 0, "Data Média": pd.to_datetime(d['ts_medio'], unit='s').date()} for t, d in posicoes.items() if d['qtd']>0]
    return consolidar_carteira(pd.DataFrame(ativos))

def to_excel(df, sheet_name='Sheet1'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# ==========================================
# 5. SIDEBAR: UPLOAD, LOGIN E DB
# ==========================================
st.sidebar.markdown(f"👤 **{st.session_state.username.upper()}**")
if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.df_base = pd.DataFrame()
    st.rerun()

st.sidebar.divider()
st.sidebar.markdown("### 💾 Banco de Dados")
if not st.session_state.df_base.empty:
    if st.sidebar.button("Salvar Estado Atual no DB", type="primary", use_container_width=True):
        salvar_carteira_db(st.session_state.username, st.session_state.df_base)
        st.sidebar.success("Sincronizado!")

st.sidebar.divider()
st.sidebar.header("1. Upload de Arquivos")
arquivo_principal = st.sidebar.file_uploader("Substituir Base Completa", type=["xlsx", "csv"])
arquivo_novo = st.sidebar.file_uploader("Apenas Novas Operações", type=["xlsx", "csv"])
data_corte = st.sidebar.date_input("Filtrar a partir de:", pd.Timestamp.now().date() - pd.Timedelta(days=15)) if arquivo_novo else None

if st.sidebar.button("🚀 Processar", use_container_width=True):
    base_atual = st.session_state.df_base.copy()
    if arquivo_principal:
        txt = arquivo_principal.getvalue().decode('utf-8-sig', errors='ignore') if arquivo_principal.name.endswith('.csv') else None
        df_p = pd.read_csv(io.StringIO(txt), sep=';' if txt and ';' in txt else ',') if txt else pd.read_excel(arquivo_principal)
        
        # VALIDAÇÃO DE COLUNAS RESTAURADA PARA EVITAR KEYERROR
        if 'Data Média' in df_p.columns:
            base_atual = consolidar_carteira(df_p)
        elif 'Data do Negócio' in df_p.columns:
            base_atual = processar_planilha_b3(df_p)
        else:
            st.sidebar.error("Formato de planilha inválido. Use o backup do sistema ou o extrato padrão da B3.")
            st.stop()
            
    if arquivo_novo and not base_atual.empty:
        txt_n = arquivo_novo.getvalue().decode('utf-8-sig', errors='ignore') if arquivo_novo.name.endswith('.csv') else None
        df_n = pd.read_csv(io.StringIO(txt_n), sep=';' if ';' in txt_n else ',') if txt_n else pd.read_excel(arquivo_novo)
        df_n['Data do Negócio'] = pd.to_datetime(df_n['Data do Negócio'], dayfirst=True, errors='coerce')
        df_n = df_n[df_n['Data do Negócio'].dt.date >= data_corte]
        linhas_b = [{"Código de Negociação": r['Ativo'], "Tipo de Movimentação": "Compra", "Data do Negócio": pd.to_datetime(r['Data Média']), "Quantidade": r['Quantidade'], "Valor": r['Quantidade']*r['Preço Médio']} for _, r in base_atual.iterrows()]
        base_atual = processar_planilha_b3(pd.concat([pd.DataFrame(linhas_b), df_n], ignore_index=True))
    st.session_state.df_base = base_atual
    st.sidebar.warning("Memória atualizada. Salve no DB para manter.")
    st.rerun()

# ==========================================
# 6. PAINEL MACRO (SEMPRE VISÍVEL)
# ==========================================
proj_focus, ano_atual = obter_projecoes_focus()
selic_hoje, ipca_12m_hoje = obter_macro_atual()

st.markdown("### 🇧🇷 Conjuntura Macroeconômica")
c_m1, c_m2 = st.columns([1, 2])
c_m1.success(f"🎯 **Cenário Atual (Vigente)**\n\nSelic Atual: **{f_pct(selic_hoje)} a.a.**\n\nIPCA 12 meses: **{f_pct(ipca_12m_hoje)}**")
c_m2.info(
    f"🔮 **Projeções do Mercado (Focus)**\n\n"
    f"**Selic:** {ano_atual}: **{f_pct(proj_focus.get(f'Selic_{ano_atual}', 0))}** |  {ano_atual+1}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+1}', 0))}** |  {ano_atual+2}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+2}', 0))}**\n\n"
    f"**IPCA:** {ano_atual}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual}', 0))}** |  {ano_atual+1}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+1}', 0))}** |  {ano_atual+2}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+2}', 0))}**"
)
st.write("---")

# ==========================================
# 7. CONTROLE DA CARTEIRA E CONEXÃO
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Controle Operacional")
    ca, cb, cc = st.columns([1, 1, 1])
    with ca:
        tdel = st.selectbox("Excluir Ativo:", [""] + sorted(st.session_state.df_base["Ativo"].tolist()))
        if st.button("Remover") and tdel:
            st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != tdel]
            st.rerun()
    with cb:
        nt = st.text_input("Nova Compra (Ticker)")
        cq, cp = st.columns(2)
        nq = cq.number_input("Qtd", min_value=1)
        np_v = cp.number_input("PM (R$)", min_value=0.01)
        if st.button("Adicionar") and nt:
            nl = pd.DataFrame([{"Ativo": nt.upper(), "Quantidade": float(nq), "Preço Médio": float(np_v), "Data Média": pd.Timestamp.now().date()}])
            st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, nl], ignore_index=True))
            st.rerun()
    with cc:
        st.download_button("📥 Baixar Carteira Atual (CSV)", data=st.session_state.df_base.to_csv(index=False, sep=';', encoding='utf-8-sig'), file_name="Carteira_Backup.csv", use_container_width=True)

    df_editado = st.data_editor(st.session_state.df_base, use_container_width=True, hide_index=True)

    if st.button("🚀 Conectar ao Mercado Vivo", type="primary"):
        st.session_state.df_base = consolidar_carteira(df_editado) 
        df_macro, fundamentos_br = carregar_macro(), obter_fundamentos_brasil()
        progresso, total = st.progress(0), len(st.session_state.df_base)
        dados_mercado, linhas_simul_iniciais = {}, []

        for i, row in st.session_state.df_base.iterrows():
            ticker, data_compra = str(row['Ativo']).strip().upper(), pd.to_datetime(row['Data Média']) if pd.notna(row['Data Média']) else pd.Timestamp.now()
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
        st.success("Sincronizado!")

# ==========================================
# 8. DASHBOARD E RELATÓRIOS
# ==========================================
    if st.session_state.dados_mercado:
        linhas_perf = []
        for t, dm in st.session_state.dados_mercado.items():
            investido, saldo = dm['Qtd'] * dm['PM'], dm['Qtd'] * dm['Preço Atual']
            linhas_perf.append({
                "Ativo": t, "Tipo": dm["Tipo"], "Setor": dm["Setor"], "Qtd": int(dm['Qtd']), 
                "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
                "Total Investido": investido, "Saldo Atual": saldo, "Resultado (R$)": saldo - investido,
                "Data Média": dm['Data'].strftime('%d/%m/%Y'), "Total Div. (R$)": dm['Div_Total'], 
                "DY on Cost (%)": (dm['Div_Total'] / investido)*100 if investido>0 else 0, 
                "Evolução c/ Div (%)": (((saldo + dm['Div_Total']) / investido)-1)*100 if investido>0 else 0,
                "IPCA Acum. (%)": dm['IPCA'], "CDI Acum. (%)": dm['CDI']
            })
        df_perf_final = pd.DataFrame(linhas_perf)

        st.markdown("### 🏆 Visão Global")
        df_acoes, df_fiis = df_perf_final[df_perf_final['Tipo'] == 'Ação'], df_perf_final[df_perf_final['Tipo'] == 'FII']
        ev_acoes = (df_acoes['Saldo Atual'].sum() / df_acoes['Total Investido'].sum() - 1)*100 if df_acoes['Total Investido'].sum()>0 else 0
        ev_fiis = (df_fiis['Saldo Atual'].sum() / df_fiis['Total Investido'].sum() - 1)*100 if df_fiis['Total Investido'].sum()>0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📈 Patrimônio Ações", f_brl(df_acoes['Saldo Atual'].sum()), f_pct(ev_acoes))
        m2.metric("🏢 Patrimônio FIIs", f_brl(df_fiis['Saldo Atual'].sum()), f_pct(ev_fiis))
        m3.metric("💸 Renda Ações", f_brl(df_acoes['Total Div. (R$)'].sum()))
        m4.metric("💸 Renda FIIs", f_brl(df_fiis['Total Div. (R$)'].sum()))

        t1, t2, t3, t4, t5, t6 = st.tabs(["📊 Visão Geral", "💰 Valuation", "🎯 Radar & Projeção", "📈 Gráficos", "💸 Proventos", "💬 IA"])
        
        with t1:
            st.dataframe(df_perf_final.drop(columns=['Tipo', 'Setor']).style.format({c: f_brl for c in ["Preço Médio", "Preço Atual", "Total Investido", "Saldo Atual", "Resultado (R$)", "Total Div. (R$)"]}|{c: f_pct for c in ["DY on Cost (%)", "Evolução c/ Div (%)", "IPCA Acum. (%)", "CDI Acum. (%)"]}), use_container_width=True, hide_index=True)

        with t2:
            st.markdown("#### Bazin (Renda) & Graham (Valor)")
            yd = st.number_input("Taxa Risco (%):", value=6.0, step=0.5) / 100.0
            df_edit_v = st.data_editor(st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]], use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"])
            st.session_state.df_simul[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]] = df_edit_v[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]]
            
            recs_val = []
            for _, r in df_edit_v.iterrows():
                bz = (r["Div. Projetado (R$)"] / yd) if r["Div. Projetado (R$)"]>0 else 0
                gh = (22.5 * r["LPA Projetado"] * r["VPA (Contábil)"])**0.5 if r["LPA Projetado"]>0 and r["VPA (Contábil)"]>0 else 0
                mbz = ((bz/r["Cotação Atual"])-1)*100 if bz>0 else 0
                mgh = ((gh/r["Cotação Atual"])-1)*100 if gh>0 else 0
                recs_val.append({"Ativo": r['Ativo'], "Teto Bazin": bz, "Margem Bazin (%)": mbz, "Justo Graham": gh, "Margem Graham (%)": mgh})
            st.session_state.df_recs_val = pd.DataFrame(recs_val)
            st.dataframe(st.session_state.df_recs_val.style.format({"Teto Bazin": f_brl, "Justo Graham": f_brl, "Margem Bazin (%)": f_pct, "Margem Graham (%)": f_pct}), use_container_width=True, hide_index=True)

        with t3: 
            c_p1, c_p2, c_p3, c_p4 = st.columns(4)
            patr_fora = c_p1.number_input("Patrimônio Fora (R$):", value=0.0, step=1000.0)
            aporte = c_p2.number_input("Aporte Mensal (R$):", value=2000.0, step=500.0)
            rent = c_p3.number_input("Rentab. Mensal (%):", value=0.8, step=0.1) / 100.0
            cresc_div = c_p4.number_input("Cresc. Anual Div (%):", value=5.0, step=1.0) / 100.0

            st.markdown("##### Radar de Oportunidades")
            df_radar = pd.merge(df_perf_final[['Ativo', 'Tipo']], st.session_state.df_recs_val, on='Ativo')
            mg_ex = st.number_input("Margem Graham Exigida (%):", value=15.0)
            mb_ex = st.number_input("Margem Bazin Exigida (%):", value=5.0)
            
            sts_list = []
            for _, r in df_radar.iterrows():
                if r['Tipo'] == 'Ação': sts_list.append("COMPRA 🟢" if r['Margem Graham (%)']>mg_ex and r['Margem Bazin (%)']>mb_ex else ("MANTER 🟡" if r['Margem Graham (%)']>0 else "VENDA 🔴"))
                else: sts_list.append("COMPRA 🟢" if r['Margem Bazin (%)']>mb_ex else ("MANTER 🟡" if r['Margem Bazin (%)']>-5 else "VENDA 🔴"))
            df_radar['Status'] = sts_list
            st.dataframe(df_radar.style.format({"Teto Bazin": f_brl, "Justo Graham": f_brl, "Margem Bazin (%)": f_pct, "Margem Graham (%)": f_pct}), use_container_width=True, hide_index=True)

            st.markdown("##### ❄️ Projeção Bola de Neve (1 Ano)")
            saldo_corr = df_perf_final['Saldo Atual'].sum() + patr_fora
            base_div = st.session_state.df_simul['Div. Projetado (R$)'].sum() / 12
            ac_ap, ac_jd, linhas_proj = 0.0, 0.0, []
            
            for m in range(13):
                linhas_proj.append({"Mês": f"Mês {m}", "Patrimônio Base": saldo_corr, "Aportes Acum.": ac_ap, "Juros/Divs Acum.": ac_jd})
                gc = saldo_corr * rent
                div_m = base_div * ((1 + cresc_div) ** (m/12))
                ac_jd += (gc + div_m)
                ac_ap += aporte
                saldo_corr += (gc + div_m + aporte)
            
            fig_proj = px.area(pd.DataFrame(linhas_proj).melt(id_vars=["Mês"], var_name="Componente", value_name="Valor"), x="Mês", y="Valor", color="Componente")
            st.plotly_chart(fig_proj, use_container_width=True)

        with t4:
            st.markdown("#### Gráficos da Distribuição")
            c_g1, c_g2 = st.columns(2)
            c_g1.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Ativo', title="Por Ativo"), use_container_width=True)
            c_g2.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Setor', title="Por Setor"), use_container_width=True)
            
            st.markdown("#### 📊 Rentabilidade: Ativos vs Benchmarks (CDI e IPCA)")
            st.markdown("Comparativo do Ganho Real e Nominal desde a aquisição de cada ativo.")
            
            # Preparação dos dados para o Gráfico Restaurado
            df_comp = df_perf_final[['Ativo', 'Evolução c/ Div (%)', 'CDI Acum. (%)', 'IPCA Acum. (%)']].copy()
            df_comp = df_comp.rename(columns={'Evolução c/ Div (%)': 'Carteira (c/ Div)', 'CDI Acum. (%)': 'CDI', 'IPCA Acum. (%)': 'IPCA'})
            df_melt = df_comp.melt(id_vars='Ativo', var_name='Indicador', value_name='Rentabilidade (%)')
            
            fig_comp = px.bar(
                df_melt, 
                x='Ativo', 
                y='Rentabilidade (%)', 
                color='Indicador', 
                barmode='group',
                color_discrete_map={'Carteira (c/ Div)': '#1f77b4', 'CDI': '#ff7f0e', 'IPCA': '#2ca02c'},
                title="Rentabilidade Acumulada por Ativo vs Indexadores"
            )
            fig_comp.update_layout(xaxis_title="Ativo", yaxis_title="Rentabilidade Acumulada (%)", legend_title="Indicador")
            st.plotly_chart(fig_comp, use_container_width=True)

        with t5:
            st.markdown("### 💸 Proventos (Mensais e Exportação)")
            c_f1, c_f2, c_btn = st.columns([2, 2, 2])
            meses_map = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
            m_hoje, a_hoje = pd.Timestamp.now().month, pd.Timestamp.now().year
            
            m_sel = c_f1.selectbox("Mês:", options=list(meses_map.keys()), format_func=lambda x: meses_map[x], index=m_hoje-1)
            a_sel = c_f2.selectbox("Ano:", options=[a_hoje, a_hoje-1, a_hoje-2])
            
            if 'divs_df' not in st.session_state: st.session_state.divs_df = None
            
            if c_btn.button("🔄 Processar Proventos", use_container_width=True) or st.session_state.divs_df is None:
                with st.spinner("Lendo histórico B3..."):
                    l_div = []
                    for t, dm in st.session_state.dados_mercado.items():
                        try:
                            divs = yf.Ticker(f"{t}.SA").dividends
                            if not divs.empty:
                                if divs.index.tz is not None: divs.index = divs.index.tz_localize(None)
                                dm_val = divs[(divs.index.month == m_sel) & (divs.index.year == a_sel)].sum()
                                if dm_val > 0:
                                    yoc = ((dm_val * dm['Qtd']) / (dm['Qtd'] * dm['PM'])) * 100 if dm['PM']>0 else 0
                                    dy = (dm_val / dm['Preço Atual']) * 100 if dm['Preço Atual']>0 else 0
                                    l_div.append({"Ativo": t, "Unitário (R$)": float(dm_val), "Qtd": int(dm['Qtd']), "Recebido (R$)": float(dm_val * dm['Qtd']), "Yield on Cost (%)": float(yoc), "DY Atual (%)": float(dy)})
                        except: pass
                    st.session_state.divs_df = pd.DataFrame(l_div)
                    st.session_state.divs_m, st.session_state.divs_a = m_sel, a_sel
            
            df_d = st.session_state.divs_df
            if df_d is not None and not df_d.empty:
                st.dataframe(df_d.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct, "DY Atual (%)": f_pct}), use_container_width=True, hide_index=True)
                st.success(f"**Total {meses_map[st.session_state.divs_m]}/{st.session_state.divs_a}:** {f_brl(df_d['Recebido (R$)'].sum())}")
                
                xls = to_excel(df_d, sheet_name=f"Proventos_{meses_map[st.session_state.divs_m]}")
                st.download_button(label="📥 Baixar Relatório de Proventos (Excel)", data=xls, file_name=f"Proventos_{st.session_state.username}_{meses_map[st.session_state.divs_m]}_{st.session_state.divs_a}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
            elif df_d is not None: st.info("Sem proventos no período selecionado.")

        with t6:
            st.info("Utilize a caixa de chat abaixo para análise da carteira com IA.")

# ==========================================
# 9. COMITÊ DE IA (CHAT)
# ==========================================
st.write("---")
cc1, cc2 = st.columns([10, 2])
with cc1: st.markdown("### 💬 Comitê de IA")
with cc2:
    if st.button("🗑️ Limpar Chat", use_container_width=True):
        st.session_state.historico_chat = MENSAGEM_INICIAL.copy()
        if os.path.exists(ARQUIVO_CHAT): os.remove(ARQUIVO_CHAT)
        st.rerun()

try: api_key = st.secrets["GEMINI_API_KEY"]
except: api_key = st.text_input("Gemini API Key:", type="password")

for msg in st.session_state.historico_chat:
    with st.chat_message(msg["role"]): st.write(msg["content"])
    
if prompt := st.chat_input("Pergunte à Gestora IA..."):
    st.session_state.historico_chat.append({"role": "user", "content": prompt})
    salvar_chat()
    with st.chat_message("user"): st.write(prompt)
    
    ctx_c = "Vazia." if st.session_state.df_base.empty else df_perf_final[['Ativo', 'Qtd', 'Preço Médio', 'Preço Atual', 'Evolução c/ Div (%)']].to_csv(index=False)
    ctx_m = f"Selic: {f_pct(selic_hoje)}|IPCA: {f_pct(ipca_12m_hoje)}. Focus {ano_atual}: Sel {f_pct(proj_focus.get(f'Selic_{ano_atual}'))}/IPCA {f_pct(proj_focus.get(f'IPCA_{ano_atual}'))}"
    sys_prompt = f"Você é Analista Sênior CNPI. [Dados]: {ctx_c}. [Macro]: {ctx_m}. Seja executiva, cruze dados."
    
    if api_key:
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            resp_ok = False
            for m in ['gemini-3.5-flash', 'gemini-3.1-flash-lite', 'gemini-2.5-flash', 'gemini-1.5-flash']:
                try:
                    resposta = genai.GenerativeModel(m).generate_content([sys_prompt, prompt]).text
                    resp_ok = True
                    break 
                except: continue
            if not resp_ok: resposta = "⚠️ Falha de rede com a IA do Google."
        except Exception as e: resposta = f"⚠️ Erro estrutural: {e}"
    else: resposta = "⚠️ Chave API ausente."
    
    st.session_state.historico_chat.append({"role": "assistant", "content": resposta})
    salvar_chat()
    st.rerun()
