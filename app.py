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

try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==========================================
# 1. CONFIGURAÇÃO E FORMATADORES
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")

def f_brl(x): return f"R$ {float(x):,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
def f_brl_4(x): return f"R$ {float(x):,.4f}".replace(",", "v").replace(".", ",").replace("v", ".")
def f_pct(x): return f"{float(x):,.2f}%".replace(",", "v").replace(".", ",").replace("v", ".")

def to_excel(df, sheet_name='Sheet1'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        for col_num, column in enumerate(worksheet.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            for row_num in range(2, worksheet.max_row + 1):
                data_cell = worksheet.cell(row=row_num, column=col_num)
                data_cell.font = Font(name="Arial", size=10)
                data_cell.border = thin_border
                if isinstance(data_cell.value, (int, float)): data_cell.alignment = Alignment(horizontal="right")
                else: data_cell.alignment = Alignment(horizontal="center")
            max_len = max(len(str(c.value or '')) for c in column)
            col_letter = get_column_letter(col_num)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 13)
    return output.getvalue()

def export_docx(historico):
    if not HAS_DOCX: return None
    doc = docx.Document()
    doc.add_heading('Histórico - Comitê de IA CNPI', 0)
    for msg in historico:
        role = "Analista:" if msg["role"] == "user" else "Gestora IA:"
        doc.add_heading(role, level=2)
        doc.add_paragraph(msg["content"])
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

MAPEAMENTO_TICKERS = {"GALG11": "GARE11", "SOMA3": "ALOS3", "ARZZ3": "ALOS3", "VVAR3": "BHIA3", "VIIA3": "BHIA3", "BRML3": "ALSO3", "BBRK11": "BRCR11", "HCTR11": "TRXD11", "TORD11": "TRXD11"}
UNITS_ACOES = ['SANB11', 'TAEE11', 'KLBN11', 'BPAC11', 'ALUP11', 'ENGI11', 'BIDI11', 'CPLE11', 'SAPR11', 'RNEW11']

if 'df_base' not in st.session_state: st.session_state.df_base = pd.DataFrame()
if 'dados_mercado' not in st.session_state: st.session_state.dados_mercado = {}
if 'df_simul' not in st.session_state: st.session_state.df_simul = pd.DataFrame()
if 'df_tesouro' not in st.session_state: st.session_state.df_tesouro = pd.DataFrame()
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'username' not in st.session_state: st.session_state.username = ""
if 'historico_chat' not in st.session_state: st.session_state.historico_chat = []

# ==========================================
# 2. MOTOR DE BANCO DE DADOS HÍBRIDO (POSTGRESQL / SQLITE)
# ==========================================
IS_POSTGRES = "POSTGRES_URL" in st.secrets
PARAM = "%s" if IS_POSTGRES else "?"

def get_db_connection():
    if IS_POSTGRES:
        import psycopg2
        return psycopg2.connect(st.secrets["POSTGRES_URL"])
    else:
        return sqlite3.connect("terminal_cnpi.db")

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (username TEXT PRIMARY KEY, password TEXT)''')
    if IS_POSTGRES:
        c.execute('''CREATE TABLE IF NOT EXISTS carteiras (username TEXT, ativo TEXT, quantidade DOUBLE PRECISION, preco_medio DOUBLE PRECISION, data_media TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS tesouro (username TEXT, titulo TEXT, investido DOUBLE PRECISION, taxa DOUBLE PRECISION, vencimento INTEGER)''')
        c.execute('''CREATE TABLE IF NOT EXISTS chat_ia (username TEXT, role TEXT, content TEXT)''')
    else:
        c.execute('''CREATE TABLE IF NOT EXISTS carteiras (username TEXT, Ativo TEXT, Quantidade REAL, Preco_Medio REAL, Data_Media TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS tesouro (username TEXT, titulo TEXT, investido REAL, taxa REAL, vencimento INTEGER)''')
        c.execute('''CREATE TABLE IF NOT EXISTS chat_ia (username TEXT, role TEXT, content TEXT)''')
    conn.commit()
    conn.close()

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()

def registrar_usuario(username, password):
    conn = get_db_connection()
    c = conn.cursor()
    try:
        c.execute(f"INSERT INTO usuarios (username, password) VALUES ({PARAM}, {PARAM})", (username, hash_password(password)))
        conn.commit()
        conn.close()
        return True
    except:
        conn.close()
        return False

def autenticar_usuario(username, password):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute(f"SELECT * FROM usuarios WHERE username={PARAM} AND password={PARAM}", (username, hash_password(password)))
    user = c.fetchone()
    conn.close()
    return user is not None

def atualizar_senha(username, nova_senha):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute(f"SELECT * FROM usuarios WHERE username={PARAM}", (username,))
    user = c.fetchone()
    if user is None:
        conn.close()
        return False
    c.execute(f"UPDATE usuarios SET password={PARAM} WHERE username={PARAM}", (hash_password(nova_senha), username))
    conn.commit()
    conn.close()
    return True

def salvar_dados_usuario(username):
    conn = get_db_connection()
    c = conn.cursor()
    
    # Salva Carteira
    c.execute(f"DELETE FROM carteiras WHERE username={PARAM}", (username,))
    if not st.session_state.df_base.empty:
        for _, row in st.session_state.df_base.iterrows():
            c.execute(f"INSERT INTO carteiras (username, Ativo, Quantidade, Preco_Medio, Data_Media) VALUES ({PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM})",
                      (username, row['Ativo'], float(row['Quantidade']), float(row['Preço Médio']), str(row['Data Média'])))
            
    # Salva Tesouro
    c.execute(f"DELETE FROM tesouro WHERE username={PARAM}", (username,))
    if not st.session_state.df_tesouro.empty:
        for _, row in st.session_state.df_tesouro.iterrows():
            c.execute(f"INSERT INTO tesouro (username, titulo, investido, taxa, vencimento) VALUES ({PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM})",
                      (username, row['Título'], float(row['Investimento (R$)']), float(row['Taxa Anual (%)']), int(row['Ano Venc.'])))
            
    # Salva Chat IA
    c.execute(f"DELETE FROM chat_ia WHERE username={PARAM}", (username,))
    for msg in st.session_state.historico_chat[-30:]: # Guarda ultimas 30 mensagens
        c.execute(f"INSERT INTO chat_ia (username, role, content) VALUES ({PARAM}, {PARAM}, {PARAM})", (username, msg['role'], msg['content']))
        
    conn.commit()
    conn.close()

def carregar_dados_usuario(username):
    conn = get_db_connection()
    df_carteira = pd.read_sql_query(f"SELECT Ativo, Quantidade, Preco_Medio as \"Preço Médio\", Data_Media as \"Data Média\" FROM carteiras WHERE username={PARAM}", conn, params=(username,))
    df_tesouro = pd.read_sql_query(f"SELECT titulo as \"Título\", investido as \"Investimento (R$)\", taxa as \"Taxa Anual (%)\", vencimento as \"Ano Venc.\" FROM tesouro WHERE username={PARAM}", conn, params=(username,))
    df_chat = pd.read_sql_query(f"SELECT role, content FROM chat_ia WHERE username={PARAM}", conn, params=(username,))
    conn.close()
    
    if not df_carteira.empty: df_carteira['Data Média'] = pd.to_datetime(df_carteira['Data Média']).dt.date
    st.session_state.df_base = df_carteira
    st.session_state.df_tesouro = df_tesouro
    
    if not df_chat.empty:
        st.session_state.historico_chat = df_chat.to_dict('records')
    else:
        st.session_state.historico_chat = [{"role": "assistant", "content": f"Saudações, {username}. O terminal está mapeado em tempo real. Como posso ajudar com o mercado ou com sua carteira?"}]

init_db()

# ==========================================
# 3. TELA DE AUTENTICAÇÃO (GATEKEEPER)
# ==========================================
if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center;'>🔐 Terminal de Gestão Profissional</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Acesso restrito. Identifique-se para carregar seu portfólio.</p>", unsafe_allow_html=True)
    
    col_log1, col_log2, col_log3 = st.columns([1, 1, 1])
    with col_log2:
        tab_login, tab_register, tab_forgot = st.tabs(["Acesso", "Novo Registro", "Esqueci a Senha"])
        
        with tab_login:
            login_user = st.text_input("Usuário", key="log_user")
            login_pass = st.text_input("Senha", type="password", key="log_pass")
            if st.button("Entrar", use_container_width=True):
                if autenticar_usuario(login_user, login_pass):
                    st.session_state.logged_in = True
                    st.session_state.username = login_user
                    carregar_dados_usuario(login_user)
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
                
        with tab_forgot:
            st.markdown("<p style='font-size: 14px; color: gray;'>Informe seu usuário cadastrado e a nova senha desejada.</p>", unsafe_allow_html=True)
            forgot_user = st.text_input("Confirmar Usuário", key="forgot_user")
            forgot_pass = st.text_input("Nova Senha", type="password", key="forgot_pass")
            if st.button("Redefinir Senha", use_container_width=True):
                if forgot_user and forgot_pass:
                    if atualizar_senha(forgot_user, forgot_pass):
                        st.success("Senha redefinida com sucesso! Volte na aba 'Acesso' para entrar.")
                    else: st.error("Usuário não encontrado no sistema.")
                else: st.warning("Preencha ambos os campos.")
                
    st.stop()

# ==========================================
# 4. APP PRINCIPAL: FUNÇÕES DE DADOS B3 E MACRO
# ==========================================
st.markdown(f"""
    <div style="text-align: center; margin-bottom: 20px;">
        <h3 style="font-weight: 400; margin-bottom: 0;">💼 Terminal de Gestão</h3>
        <h6 style="color: #666; font-weight: 300;">Analista: {st.session_state.username.upper()}</h6>
    </div>
""", unsafe_allow_html=True)
st.write("---")

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

def corrigir_cabecalho_b3(df):
    if df.empty: return df
    if 'Data do Negócio' in df.columns or 'Data Média' in df.columns: 
        return df
        
    def tentar_mapear(lista_colunas):
        lista_upper = [str(c).strip().upper() for c in lista_colunas]
        is_negociacao = any('DATA DO' in c or 'NEGÓCIO' in c or 'NEGOCIO' in c for c in lista_upper)
        is_movimentacao = any('PRODUTO' in c or 'ATIVO' in c for c in lista_upper) and any('MOVIMENTA' in c or 'MOVIMEN' in c for c in lista_upper)
        
        if is_negociacao or is_movimentacao:
            mapeamento = {}
            for col in lista_colunas:
                c_up = str(col).strip().upper()
                if 'DAT' in c_up: mapeamento[col] = 'Data do Negócio'
                elif 'PROD' in c_up or 'ATIV' in c_up or 'CÓD' in c_up or 'COD' in c_up or 'PAPEL' in c_up: mapeamento[col] = 'Código de Negociação'
                elif 'MOV' in c_up or 'TIP' in c_up: mapeamento[col] = 'Tipo de Movimentação'
                elif 'VAL' in c_up: mapeamento[col] = 'Valor'
                elif 'QUA' in c_up or 'QTD' in c_up: mapeamento[col] = 'Quantidade'
                elif 'PRE' in c_up: mapeamento[col] = 'Preço Unitário'
                elif 'ENTRADA' in c_up or 'SAÍDA' in c_up or 'SAIDA' in c_up: mapeamento[col] = 'Entrada/Saída'
            return mapeamento, is_movimentacao
        return None, False

    mapeamento, is_mov = tentar_mapear(df.columns.tolist())
    if mapeamento and ('Data do Negócio' in mapeamento.values() or 'Código de Negociação' in mapeamento.values()):
        df = df.rename(columns=mapeamento)
        if is_mov and 'Valor' not in df.columns and 'Preço Unitário' in df.columns and 'Quantidade' in df.columns:
            df['Valor'] = df['Quantidade'].apply(limpar_numero) * df['Preço Unitário'].apply(limpar_numero)
        return df

    for i in range(min(30, len(df))):
        linha = df.iloc[i].tolist()
        mapeamento, is_mov = tentar_mapear(linha)
        if mapeamento:
            df.columns = linha
            df_sub = df.iloc[i+1:].reset_index(drop=True)
            df_sub = df_sub.rename(columns=mapeamento)
            if is_mov and 'Valor' not in df_sub.columns and 'Preço Unitário' in df_sub.columns and 'Quantidade' in df_sub.columns:
                df_sub['Valor'] = df_sub['Quantidade'].apply(limpar_numero) * df_sub['Preço Unitário'].apply(limpar_numero)
            return df_sub
            
    return df

def processar_planilha_b3(df):
    if df.empty: return pd.DataFrame()
    
    df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
    df['Quantidade'] = df['Quantidade'].apply(limpar_numero)
    
    if 'Valor' in df.columns: df['Valor'] = df['Valor'].apply(limpar_numero)
    elif 'Preço Unitário' in df.columns: df['Valor'] = df['Quantidade'] * df['Preço Unitário'].apply(limpar_numero)
    else: df['Valor'] = 0.0

    df = df.sort_values('Data do Negócio')
    
    posicoes = {}
    for _, row in df.iterrows():
        if pd.isna(row['Código de Negociação']): continue
        
        ticker_raw = str(row['Código de Negociação']).strip().upper()
        if " - " in ticker_raw: ticker_raw = ticker_raw.split(" - ")[0].strip()
        elif " " in ticker_raw: ticker_raw = ticker_raw.split(" ")[0].strip()
            
        ticker = MAPEAMENTO_TICKERS.get(ticker_raw[:-1] if ticker_raw.endswith('F') and len(ticker_raw) > 4 else ticker_raw, ticker_raw[:-1] if ticker_raw.endswith('F') and len(ticker_raw) > 4 else ticker_raw)
        
        if not re.match(r'^[A-Z]{4}\d{1,2}$', ticker): continue
            
        if 'Tipo de Movimentação' in df.columns:
            mov_tipo_str = str(row['Tipo de Movimentação']).strip().upper()
            if any(p in mov_tipo_str for p in ['RENDIMENTO', 'JUROS', 'DIVIDENDO', 'JCP', 'REEMBOLSO']):
                continue
        
        qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio'] if pd.notna(row['Data do Negócio']) else pd.Timestamp.now()
        if ticker not in posicoes: posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'ts_medio': 0.0}
        
        is_compra = False
        is_venda = False
        
        if 'Entrada/Saída' in df.columns and pd.notna(row['Entrada/Saída']):
            io_dir = str(row['Entrada/Saída']).strip().upper()
            if 'CRED' in io_dir or 'ENT' in io_dir: is_compra = True
            elif 'DEB' in io_dir or 'SAI' in io_dir: is_venda = True
        
        if not is_compra and not is_venda and 'Tipo de Movimentação' in df.columns:
            tipo_mov = str(row['Tipo de Movimentação']).strip().upper()
            if any(term in tipo_mov for term in ['COMPRA', 'CREDITO', 'CRÉDITO', 'LIQUIDAÇÃO', 'LIQUIDACAO', 'TRANSFERENCIA', 'TRANSFERÊNCIA']) or tipo_mov == 'C':
                is_compra = True
            elif any(term in tipo_mov for term in ['VENDA', 'DEBITO', 'DÉBITO']) or tipo_mov == 'V':
                is_venda = True
        
        if is_compra:
            q_ant, ts_ant = posicoes[ticker]['qtd'], posicoes[ticker]['ts_medio']
            ts_novo = pd.Timestamp(data).timestamp()
            posicoes[ticker]['ts_medio'] = ts_novo if q_ant == 0 else ((ts_ant * q_ant) + (ts_novo * qtd)) / (q_ant + qtd)
            posicoes[ticker]['qtd'] += qtd
            posicoes[ticker]['valor'] += valor
        elif is_venda:
            if qtd >= (posicoes[ticker]['qtd'] - 0.001): posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'ts_medio': 0.0}
            else:
                pm = posicoes[ticker]['valor'] / posicoes[ticker]['qtd'] if posicoes[ticker]['qtd'] > 0 else 0
                posicoes[ticker]['qtd'] -= qtd
                posicoes[ticker]['valor'] -= (qtd * pm)
                
    ativos = [{"Ativo": t, "Quantidade": d['qtd'], "Preço Médio": d['valor']/d['qtd'] if d['qtd']>0 else 0, "Data Média": pd.to_datetime(d['ts_medio'], unit='s').date()} for t, d in posicoes.items() if d['qtd']>0]
    return consolidar_carteira(pd.DataFrame(ativos))

# ==========================================
# 5. SIDEBAR: UPLOAD E DB
# ==========================================
st.sidebar.markdown(f"### 👤 ANALISTA OPERACIONAL")
if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.df_base = pd.DataFrame()
    st.rerun()

st.sidebar.divider()
st.sidebar.markdown("### 💾 Banco de Dados")
if st.sidebar.button("Salvar Estado Atual no DB", type="primary", use_container_width=True):
    salvar_dados_usuario(st.session_state.username)
    st.sidebar.success("Sincronizado e Salvo!")

st.sidebar.divider()
st.sidebar.header("1. Upload de Arquivos")
arquivo_principal = st.sidebar.file_uploader("Substituir Base Completa", type=["xlsx", "csv"])
arquivo_novo = st.sidebar.file_uploader("Apenas Novas Operações", type=["xlsx", "csv"])
data_corte = st.sidebar.date_input("Filtrar a partir de:", pd.Timestamp.now().date() - pd.Timedelta(days=15)) if arquivo_novo else None

if st.sidebar.button("🚀 Processar Excel B3", use_container_width=True):
    base_atual = st.session_state.df_base.copy()
    
    if arquivo_principal:
        txt = arquivo_principal.getvalue().decode('utf-8-sig', errors='ignore') if arquivo_principal.name.endswith('.csv') else None
        if txt:
            linhas_amostra = "".join(txt.split('\n')[:5])
            sep_detectado = '\t' if linhas_amostra.count('\t') > linhas_amostra.count(';') and linhas_amostra.count('\t') > linhas_amostra.count(',') else (';' if linhas_amostra.count(';') >= linhas_amostra.count(',') else ',')
            df_p = pd.read_csv(io.StringIO(txt), sep=sep_detectado)
        else:
            df_p = pd.read_excel(arquivo_principal)
        
        df_p = corrigir_cabecalho_b3(df_p)
        if 'Data Média' in df_p.columns: base_atual = consolidar_carteira(df_p)
        elif 'Data do Negócio' in df_p.columns: base_atual = processar_planilha_b3(df_p)
        else:
            st.sidebar.error("Formato inválido. Assegure-se de que a planilha possui colunas nativas da B3.")
            st.stop()
            
    if arquivo_novo and not base_atual.empty:
        txt_n = arquivo_novo.getvalue().decode('utf-8-sig', errors='ignore') if arquivo_novo.name.endswith('.csv') else None
        if txt_n:
            linhas_amostra_n = "".join(txt_n.split('\n')[:5])
            sep_detectado_n = '\t' if linhas_amostra_n.count('\t') > linhas_amostra_n.count(';') and linhas_amostra_n.count('\t') > linhas_amostra_n.count(',') else (';' if linhas_amostra_n.count(';') >= linhas_amostra_n.count(',') else ',')
            df_n = pd.read_csv(io.StringIO(txt_n), sep=sep_detectado_n)
        else:
            df_n = pd.read_excel(arquivo_novo)
        
        df_n = corrigir_cabecalho_b3(df_n)
        if not df_n.empty and 'Data do Negócio' in df_n.columns:
            df_n['Data do Negócio'] = pd.to_datetime(df_n['Data do Negócio'], dayfirst=True, errors='coerce')
            df_n = df_n[df_n['Data do Negócio'].dt.date >= data_corte]
            linhas_b = [{"Código de Negociação": r['Ativo'], "Tipo de Movimentação": "Compra", "Data do Negócio": pd.to_datetime(r['Data Média']), "Quantidade": r['Quantidade'], "Valor": r['Quantidade']*r['Preço Médio']} for _, r in base_atual.iterrows()]
            base_atual = processar_planilha_b3(pd.concat([pd.DataFrame(linhas_b), df_n], ignore_index=True))
            
    st.session_state.df_base = base_atual
    st.sidebar.warning("Memória atualizada. Salve no DB para manter.")
    st.rerun()

# ==========================================
# 6. PAINEL MACRO E CONTROLE MANUAL
# ==========================================
proj_focus, ano_atual = obter_projecoes_focus()
selic_hoje, ipca_12m_hoje = obter_macro_atual()

st.markdown("### 👑 Conjuntura Macroeconômica")
c_m1, c_m2 = st.columns([1, 2])
c_m1.success(f"🎯 **Cenário Atual (Vigente)**\n\nSelic Atual: **{f_pct(selic_hoje)} a.a.**\n\nIPCA 12 meses: **{f_pct(ipca_12m_hoje)}**")
c_m2.info(f"🔮 **Projeções do Mercado (Focus)**\n\n**Selic:** {ano_atual}: **{f_pct(proj_focus.get(f'Selic_{ano_atual}', 0))}** |  {ano_atual+1}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+1}', 0))}** |  {ano_atual+2}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+2}', 0))}**\n\n**IPCA:** {ano_atual}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual}', 0))}** |  {ano_atual+1}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+1}', 0))}** |  {ano_atual+2}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+2}', 0))}**")
st.write("---")

st.markdown("### 2. Controle Operacional Manual")
ca, cb, cc = st.columns([1, 1, 1])
with ca:
    tdel = st.selectbox("Excluir Ativo:", [""] + sorted(st.session_state.df_base["Ativo"].tolist()) if not st.session_state.df_base.empty else [""])
    if st.button("Remover", use_container_width=True) and tdel:
        st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != tdel]
        st.rerun()
with cb:
    nt = st.text_input("Nova Compra (Ticker)")
    cq, cp = st.columns(2)
    nq = cq.number_input("Qtd", min_value=1)
    np_v = cp.number_input("PM (R$)", min_value=0.01)
    if st.button("Adicionar à Carteira", use_container_width=True) and nt:
        nl = pd.DataFrame([{"Ativo": nt.upper(), "Quantidade": float(nq), "Preço Médio": float(np_v), "Data Média": pd.Timestamp.now().date()}])
        st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, nl], ignore_index=True))
        st.rerun()
with cc:
    st.info("Insira ativos manualmente se não quiser usar a planilha da B3. Todas as abas são acessíveis livremente.")
    if st.button("🚀 Conectar ao Mercado Vivo", type="primary", use_container_width=True):
        if not st.session_state.df_base.empty:
            st.session_state.df_base = consolidar_carteira(st.session_state.df_base) 
            df_macro, fundamentos_br = carregar_macro(), obter_fundamentos_brasil()
            progresso, total = st.progress(0), len(st.session_state.df_base)
            dados_mercado, lines_simul_iniciais = {}, []

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
                lines_simul_iniciais.append({"Ativo": ticker, "Cotação Atual": preco_atual, "VPA (Contábil)": vpa, "LPA Projetado": lpa, "Div. Projetado (R$)": divs_12m})
                progresso.progress((i + 1) / total)
                
            st.session_state.dados_mercado = dados_mercado
            st.session_state.df_simul = pd.DataFrame(lines_simul_iniciais)
            st.success("Sincronizado!")
        else:
            st.warning("Adicione ativos na carteira primeiro.")

st.write("---")

# ==========================================
# 7. DASHBOARD E RELATÓRIOS (TABS INDEPENDENTES)
# ==========================================
t1, t2, t3, t4, t5, t_tesouro, t6 = st.tabs(["📊 Visão Geral", "💰 Valuation", "🎯 Radar & Projeção", "📈 Gráficos", "💸 Proventos B3", "🏛️ Tesouro Direto", "💬 Gestora IA (CNPI)"])

if st.session_state.dados_mercado:
    linhas_perf = []
    for t, dm in st.session_state.dados_mercado.items():
        investido, saldo = dm['Qtd'] * dm['PM'], dm['Qtd'] * dm['Preço Atual']
        linhas_perf.append({
            "Ativo": t, "Tipo": dm["Tipo"], "Setor": dm["Setor"], "Qtd": int(dm['Qtd']), 
            "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
            "Total Investido": investido, "Saldo Atual": saldo, 
            "Saldo C/ Dividendos": saldo + dm['Div_Total'],
            "Resultado (R$)": saldo - investido,
            "Resultado C/ Dividendos": (saldo - investido) + dm['Div_Total'],
            "Data Média": dm['Data'].strftime('%d/%m/%Y'), "Total Div. (R$)": dm['Div_Total'], 
            "DY on Cost (%)": (dm['Div_Total'] / investido)*100 if investido>0 else 0, 
            "Evolução c/ Div (%)": (((saldo + dm['Div_Total']) / investido)-1)*100 if investido>0 else 0,
            "IPCA Acum. (%)": dm['IPCA'], "CDI Acum. (%)": dm['CDI']
        })
    df_perf_final = pd.DataFrame(linhas_perf)

    with t1:
        st.markdown("### 🏆 Visão Global")
        df_acoes, df_fiis = df_perf_final[df_perf_final['Tipo'] == 'Ação'], df_perf_final[df_perf_final['Tipo'] == 'FII']
        ev_acoes = (df_acoes['Saldo Atual'].sum() / df_acoes['Total Investido'].sum() - 1)*100 if df_acoes['Total Investido'].sum()>0 else 0
        ev_fiis = (df_fiis['Saldo Atual'].sum() / df_fiis['Total Investido'].sum() - 1)*100 if df_fiis['Total Investido'].sum()>0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📈 Patrimônio Ações", f_brl(df_acoes['Saldo Atual'].sum()), f_pct(ev_acoes))
        m2.metric("🏢 Patrimônio FIIs", f_brl(df_fiis['Saldo Atual'].sum()), f_pct(ev_fiis))
        m3.metric("💸 Renda Ações", f_brl(df_acoes['Total Div. (R$)'].sum()))
        m4.metric("💸 Renda FIIs", f_brl(df_fiis['Total Div. (R$)'].sum()))

        st.dataframe(df_perf_final.drop(columns=['Tipo', 'Setor']).style.format({c: f_brl for c in ["Preço Médio", "Preço Atual", "Total Investido", "Saldo Atual", "Saldo C/ Dividendos", "Resultado (R$)", "Resultado C/ Dividendos", "Total Div. (R$)"]}|{c: f_pct for c in ["DY on Cost (%)", "Evolução c/ Div (%)", "IPCA Acum. (%)", "CDI Acum. (%)"]}), use_container_width=True, hide_index=True)

    with t2:
        st.markdown("#### Métodos Certificados de Valuation")
        st.markdown("""
        * **Preço Teto Decio Bazin:** Avalia se a empresa paga bons dividendos hoje. Calcula o preço máximo ideal de compra para garantir um retorno mínimo.
        * **Preço Justo Benjamin Graham:** Avalia o valor real de fábrica da empresa. Indica se o preço da ação está barato ou caro na bolsa. *(Não se aplica a FIIs).*
        """)
        yd = st.number_input("Taxa de Retorno Mínima Exigida Bazin (%):", value=6.0, step=0.5) / 100.0
        
        df_edit_v = st.data_editor(st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]], use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"])
        st.session_state.df_simul[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]] = df_edit_v[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]]
        
        recs_val = []
        for _, r in df_edit_v.iterrows():
            t_ticker = str(r['Ativo']).strip().upper()
            is_fii = t_ticker.endswith('11') and t_ticker not in UNITS_ACOES
            bz = (float(r["Div. Projetado (R$)"]) / yd) if float(r["Div. Projetado (R$)"]) > 0 else 0.0
            mbz = ((bz / float(r["Cotação Atual"])) - 1) * 100 if bz > 0 else 0.0
            
            if not is_fii:
                gh = (22.5 * float(r["LPA Projetado"]) * float(r["VPA (Contábil)"]))**0.5 if float(r["LPA Projetado"]) > 0 and float(r["VPA (Contábil)"]) > 0 else 0.0
                mgh = ((gh / float(r["Cotação Atual"])) - 1) * 100 if gh > 0 else 0.0
            else:
                gh = np.nan
                mgh = np.nan
            recs_val.append({"Ativo": t_ticker, "Teto Bazin": bz, "Margem Bazin (%)": mbz, "Justo Graham": gh, "Margem Graham (%)": mgh})
            
        st.session_state.df_recs_val = pd.DataFrame(recs_val)
        st.dataframe(st.session_state.df_recs_val.style.format({"Teto Bazin": lambda x: f_brl(x) if x > 0 else "-", "Justo Graham": lambda x: f_brl(x) if pd.notna(x) and x > 0 else "-", "Margem Bazin (%)": lambda x: f_pct(x) if x != 0 else "-", "Margem Graham (%)": lambda x: f_pct(x) if pd.notna(x) and x != 0 else "-"}), use_container_width=True, hide_index=True)

    with t3: 
        st.markdown("##### Parametrização do Radar Operacional")
        c_p1, c_p2, c_p3, c_p4 = st.columns(4)
        patr_fora = c_p1.number_input("Patrimônio Externo (R$):", value=0.0, step=1000.0)
        aporte = c_p2.number_input("Aporte Mensal (R$):", value=2000.0, step=500.0)
        rent = c_p3.number_input("Rentabilidade Mensal Alvo (%):", value=0.8, step=0.1) / 100.0
        cresc_div = c_p4.number_input("Cresc. Anual Dividendos (%):", value=5.0, step=1.0) / 100.0

        st.markdown("##### 🎯 Triagem Estratégica e Teoria das Margens")
        st.info("**Por que olhar a Margem de Segurança?**\n\n**Graham (>15% a 20%):** Protege seu patrimônio contra erros contábeis e flutuações bruscas. Você compra o ativo com 'desconto de fábrica'.\n\n**Bazin (>0% a 5%):** Garante que, mesmo que a cotação congele, o dinheiro vivo em dividendos será no mínimo o que você exigiu.")
        
        c_m1, c_m2 = st.columns(2)
        mb_ex = c_m1.number_input("Margem Mínima Bazin Exigida (%):", value=5.0)
        mg_ex = c_m2.number_input("Margem Mínima Graham Exigida (%):", value=15.0)
        
        # Radar recálculo dinâmico baseado no input atual
        df_radar = pd.merge(df_perf_final[['Ativo', 'Tipo', 'Preço Atual']], st.session_state.df_recs_val, on='Ativo')
        
        df_radar['Status Bazin'] = df_radar.apply(lambda r: "COMPRA 🟢" if r['Teto Bazin']>0 and r['Margem Bazin (%)'] >= mb_ex else ("MANTER 🟡" if r['Teto Bazin']>0 and r['Margem Bazin (%)'] >= -5 else "VENDA 🔴"), axis=1)
        df_radar['Status Graham'] = df_radar.apply(lambda r: "COMPRA 🟢" if r['Tipo']=='Ação' and pd.notna(r['Justo Graham']) and r['Margem Graham (%)'] >= mg_ex else ("MANTER 🟡" if r['Tipo']=='Ação' and pd.notna(r['Justo Graham']) and r['Margem Graham (%)'] >= 0 else ("VENDA 🔴" if r['Tipo']=='Ação' else "-")), axis=1)
        
        st.dataframe(df_radar[['Ativo', 'Tipo', 'Preço Atual', 'Teto Bazin', 'Margem Bazin (%)', 'Status Bazin', 'Justo Graham', 'Margem Graham (%)', 'Status Graham']].style.format({"Preço Atual": f_brl, "Teto Bazin": lambda x: f_brl(x) if x > 0 else "-", "Justo Graham": lambda x: f_brl(x) if pd.notna(x) and x > 0 else "-", "Margem Bazin (%)": lambda x: f_pct(x) if x != 0 else "-", "Margem Graham (%)": lambda x: f_pct(x) if pd.notna(x) and x != 0 else "-"}), use_container_width=True, hide_index=True)

        st.markdown("##### ❄️ Projeção Bola de Neve (1 Ano)")
        saldo_inicial = df_perf_final['Saldo Atual'].sum() + patr_fora
        base_div = st.session_state.df_simul['Div. Projetado (R$)'].sum() / 12 if not st.session_state.df_simul.empty else 0.0
        
        ac_ap, ac_jd, linhas_proj = 0.0, 0.0, []
        saldo_dinamico = saldo_inicial
        
        for m in range(13):
            if m > 0:
                gc = saldo_dinamico * rent
                div_m = base_div * ((1 + cresc_div) ** (m/12))
                ac_jd += (gc + div_m)
                ac_ap += aporte
                saldo_dinamico += (gc + div_m + aporte)
            linhas_proj.append({"Mês": f"Mês {m}", "Capital Inicial": saldo_inicial, "Aportes Acumulados": ac_ap, "Juros/Divs Acumulados": ac_jd})
        
        df_melt_proj = pd.DataFrame(linhas_proj).melt(id_vars=["Mês"], value_vars=["Capital Inicial", "Aportes Acumulados", "Juros/Divs Acumulados"], var_name="Componente", value_name="Valor (R$)")
        fig_proj = px.bar(df_melt_proj, x="Mês", y="Valor (R$)", color="Componente", title="Evolução Patrimonial Controlada", color_discrete_sequence=['#1f4e78', '#00a896', '#f4a261'])
        st.plotly_chart(fig_proj, use_container_width=True)

    with t4:
        st.markdown("#### Gráficos de Distribuição Patrimonial")
        c_g1, c_g2 = st.columns(2)
        cores_modernas = ['#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087', '#f95d6a', '#ff7c43', '#ffa600']
        cores_tipo = ['#1f4e78', '#00a896']
        c_g1.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Ativo', title="Por Ativo", color_discrete_sequence=cores_modernas), use_container_width=True)
        c_g2.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Tipo', title="Por Classe (FII vs Ação)", color_discrete_sequence=cores_tipo), use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 📊 Gráfico Dinâmico Comparativo Histórico")
        ativos_disponiveis = sorted(df_perf_final['Ativo'].unique().tolist())
        c_f_g1, c_f_g2 = st.columns(2)
        ativos_sel = c_f_g1.multiselect("Selecione Ativos:", options=ativos_disponiveis, default=ativos_disponiveis[:5] if len(ativos_disponiveis)>=5 else ativos_disponiveis)
        indexadores_sel = c_f_g2.multiselect("Indexadores:", ['CDI', 'IPCA'], default=['CDI', 'IPCA'])
        
        if ativos_sel:
            df_comp = df_perf_final[df_perf_final['Ativo'].isin(ativos_sel)][['Ativo', 'Evolução c/ Div (%)', 'CDI Acum. (%)', 'IPCA Acum. (%)']].rename(columns={'Evolução c/ Div (%)': 'Carteira (c/ Div)', 'CDI Acum. (%)': 'CDI', 'IPCA Acum. (%)': 'IPCA'})
            col_manter = ['Ativo', 'Carteira (c/ Div)'] + [ind for ind in indexadores_sel]
            df_melt = df_comp[col_manter].melt(id_vars='Ativo', var_name='Indicador', value_name='Rentabilidade (%)')
            fig_comp = px.bar(df_melt, x='Ativo', y='Rentabilidade (%)', color='Indicador', barmode='group', color_discrete_map={'Carteira (c/ Div)': '#003f5c', 'CDI': '#00a896', 'IPCA': '#f4a261'}, title="Rentabilidade Tempo de Posse")
            st.plotly_chart(fig_comp, use_container_width=True)

    with t5:
        st.markdown("### 💸 Proventos Mensais e Status de Pagamento")
        c_f1, c_f2, c_btn = st.columns([2, 2, 2])
        meses_map = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
        m_hoje, a_hoje = pd.Timestamp.now().month, pd.Timestamp.now().year
        m_sel = c_f1.selectbox("Mês:", options=list(meses_map.keys()), format_func=lambda x: meses_map[x], index=m_hoje-1)
        a_sel = c_f2.selectbox("Ano:", options=[a_hoje, a_hoje-1, a_hoje-2])
        
        if c_btn.button("🔄 Processar Proventos", use_container_width=True):
            with st.spinner("Lendo histórico B3..."):
                la, lf = [], []
                for t_tk, dm in st.session_state.dados_mercado.items():
                    val = 0.0
                    try:
                        divs = yf.Ticker(f"{t_tk}.SA").dividends
                        if not divs.empty:
                            if divs.index.tz is not None: divs.index = divs.index.tz_localize(None)
                            val = float(divs[(divs.index.month == m_sel) & (divs.index.year == a_sel)].sum())
                    except: pass
                    
                    rec = val * dm['Qtd']
                    yoc = (rec / (dm['Qtd'] * dm['PM'])) * 100 if dm['PM']>0 else 0
                    
                    if dm['Tipo'] == 'FII':
                        lf.append({"FII": t_tk, "Unitário (R$)": val, "Recebido (R$)": rec, "Yield on Cost (%)": yoc, "Status": "Divulgado / Pago 🟢" if val>0 else "Aguardando 🟡"})
                    else:
                        if val > 0: la.append({"Ação": t_tk, "Unitário (R$)": val, "Recebido (R$)": rec, "Yield on Cost (%)": yoc, "Status": "Pago 🟢"})
                
                st.session_state.divs_acoes = pd.DataFrame(la)
                st.session_state.divs_fiis = pd.DataFrame(lf)
                st.session_state.divs_m = m_sel
        
        if 'divs_fiis' in st.session_state and not st.session_state.divs_fiis.empty:
            st.markdown("#### 🏢 Status dos Fundos Imobiliários (FIIs)")
            st.dataframe(st.session_state.divs_fiis.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct}), use_container_width=True, hide_index=True)
            
        if 'divs_acoes' in st.session_state and not st.session_state.divs_acoes.empty:
            st.markdown("#### 📈 Ações Pagadoras")
            st.dataframe(st.session_state.divs_acoes.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct}), use_container_width=True, hide_index=True)

else:
    for t in [t1, t2, t3, t4, t5]:
        with t: st.info("ℹ️ Para ativar as visões gráficas e proventos, adicione ativos ao lado e clique em **Conectar ao Mercado Vivo**.")

# ==========================================
# 8. ABAS INDEPENDENTES (SEMPRE ATIVAS)
# ==========================================
with t_tesouro:
    st.markdown("### 🏛️ Simulador e Controle de Tesouro Direto")
    st.info("Insira manualmente seus títulos de renda fixa abaixo. O sistema projetará o efeito composto até a data de vencimento de cada um.")
    
    if st.session_state.df_tesouro.empty:
        st.session_state.df_tesouro = pd.DataFrame([{"Título": "Tesouro IPCA+ 2029", "Investimento (R$)": 1000.0, "Taxa Anual (%)": 6.0, "Ano Venc.": 2029}])
        
    df_t = st.data_editor(st.session_state.df_tesouro, num_rows="dynamic", use_container_width=True, hide_index=True)
    st.session_state.df_tesouro = df_t
    
    if st.button("Projetar Títulos até Vencimento"):
        res_t = []
        for _, rt in df_t.iterrows():
            anos = max(1, int(rt['Ano Venc.']) - pd.Timestamp.now().year)
            v_final = float(rt['Investimento (R$)']) * ((1 + (float(rt['Taxa Anual (%)'])/100)) ** anos)
            res_t.append({"Título": rt['Título'], "Anos P/ Vencer": anos, "Investido": float(rt['Investimento (R$)']), "Valor Bruto no Vencimento": v_final, "Lucro Bruto Projetado": v_final - float(rt['Investimento (R$)'])})
        st.dataframe(pd.DataFrame(res_t).style.format({"Investido": f_brl, "Valor Bruto no Vencimento": f_brl, "Lucro Bruto Projetado": f_brl}), use_container_width=True, hide_index=True)

with t6:
    st.markdown("### 💬 Comitê de IA - Análise CNPI Avançada")
    c_btn1, c_btn2 = st.columns([1, 1])
    
    if c_btn1.button("🗑️ Limpar Histórico de Chat", use_container_width=True):
        st.session_state.historico_chat = [{"role": "assistant", "content": f"Saudações, {st.session_state.username}. O terminal foi reiniciado."}]
        st.rerun()
        
    if HAS_DOCX and len(st.session_state.historico_chat) > 1:
        c_btn2.download_button("📄 Exportar Conversa (Word)", data=export_docx(st.session_state.historico_chat), file_name=f"Relatorio_IA_{st.session_state.username}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    elif not HAS_DOCX:
        c_btn2.caption("⚠️ Adicione 'python-docx' no requirements.txt para habilitar o download em Word.")

    try: api_key = st.secrets.get("GEMINI_API_KEY", "")
    except: api_key = ""
        
    if not api_key:
        api_key = st.text_input("Insira sua Gemini API Key para ativar a Gestora IA:", type="password")
        
    for msg in st.session_state.historico_chat:
        with st.chat_message(msg["role"]): st.write(msg["content"])
        
    if prompt := st.chat_input("Ex: 'A ação BBAS3 está barata?' ou 'Avalie minha carteira'"):
        with st.chat_message("user"): st.write(prompt)
        st.session_state.historico_chat.append({"role": "user", "content": prompt})
        
        with st.chat_message("assistant"):
            with st.spinner("O Comitê de IA está processando..."):
                ctx_c = str(st.session_state.dados_mercado) if st.session_state.dados_mercado else "O usuário NÃO tem dados de carteira cadastrados no momento."
                ctx_m = f"Selic Vigente: {f_pct(selic_hoje)}|IPCA: {f_pct(ipca_12m_hoje)}. Focus {ano_atual}: Sel {f_pct(proj_focus.get(f'Selic_{ano_atual}'))}/IPCA {f_pct(proj_focus.get(f'IPCA_{ano_atual}'))}"
                historico_texto = "\n".join([f"{'Usuário' if m['role']=='user' else 'Gestora IA'}: {m['content']}" for m in st.session_state.historico_chat[-6:-1]])
                
                sys_prompt = (
                    f"Você é um Analista Sênior CNPI de alta performance. [Dados da Carteira]: {ctx_c}. [Macro]: {ctx_m}.\n"
                    f"REGRA ESTRITA DE ESCOPO:\n"
                    f"1) Se o usuário usar explicitamente palavras como 'minha carteira', 'meus ativos' ou perguntar de ativos que ele possui nos [Dados da Carteira], VOCÊ DEVE analisar o portfólio dele.\n"
                    f"2) Se o usuário NÃO citar a própria carteira e fizer uma pergunta genérica de mercado, IGNORE TOTALMENTE a carteira dele e responda de forma neutra e generalista.\n\n"
                    f"=== HISTÓRICO RECENTE (Para manter continuidade) ===\n{historico_texto}"
                )
                
                resposta = "⚠️ Chave API ausente ou não configurada."
                if api_key:
                    try:
                        import google.generativeai as genai
                        genai.configure(api_key=api_key)
                        resp_ok, ultimo_erro = False, ""
                        for m in ['gemini-2.5-flash', 'gemini-1.5-flash']:
                            try:
                                resposta = genai.GenerativeModel(m).generate_content([sys_prompt, prompt]).text
                                resp_ok = True
                                break 
                            except Exception as err:
                                ultimo_erro = str(err)
                                continue
                        if not resp_ok: resposta = f"⚠️ Falha de comunicação com IA. Erro: {ultimo_erro}"
                    except Exception as e: resposta = f"⚠️ Erro estrutural: {e}"
                
                st.write(resposta)
                
        st.session_state.historico_chat.append({"role": "assistant", "content": resposta})
        salvar_dados_usuario(st.session_state.username) # Salva histórico no DB logo após a resposta
        st.rerun()
