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

def f_brl(x): 
    return f"R$ {float(x):,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def f_brl_4(x): 
    return f"R$ {float(x):,.4f}".replace(",", "v").replace(".", ",").replace("v", ".")

def f_pct(x): 
    return f"{float(x):,.2f}%".replace(",", "v").replace(".", ",").replace("v", ".")

def to_excel(df, sheet_name='Sheet1'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'), 
            right=Side(style='thin', color='D9D9D9'), 
            top=Side(style='thin', color='D9D9D9'), 
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        for col_num, column in enumerate(worksheet.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            for row_num in range(2, worksheet.max_row + 1):
                data_cell = worksheet.cell(row=row_num, column=col_num)
                data_cell.font = Font(name="Arial", size=10)
                data_cell.border = thin_border
                
                if isinstance(data_cell.value, (int, float)): 
                    data_cell.alignment = Alignment(horizontal="right")
                else: 
                    data_cell.alignment = Alignment(horizontal="center")
                    
            max_len = max(len(str(c.value or '')) for c in column)
            col_letter = get_column_letter(col_num)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 13)
            
    return output.getvalue()

def export_docx(historico):
    if not HAS_DOCX: 
        return None
        
    doc = docx.Document()
    doc.add_heading('Relatório Analítico Institucional - Comitê de IA CNPI', 0)
    
    for msg in historico:
        if msg["role"] == "user":
            doc.add_heading("Consulta Operacional:", level=2)
        else:
            doc.add_heading("Parecer da Gestora IA:", level=2)
            
        doc.add_paragraph(msg["content"])
        
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

MAPEAMENTO_TICKERS = {
    "GALG11": "GARE11", "SOMA3": "ALOS3", "ARZZ3": "ALOS3", 
    "VVAR3": "BHIA3", "VIIA3": "BHIA3", "BRML3": "ALSO3", 
    "BBRK11": "BRCR11", "HCTR11": "TRXD11", "TORD11": "TRXD11"
}

UNITS_ACOES = [
    'SANB11', 'TAEE11', 'KLBN11', 'BPAC11', 'ALUP11', 
    'ENGI11', 'BIDI11', 'CPLE11', 'SAPR11', 'RNEW11'
]

# Inicialização de Variáveis de Sessão para garantir a persistência na navegação
if 'df_base' not in st.session_state: 
    st.session_state.df_base = pd.DataFrame()
    
if 'df_tesouro' not in st.session_state: 
    st.session_state.df_tesouro = pd.DataFrame()
    
if 'dados_mercado' not in st.session_state: 
    st.session_state.dados_mercado = {}
    
if 'df_simul' not in st.session_state: 
    st.session_state.df_simul = pd.DataFrame()
    
if 'logged_in' not in st.session_state: 
    st.session_state.logged_in = False
    
if 'username' not in st.session_state: 
    st.session_state.username = ""
    
if 'historico_chat' not in st.session_state: 
    st.session_state.historico_chat = []

# ==========================================
# 2. MOTOR DE BANCO DE DADOS HÍBRIDO (SQLite / PostgreSQL)
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
        c.execute('''CREATE TABLE IF NOT EXISTS tesouro_v2 (username TEXT, titulo TEXT, data_compra TEXT, tipo_taxa TEXT, investido DOUBLE PRECISION, taxa DOUBLE PRECISION, vencimento INTEGER)''')
    else:
        c.execute('''CREATE TABLE IF NOT EXISTS carteiras (username TEXT, Ativo TEXT, Quantidade REAL, Preco_Medio REAL, Data_Media TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS tesouro_v2 (username TEXT, titulo TEXT, data_compra TEXT, tipo_taxa TEXT, investido REAL, taxa REAL, vencimento INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS chat_ia (username TEXT, role TEXT, content TEXT)''')
    conn.commit()
    conn.close()

def salvar_dados_completos_db(username):
    conn = get_db_connection(); c = conn.cursor()
    usr = username.strip()
    
    # Salvar Carteira B3
    c.execute(f"DELETE FROM carteiras WHERE username={PARAM}", (usr,))
    if not st.session_state.df_base.empty:
        for _, r in st.session_state.df_base.iterrows():
            c.execute(f"INSERT INTO carteiras (username, ativo, quantidade, preco_medio, data_media) VALUES ({PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM})",
                      (usr, str(r.get('Ativo', '')), float(limpar_numero(r.get('Quantidade', 0))), float(limpar_numero(r.get('Preço Médio', 0))), str(r.get('Data Média', ''))))
            
    # Salvar Tesouro Direto V2 de Forma Segura contra Nulos
    c.execute(f"DELETE FROM tesouro_v2 WHERE username={PARAM}", (usr,))
    if isinstance(st.session_state.df_tesouro, pd.DataFrame) and not st.session_state.df_tesouro.empty:
        for _, r in st.session_state.df_tesouro.iterrows():
            if pd.isna(r.get('Título')) or str(r.get('Título')).strip() == '':
                continue
            titulo = str(r.get('Título', 'Tesouro'))
            dt_compra = str(r.get('Data Compra', pd.Timestamp.now().date()))
            tipo_taxa = str(r.get('Tipo Taxa', 'Pré-fixado'))
            investido = float(limpar_numero(r.get('Valor Investido (R$)', 0)))
            taxa = float(limpar_numero(r.get('Taxa Contratada (%)', 0)))
            vencimento = int(limpar_numero(r.get('Ano Vencimento', pd.Timestamp.now().year + 5)))
            
            c.execute(f"INSERT INTO tesouro_v2 (username, titulo, data_compra, tipo_taxa, investido, taxa, vencimento) VALUES ({PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM})",
                      (usr, titulo, dt_compra, tipo_taxa, investido, taxa, vencimento))
            
    # Salvar Histórico do Comitê de IA
    c.execute(f"DELETE FROM chat_ia WHERE username={PARAM}", (usr,))
    for msg in st.session_state.historico_chat[-30:]:
        c.execute(f"INSERT INTO chat_ia (username, role, content) VALUES ({PARAM}, {PARAM}, {PARAM})", (usr, msg['role'], msg['content']))
        
    conn.commit(); conn.close()

def carregar_dados_completos_db(username):
    conn = get_db_connection()
    usr = username.strip()
    
    # Carregar Carteira B3
    query_cart = f"SELECT Ativo, Quantidade, Preco_Medio, Data_Media FROM carteiras WHERE username={PARAM}"
    df_cart = pd.read_sql_query(query_cart, conn, params=(usr,))
    df_cart = df_cart.rename(columns={"Preco_Medio": "Preço Médio", "preco_medio": "Preço Médio", "Data_Media": "Data Média", "data_media": "Data Média", "ativo": "Ativo", "quantidade": "Quantidade"})
    if not df_cart.empty: df_cart['Data Média'] = pd.to_datetime(df_cart['Data Média']).dt.date
    else: df_cart = pd.DataFrame(columns=["Ativo", "Quantidade", "Preço Médio", "Data Média"])
    st.session_state.df_base = df_cart
    
    # Carregar Tesouro Direto V2 com as colunas corretas solicitadas
    query_tes = f"SELECT titulo, data_compra, tipo_taxa, investido, taxa, vencimento FROM tesouro_v2 WHERE username={PARAM}"
    df_tes = pd.read_sql_query(query_tes, conn, params=(usr,))
    df_tes = df_tes.rename(columns={"titulo": "Título", "data_compra": "Data Compra", "tipo_taxa": "Tipo Taxa", "investido": "Valor Investido (R$)", "taxa": "Taxa Contratada (%)", "vencimento": "Ano Vencimento"})
    if not df_tes.empty: df_tes['Data Compra'] = pd.to_datetime(df_tes['Data Compra']).dt.date
    else: df_tes = pd.DataFrame(columns=["Título", "Data Compra", "Tipo Taxa", "Valor Investido (R$)", "Taxa Contratada (%)", "Ano Vencimento"])
    st.session_state.df_tesouro = df_tes
    
    # Carregar Chat IA
    query_chat = f"SELECT role, content FROM chat_ia WHERE username={PARAM}"
    df_chat = pd.read_sql_query(query_chat, conn, params=(usr,))
    if not df_chat.empty: st.session_state.historico_chat = df_chat.to_dict('records')
    else: st.session_state.historico_chat = [{"role": "assistant", "content": f"Saudações, {username}. O terminal está mapeado e online. Como posso ajudar?"}]
        
    conn.close()

init_db()

# ==========================================
# 3. GATEKEEPER - TELA DE AUTENTICAÇÃO
# ==========================================
if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center;'>🔐 Terminal de Gestão Profissional</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Acesso restrito. Identifique-se para carregar seu portfólio.</p>", unsafe_allow_html=True)
    
    col_log1, col_log2, col_log3 = st.columns([1, 1, 1])
    
    with col_log2:
        tab_login, tab_register, tab_forgot = st.tabs(["Acesso", "Novo Registro", "Recuperar Senha"])
        
        with tab_login:
            login_user = st.text_input("Usuário", key="log_user").strip()
            login_pass = st.text_input("Senha", type="password", key="log_pass")
            
            if st.button("Entrar no Terminal", use_container_width=True):
                if autenticar_usuario(login_user, login_pass):
                    st.session_state.logged_in = True
                    st.session_state.username = login_user
                    carregar_dados_completos_db(login_user)
                    st.rerun()
                else: 
                    st.error("Credenciais inválidas.")
                
        with tab_register:
            reg_user = st.text_input("Novo Usuário", key="reg_user").strip()
            reg_pass = st.text_input("Nova Senha", type="password", key="reg_pass")
            
            if st.button("Registrar Acesso", use_container_width=True):
                if reg_user and reg_pass:
                    if registrar_usuario(reg_user, reg_pass): 
                        st.success("Conta provisionada com sucesso! Você já pode fazer o login.")
                    else: 
                        st.error("O nome de usuário já está em uso no sistema.")
                else: 
                    st.warning("Preencha ambos os campos corretamente.")
                
        with tab_forgot:
            forgot_user = st.text_input("Usuário Cadastrado", key="for_user").strip()
            forgot_pass = st.text_input("Nova Senha Desejada", type="password", key="for_pass")
            
            if st.button("Redefinir Senha", use_container_width=True):
                if atualizar_senha(forgot_user, forgot_pass): 
                    st.success("Sua senha foi redefinida com sucesso.")
                else: 
                    st.error("Usuário não encontrado no banco de dados.")
                    
    st.stop()
    # ==========================================
# 4. FUNÇÕES DE PROCESSAMENTO B3 E MACRO
# ==========================================
st.markdown(f"""
    <div style="text-align: center; margin-bottom: 20px;">
        <h3 style="font-weight: 400; margin-bottom: 0;">💼 Terminal de Gestão</h3>
        <h6 style="color: #666; font-weight: 300;">Analista Sênior: {st.session_state.username.upper()}</h6>
    </div>
""", unsafe_allow_html=True)
st.write("---")

@st.cache_data(ttl=86400)
def obter_macro_atual():
    selic_atual = 10.50
    ipca_12m = 4.00
    
    try:
        res = requests.get("https://brasilapi.com.br/api/taxas/v1", timeout=5)
        if res.status_code == 200:
            for taxa in res.json():
                if taxa['nome'] == 'Selic': 
                    selic_atual = float(taxa['valor'])
    except:
        try:
            selic_df = sgs.get({'selic': 432}, last=1)
            if not selic_df.empty: 
                selic_atual = float(selic_df['selic'].iloc[-1])
        except: 
            pass
            
    try:
        ipca_df = sgs.get({'IPCA_12M': 13522}, last=1)
        if not ipca_df.empty: 
            ipca_12m = float(ipca_df['IPCA_12M'].iloc[-1])
    except: 
        pass
        
    return selic_atual, ipca_12m

@st.cache_data(ttl=86400)
def carregar_dados_mercado():
    macro = pd.DataFrame()
    fundamentos = {}
    
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'] = macro['CDI'] / 100
        macro['IPCA'] = macro['IPCA'] / 100
        
        url_fundamentus = 'https://www.fundamentus.com.br/resultado.php'
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        df_fund = pd.read_html(io.StringIO(requests.get(url_fundamentus, headers=headers, timeout=10).text), decimal=',', thousands='.')[0]
        
        for _, row in df_fund.iterrows():
            t = str(row['Papel']).strip().upper()
            c = float(row['Cotação'])
            pl = float(row['P/L'])
            pvp = float(row['P/VP'])
            fundamentos[t] = {
                'vpa': c/pvp if pvp > 0 else 0.0, 
                'lpa': c/pl if pl > 0 else 0.0
            }
    except Exception as e: 
        pass
    
    ano_at = pd.Timestamp.now().year
    proj = {}
    
    try:
        url_focus = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$top=300&$filter=Indicador%20eq%20'IPCA'%20or%20Indicador%20eq%20'Selic'&$orderby=Data%20desc&$format=json"
        df_p = pd.DataFrame(requests.get(url_focus, timeout=5).json().get('value', []))
        
        if not df_p.empty:
            df_p = df_p[df_p['Data'] == df_p['Data'].max()]
            for ao in [0, 1, 2]:
                aa = str(ano_at + ao)
                df_a = df_p[df_p['DataReferencia'] == aa]
                if not df_a[df_a['Indicador'] == 'IPCA'].empty: 
                    proj[f"IPCA_{aa}"] = float(df_a[df_a['Indicador'] == 'IPCA']['Mediana'].values[0])
                if not df_a[df_a['Indicador'] == 'Selic'].empty: 
                    proj[f"Selic_{aa}"] = float(df_a[df_a['Indicador'] == 'Selic']['Mediana'].values[0])
    except: 
        pass
        
    return macro, fundamentos, proj, ano_at

selic_hoje, ipca_12m_hoje = obter_macro_atual()
df_macro, fundamentos_br, proj_focus, ano_atual = carregar_dados_mercado()

def limpar_numero(x):
    if pd.isna(x): 
        return 0.0
    if isinstance(x, (int, float, np.number)): 
        return float(x)
    try: 
        return float(str(x).replace('R$', '').replace('.', '').replace(',', '.').strip())
    except: 
        return 0.0

def traduzir_setor(setor_en):
    dicionario = {
        "Banks": "Bancos", "Utilities - Regulated Electric": "Energia", 
        "Real Estate - Retail": "Shoppings/Varejo", "REIT - Retail": "Shoppings/Varejo", 
        "Real Estate - Industrial": "Logística", "REIT - Industrial": "Logística", 
        "REIT - Office": "Lajes Corporativas", "REIT - Diversified": "Fundo Híbrido", 
        "Financial Data & Stock Exchanges": "Bolsa de Valores", 
        "Insurance": "Seguradoras", "Oil & Gas Integrated": "Petróleo e Gás"
    }
    return dicionario.get(setor_en, "Outros Setores")

def consolidar_carteira(df):
    if df.empty: 
        return df
        
    df['Ativo'] = df['Ativo'].astype(str).str.strip().str.upper().apply(lambda x: MAPEAMENTO_TICKERS.get(x, x))
    linhas = []
    
    for ativo, group in df.groupby('Ativo'):
        qtd = float(group['Quantidade'].sum())
        if qtd > 0:
            pm = (group['Quantidade'] * group['Preço Médio']).sum() / qtd
            ts = sum((pd.Timestamp(r['Data Média']).timestamp() * r['Quantidade']) for _, r in group.iterrows() if pd.notna(r['Data Média']))
            data_media = pd.to_datetime(ts/qtd, unit='s').date()
            linhas.append({
                "Ativo": ativo, 
                "Quantidade": qtd, 
                "Preço Médio": float(pm), 
                "Data Média": data_media
            })
            
    return pd.DataFrame(linhas)

def corrigir_cabecalho_b3(df):
    if df.empty: 
        return df
    if 'Data do Negócio' in df.columns or 'Data Média' in df.columns: 
        return df
        
    def tentar_mapear(lista_colunas):
        lista_upper = [str(c).strip().upper() for c in lista_colunas]
        is_negociacao = any('DATA DO' in c or 'NEGÓCIO' in c or 'NEGOCIO' in c for c in lista_upper)
        is_mov = any('PRODUTO' in c or 'ATIVO' in c for c in lista_upper) and any('MOVIMENTA' in c or 'MOVIMEN' in c for c in lista_upper)
        
        if is_negociacao or is_mov:
            mapeamento = {}
            for col in lista_colunas:
                c_up = str(col).strip().upper()
                if 'DAT' in c_up: mapeamento[col] = 'Data do Negócio'
                elif 'PROD' in c_up or 'ATIV' in c_up or 'CÓD' in c_up or 'COD' in c_up or 'PAPEL' in c_up: mapeamento[col] = 'Código de Negociação'
                elif 'MOV' in c_up or 'TIP' in c_up: mapeamento[col] = 'Tipo de Movimentação'
                elif 'VAL' in c_up: mapeamento[col] = 'Valor'
                elif 'QUA' in c_up or 'QTD' in c_up: mapeamento[col] = 'Quantidade'
                elif 'PRE' in c_up or 'AJUSTE' in c_up: mapeamento[col] = 'Preço Unitário'
                elif 'ENTRADA' in c_up or 'SAÍDA' in c_up or 'SAIDA' in c_up or 'C/D' in c_up: mapeamento[col] = 'Entrada/Saída'
            return mapeamento, is_mov
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
            df_sub = df.iloc[i+1:].reset_index(drop=True).rename(columns=mapeamento)
            if is_mov and 'Valor' not in df_sub.columns and 'Preço Unitário' in df_sub.columns and 'Quantidade' in df_sub.columns:
                df_sub['Valor'] = df_sub['Quantidade'].apply(limpar_numero) * df_sub['Preço Unitário'].apply(limpar_numero)
            return df_sub
            
    return df

def processar_planilha_b3(df):
    if df.empty: 
        return pd.DataFrame()
        
    df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
    df['Quantidade'] = df['Quantidade'].apply(limpar_numero)
    
    if 'Valor' in df.columns: 
        df['Valor'] = df['Valor'].apply(limpar_numero)
    elif 'Preço Unitário' in df.columns: 
        df['Valor'] = df['Quantidade'] * df['Preço Unitário'].apply(limpar_numero)
    else: 
        df['Valor'] = 0.0

    df = df.sort_values('Data do Negócio')
    posicoes = {}
    
    for _, row in df.iterrows():
        if pd.isna(row.get('Código de Negociação')): 
            continue
            
        ticker_raw = str(row['Código de Negociação']).strip().upper()
        if " - " in ticker_raw: 
            ticker_raw = ticker_raw.split(" - ")[0].strip()
        elif " " in ticker_raw: 
            ticker_raw = ticker_raw.split(" ")[0].strip()
            
        ticker = MAPEAMENTO_TICKERS.get(ticker_raw[:-1] if ticker_raw.endswith('F') and len(ticker_raw) > 4 else ticker_raw, ticker_raw[:-1] if ticker_raw.endswith('F') and len(ticker_raw) > 4 else ticker_raw)
        
        if not any(c.isdigit() for c in ticker) or len(ticker) < 4: 
            continue
            
        if 'Tipo de Movimentação' in df.columns:
            mov_tipo_str = str(row['Tipo de Movimentação']).strip().upper()
            if any(p in mov_tipo_str for p in ['RENDIMENTO', 'JUROS', 'DIVIDENDO', 'JCP', 'REEMBOLSO', 'BONIFICAÇÃO', 'DESDOBRAMENTO', 'GRUPAMENTO', 'FRAÇÃO']):
                continue
        
        qtd = row['Quantidade']
        valor = row['Valor']
        data = row['Data do Negócio'] if pd.notna(row['Data do Negócio']) else pd.Timestamp.now()
        
        if ticker not in posicoes: 
            posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'ts_medio': 0.0}
        
        is_compra = False
        is_venda = False
        
        if 'Entrada/Saída' in df.columns and pd.notna(row['Entrada/Saída']):
            io_dir = str(row['Entrada/Saída']).strip().upper()
            if 'CRED' in io_dir or 'ENT' in io_dir or 'C' == io_dir: 
                is_compra = True
            elif 'DEB' in io_dir or 'SAI' in io_dir or 'D' == io_dir: 
                is_venda = True
        
        if not is_compra and not is_venda and 'Tipo de Movimentação' in df.columns:
            tipo_mov = str(row['Tipo de Movimentação']).strip().upper()
            if any(term in tipo_mov for term in ['COMPRA', 'CREDITO', 'CRÉDITO', 'TRANSFERÊNCIA', 'TRANSFERENCIA']) or tipo_mov == 'C':
                is_compra = True
            elif any(term in tipo_mov for term in ['VENDA', 'DEBITO', 'DÉBITO']) or tipo_mov == 'V':
                is_venda = True
        
        if is_compra:
            q_ant = posicoes[ticker]['qtd']
            ts_ant = posicoes[ticker]['ts_medio']
            ts_novo = pd.Timestamp(data).timestamp()
            
            posicoes[ticker]['ts_medio'] = ts_novo if q_ant == 0 else ((ts_ant * q_ant) + (ts_novo * qtd)) / (q_ant + qtd)
            posicoes[ticker]['qtd'] += qtd
            posicoes[ticker]['valor'] += valor
            
        elif is_venda:
            if qtd >= (posicoes[ticker]['qtd'] - 0.001): 
                posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'ts_medio': 0.0}
            else:
                pm = posicoes[ticker]['valor'] / posicoes[ticker]['qtd'] if posicoes[ticker]['qtd'] > 0 else 0
                posicoes[ticker]['qtd'] -= qtd
                posicoes[ticker]['valor'] -= (qtd * pm)
                
    ativos_finais = []
    for k, v in posicoes.items():
        if v['qtd'] > 0:
            ativos_finais.append({
                "Ativo": k, 
                "Quantidade": v['qtd'], 
                "Preço Médio": v['valor']/v['qtd'], 
                "Data Média": pd.to_datetime(v['ts_medio'], unit='s').date()
            })
            
    return consolidar_carteira(pd.DataFrame(ativos_finais))

# ==========================================
# 5. SIDEBAR: UPLOAD, INTEGRAÇÃO B3 E BACKUP
# ==========================================
with st.sidebar:
    st.markdown("### 👤 ANALISTA OPERACIONAL")
    if st.button("🚪 Sair do Terminal", use_container_width=True):
        st.session_state.clear()
        st.rerun()
        
    st.divider()
    
    st.markdown("### 💾 Segurança em Nuvem")
    if st.button("Sincronizar no Banco de Dados", type="primary", use_container_width=True):
        salvar_dados_completos_db(st.session_state.username)
        st.success("Dados blindados no banco com sucesso!")
        
    st.divider()
    
    st.markdown("### 1. Integrar Notas B3 (Opcional)")
    st.info("Faça o upload de planilhas para substituir toda a base ou apenas integrar novas operações à carteira atual.")
    
    arquivo_principal = st.file_uploader("Substituir Base Completa (B3 / CSV Backup)", type=["xlsx", "csv"])
    arquivo_novo = st.file_uploader("Integrar Novas Operações", type=["xlsx", "csv"])
    
    if arquivo_novo:
        data_corte = st.date_input("Filtrar Novas a partir de:", pd.Timestamp.now().date() - pd.Timedelta(days=15))
    else:
        data_corte = None

    if st.button("🚀 Processar Arquivos B3", use_container_width=True):
        base_atual = st.session_state.df_base.copy()
        
        if arquivo_principal:
            if arquivo_principal.name.endswith('.csv'):
                txt = arquivo_principal.getvalue().decode('utf-8-sig', errors='ignore')
                sep_detectado = '\t' if txt.count('\t') > txt.count(';') else ';'
                df_p = pd.read_csv(io.StringIO(txt), sep=sep_detectado)
            else:
                df_p = pd.read_excel(arquivo_principal)
                
            df_p = corrigir_cabecalho_b3(df_p)
            
            if 'Data Média' in df_p.columns: 
                base_atual = consolidar_carteira(df_p)
            elif 'Data do Negócio' in df_p.columns: 
                base_atual = processar_planilha_b3(df_p)
            else: 
                st.error("Formato inválido. Planilha B3 não reconhecida.")
                st.stop()
                
        if arquivo_novo and not base_atual.empty:
            if arquivo_novo.name.endswith('.csv'):
                txt_n = arquivo_novo.getvalue().decode('utf-8-sig', errors='ignore')
                sep_detectado_n = '\t' if txt_n.count('\t') > txt_n.count(';') else ';'
                df_n = pd.read_csv(io.StringIO(txt_n), sep=sep_detectado_n)
            else:
                df_n = pd.read_excel(arquivo_novo)
                
            df_n = corrigir_cabecalho_b3(df_n)
            
            if not df_n.empty and 'Data do Negócio' in df_n.columns:
                df_n['Data do Negócio'] = pd.to_datetime(df_n['Data do Negócio'], dayfirst=True, errors='coerce')
                df_n = df_n[df_n['Data do Negócio'].dt.date >= data_corte]
                
                linhas_b = []
                for _, r in base_atual.iterrows():
                    linhas_b.append({
                        "Código de Negociação": r['Ativo'], 
                        "Tipo de Movimentação": "Compra", 
                        "Data do Negócio": pd.to_datetime(r['Data Média']), 
                        "Quantidade": r['Quantidade'], 
                        "Valor": r['Quantidade'] * r['Preço Médio']
                    })
                    
                base_atual = processar_planilha_b3(pd.concat([pd.DataFrame(linhas_b), df_n], ignore_index=True))
                
        st.session_state.df_base = base_atual
        st.warning("Memória atualizada. Lembre-se de salvar no DB.")
        st.rerun()

# ==========================================
# 6. PAINEL MACRO E CONTROLE MANUAL DE ATIVOS
# ==========================================
st.markdown("### 👑 Conjuntura Macroeconômica")
c_m1, c_m2 = st.columns([1, 2])

c_m1.success(
    f"🎯 **Cenário Atual (Vigente)**\n\n"
    f"Selic Vigente: **{f_pct(selic_hoje)} a.a.**\n\n"
    f"IPCA 12 meses: **{f_pct(ipca_12m_hoje)}**"
)

c_m2.info(
    f"🔮 **Projeções do Mercado (Boletim Focus)**\n\n"
    f"**Taxa Selic:** {ano_atual}: **{f_pct(proj_focus.get(f'Selic_{ano_atual}', 0))}** |  "
    f"{ano_atual+1}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+1}', 0))}** |  "
    f"{ano_atual+2}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+2}', 0))}**\n\n"
    f"**Inflação (IPCA):** {ano_atual}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual}', 0))}** |  "
    f"{ano_atual+1}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+1}', 0))}** |  "
    f"{ano_atual+2}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+2}', 0))}**"
)

st.write("---")
st.markdown("### 2. Controle Operacional da Carteira")

c_op1, c_op2, c_op3 = st.columns([1, 1.5, 1])

with c_op1:
    ativos_lista = sorted(st.session_state.df_base["Ativo"].tolist()) if not st.session_state.df_base.empty else []
    tdel = st.selectbox("Excluir Ativo da Memória:", [""] + ativos_lista)
    if st.button("Remover Ativo Selecionado", use_container_width=True) and tdel:
        st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != tdel]
        st.rerun()

with c_op2:
    nt = st.text_input("Novo Aporte ou Ativo Manual (Ticker)")
    cq, cp = st.columns(2)
    nq = cq.number_input("Quantidade", min_value=1)
    np_v = cp.number_input("Preço Médio (R$)", min_value=0.01)
    
    if st.button("Adicionar à Carteira / Integrar", use_container_width=True) and nt:
        nova_linha = pd.DataFrame([{
            "Ativo": nt.upper(), 
            "Quantidade": float(nq), 
            "Preço Médio": float(np_v), 
            "Data Média": pd.Timestamp.now().date()
        }])
        st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, nova_linha], ignore_index=True))
        st.rerun()

with c_op3:
    if not st.session_state.df_base.empty:
        csv_data = st.session_state.df_base.to_csv(index=False, sep=';', encoding='utf-8-sig')
        st.download_button(
            label="📥 Baixar Carteira Ajustada (CSV Backup)", 
            data=csv_data, 
            file_name=f"Carteira_Ajustada_{st.session_state.username}.csv", 
            use_container_width=True
        )
    st.write("")
    if st.button("💾 Salvar no Banco de Dados", type="primary", use_container_width=True):
        salvar_dados_completos_db(st.session_state.username)
        st.success("Dados blindados no banco com sucesso!")

st.markdown("#### 📝 Tabela Editável (Ajuste Fino de Quantidade e Preço Médio)")
st.info("Ajuste quantidades ou preços médios diretamente nas células da tabela abaixo antes de conectar aos servidores da B3.")

df_editado = st.data_editor(st.session_state.df_base, use_container_width=True, hide_index=True)

if st.button("🚀 Conectar ao Mercado Vivo", type="primary", use_container_width=True):
    if not df_editado.empty:
        st.session_state.df_base = consolidar_carteira(df_editado)
        
        prg = st.progress(0)
        total = len(st.session_state.df_base)
        dm = {}
        sim = []
        
        for i, r in st.session_state.df_base.iterrows():
            t = str(r['Ativo']).upper()
            dc = pd.to_datetime(r['Data Média']) if pd.notna(r['Data Média']) else pd.Timestamp.now()
            
            p_at = float(r['Preço Médio'])
            d_tot = 0.0
            d_12m = 0.0
            is_fii = t.endswith('11') and t not in UNITS_ACOES
            setor = "Desconhecido"
            
            try:
                tk = yf.Ticker(f"{t}.SA")
                h = tk.history(period="1d")
                if not h.empty: 
                    p_at = float(h['Close'].iloc[-1])
                    
                dvs = tk.dividends
                if not dvs.empty:
                    if dvs.index.tz is not None: 
                        dvs.index = dvs.index.tz_localize(None)
                    d_tot = float(dvs[dvs.index >= dc].sum() * r['Quantidade'])
                    d_12m = float(dvs[dvs.index >= (pd.Timestamp.now() - pd.DateOffset(years=1))].sum())
                    
                setor = traduzir_setor(tk.info.get('industry', ''))
            except: 
                pass
            
            cdi_ac, ipca_ac = 0.0, 0.0
            try: 
                if not df_macro.empty:
                    f_m = df_macro.loc[dc:]
                    cdi_ac = ((1 + f_m['CDI'].dropna()).prod() - 1) * 100
                    ipca_ac = ((1 + f_m['IPCA'].dropna()).prod() - 1) * 100
            except: 
                pass
            
            vpa = fundamentos_br.get(t, {}).get('vpa', 0)
            lpa = fundamentos_br.get(t, {}).get('lpa', 0)
            
            dm[t] = {
                "Qtd": float(r['Quantidade']), 
                "PM": float(r['Preço Médio']), 
                "Data": dc, 
                "Preço Atual": p_at, 
                "Div_Total": d_tot, 
                "CDI": cdi_ac, 
                "IPCA": ipca_ac, 
                "Tipo": "FII" if is_fii else "Ação", 
                "Setor": setor
            }
            
            sim.append({
                "Ativo": t, 
                "Cotação Atual": p_at, 
                "VPA (Contábil)": vpa, 
                "LPA Projetado": lpa, 
                "Div. Projetado (R$)": d_12m
            })
            
            prg.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dm
        st.session_state.df_simul = pd.DataFrame(sim)
        st.success("Mercado Sincronizado com a B3! As abas de análise estão prontas e atualizadas.")
    else: 
        st.warning("A carteira está vazia. Adicione ativos antes de conectar.")

st.write("---")
# ==========================================
# 7. DASHBOARDS E RELATÓRIOS (TABS DE ANÁLISE)
# ==========================================
tab_visao, tab_val, tab_radar, tab_graf, tab_prov, tab_tesouro, tab_ia = st.tabs([
    "📊 Visão Geral", 
    "💰 Valuation", 
    "🎯 Radar & Projeção", 
    "📈 Gráficos", 
    "💸 Proventos B3", 
    "🏛️ Tesouro Direto", 
    "💬 Gestora IA (CNPI)"
])

if st.session_state.dados_mercado:
    l_pf = []
    for t, dm in st.session_state.dados_mercado.items():
        investido = dm['Qtd'] * dm['PM']
        saldo = dm['Qtd'] * dm['Preço Atual']
        div_acumulado = dm['Div_Total']
        
        l_pf.append({
            "Ativo": t, 
            "Tipo": dm["Tipo"], 
            "Setor": dm.get("Setor", "Desconhecido"), 
            "Qtd": int(dm['Qtd']), 
            "Preço Médio": dm['PM'], 
            "Preço Atual": dm['Preço Atual'],
            "Total Investido": investido, 
            "Saldo Atual": saldo, 
            "Saldo C/ Dividendos": saldo + div_acumulado,
            "Resultado (R$)": saldo - investido, 
            "Resultado C/ Dividendos": (saldo - investido) + div_acumulado,
            "Data Média": dm['Data'].strftime('%d/%m/%Y'), 
            "Total Div. (R$)": div_acumulado, 
            "DY on Cost (%)": (div_acumulado / investido) * 100 if investido > 0 else 0, 
            "Evolução c/ Div (%)": (((saldo + div_acumulado) / investido) - 1) * 100 if investido > 0 else 0,
            "IPCA Acum. (%)": dm['IPCA'], 
            "CDI Acum. (%)": dm['CDI']
        })
        
    df_perf_final = pd.DataFrame(l_pf)

    with tab_visao:
        st.markdown("### 🏆 Visão Global da Carteira")
        
        df_a = df_perf_final[df_perf_final['Tipo'] == 'Ação']
        df_f = df_perf_final[df_perf_final['Tipo'] == 'FII']
        
        ev_a = (df_a['Saldo Atual'].sum() / df_a['Total Investido'].sum() - 1) * 100 if df_a['Total Investido'].sum() > 0 else 0
        ev_f = (df_f['Saldo Atual'].sum() / df_f['Total Investido'].sum() - 1) * 100 if df_f['Total Investido'].sum() > 0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📈 Patrimônio Ações", f_brl(df_a['Saldo Atual'].sum()), f_pct(ev_a))
        m2.metric("🏢 Patrimônio FIIs", f_brl(df_f['Saldo Atual'].sum()), f_pct(ev_f))
        m3.metric("💸 Renda Histórica Ações", f_brl(df_a['Total Div. (R$)'].sum()))
        m4.metric("💸 Renda Histórica FIIs", f_brl(df_f['Total Div. (R$)'].sum()))

        formatacao_t1 = {
            c: f_brl for c in ["Preço Médio", "Preço Atual", "Total Investido", "Saldo Atual", "Saldo C/ Dividendos", "Resultado (R$)", "Resultado C/ Dividendos", "Total Div. (R$)"]
        }
        formatacao_t1.update({
            c: f_pct for c in ["DY on Cost (%)", "Evolução c/ Div (%)", "IPCA Acum. (%)", "CDI Acum. (%)"]
        })
        
        st.dataframe(df_perf_final.drop(columns=['Tipo', 'Setor']).style.format(formatacao_t1), use_container_width=True, hide_index=True)

    with tab_val:
        st.markdown("#### Métodos Certificados de Valuation")
        st.markdown("""
        * **Preço Teto Decio Bazin:** Calcula o preço máximo ideal para compra focado em retornos via dividendos.
        * **Preço Justo Benjamin Graham:** Avalia o valor intrínseco e contábil. Proteção de patrimônio. (Não aplicável a FIIs).
        """)
        
        yd = st.number_input("Taxa de Retorno Mínima Exigida Bazin (%):", value=6.0, step=0.5) / 100.0
        
        df_sim = st.data_editor(st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]], use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"])
        st.session_state.df_simul[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]] = df_sim[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]]
        
        rv = []
        for _, r in df_sim.iterrows():
            is_fii = r['Ativo'].endswith('11') and r['Ativo'] not in UNITS_ACOES
            
            bz = (float(r["Div. Projetado (R$)"]) / yd) if float(r["Div. Projetado (R$)"]) > 0 else 0.0
            mbz = ((bz / float(r["Cotação Atual"])) - 1) * 100 if bz > 0 else 0.0
            
            if not is_fii:
                gh = (22.5 * float(r["LPA Projetado"]) * float(r["VPA (Contábil)"])) ** 0.5 if float(r["LPA Projetado"]) > 0 and float(r["VPA (Contábil)"]) > 0 else 0.0
                mgh = ((gh / float(r["Cotação Atual"])) - 1) * 100 if gh > 0 else 0.0
            else:
                gh = np.nan
                mgh = np.nan
                
            rv.append({"Ativo": r['Ativo'], "Teto Bazin": bz, "Margem Bazin (%)": mbz, "Justo Graham": gh, "Margem Graham (%)": mgh})
            
        st.session_state.df_recs_val = pd.DataFrame(rv)
        
        formatacao_t2 = {
            "Teto Bazin": lambda x: f_brl(x) if x > 0 else "-", 
            "Justo Graham": lambda x: f_brl(x) if pd.notna(x) and x > 0 else "-", 
            "Margem Bazin (%)": lambda x: f_pct(x) if x != 0 else "-", 
            "Margem Graham (%)": lambda x: f_pct(x) if pd.notna(x) and x != 0 else "-"
        }
        st.dataframe(st.session_state.df_recs_val.style.format(formatacao_t2), use_container_width=True, hide_index=True)

    with tab_radar: 
        st.markdown("##### Parametrização do Radar Operacional")
        c_p1, c_p2, c_p3, c_p4 = st.columns(4)
        patr_fora = c_p1.number_input("Patrimônio Externo / Caixa (R$):", value=0.0, step=1000.0)
        aporte = c_p2.number_input("Aporte Mensal Previsto (R$):", value=2000.0, step=500.0)
        rent = c_p3.number_input("Rentabilidade Mensal Alvo (%):", value=0.8, step=0.1) / 100.0
        cresc_div = c_p4.number_input("Crescimento Anual de Dividendos (%):", value=5.0, step=1.0) / 100.0

        st.markdown("##### 🎯 Triagem Estratégica Corporativa")
        st.info(
            "**Teoria da Margem de Segurança:**\n\n"
            "**Graham (>15% a 20%):** Protege contra quedas de mercado sistêmicas e balanços irreais. Você só entra se a ação estiver muito barata frente ao patrimônio físico dela.\n"
            "**Bazin (>0% a 5%):** Garante que, independentemente do sobe e desce da bolsa, o dinheiro depositado na sua conta via dividendos cobrirá a taxa que você estipulou."
        )
        
        c_m1, c_m2 = st.columns(2)
        mb_ex = c_m1.number_input("Margem Mínima Bazin Exigida (%):", value=5.0)
        mg_ex = c_m2.number_input("Margem Mínima Graham Exigida (%):", value=15.0)
        
        df_radar = pd.merge(df_perf_final[['Ativo', 'Tipo', 'Preço Atual']], st.session_state.df_recs_val, on='Ativo')
        
        status_bazin = []
        status_graham = []
        
        for _, row in df_radar.iterrows():
            if row['Teto Bazin'] > 0:
                if row['Margem Bazin (%)'] >= mb_ex: 
                    status_bazin.append("COMPRA 🟢")
                elif row['Margem Bazin (%)'] >= -5: 
                    status_bazin.append("MANTER 🟡")
                else: 
                    status_bazin.append("VENDA 🔴")
            else:
                status_bazin.append("MANTER 🟡")
                
            if row['Tipo'] == 'Ação' and pd.notna(row['Justo Graham']) and row['Justo Graham'] > 0:
                if row['Margem Graham (%)'] >= mg_ex: 
                    status_graham.append("COMPRA 🟢")
                elif row['Margem Graham (%)'] >= 0: 
                    status_graham.append("MANTER 🟡")
                else: 
                    status_graham.append("VENDA 🔴")
            else:
                status_graham.append("-")
                
        df_radar['Status Bazin'] = status_bazin
        df_radar['Status Graham'] = status_graham
        
        st.dataframe(df_radar[['Ativo', 'Tipo', 'Preço Atual', 'Teto Bazin', 'Margem Bazin (%)', 'Status Bazin', 'Justo Graham', 'Margem Graham (%)', 'Status Graham']].style.format(formatacao_t2 | {"Preço Atual": f_brl}), use_container_width=True, hide_index=True)

        st.markdown("##### ❄️ Projeção Bola de Neve (1 Ano)")
        saldo_dinamico = df_perf_final['Saldo Atual'].sum() + patr_fora
        b_div = st.session_state.df_simul['Div. Projetado (R$)'].sum() / 12 if not st.session_state.df_simul.empty else 0.0
        
        ac_ap = 0.0
        ac_jd = 0.0
        lp = []
        
        for m in range(13):
            if m == 0:
                lp.append({"Mês": f"Mês {m}", "Capital Inicial": saldo_dinamico, "Aportes Acumulados": 0.0, "Juros/Divs Acumulados": 0.0})
            else:
                gc = saldo_dinamico * rent
                div_m = b_div * ((1 + cresc_div) ** (m/12))
                ac_jd += (gc + div_m)
                ac_ap += aporte
                saldo_dinamico += (gc + div_m + aporte)
                
                lp.append({
                    "Mês": f"Mês {m}", 
                    "Capital Inicial": df_perf_final['Saldo Atual'].sum() + patr_fora, 
                    "Aportes Acumulados": ac_ap, 
                    "Juros/Divs Acumulados": ac_jd
                })
                
        df_proj_plot = pd.DataFrame(lp)
        df_melt_proj = df_proj_plot.melt(id_vars=["Mês"], value_vars=["Capital Inicial", "Aportes Acumulados", "Juros/Divs Acumulados"], var_name="Componente", value_name="Valor (R$)")
        
        fig_proj = px.bar(df_melt_proj, x="Mês", y="Valor (R$)", color="Componente", title="Evolução Patrimonial Controlada", color_discrete_sequence=['#1f4e78', '#00a896', '#f4a261'])
        st.plotly_chart(fig_proj, use_container_width=True)

    with tab_graf:
        st.markdown("#### Gráficos de Distribuição Patrimonial")
        cg1, cg2 = st.columns(2)
        
        paleta = ['#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087', '#f95d6a', '#ff7c43', '#ffa600']
        
        cg1.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Ativo', title="Por Ativo", color_discrete_sequence=paleta), use_container_width=True)
        cg2.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Tipo', title="Por Classe Operacional", color_discrete_sequence=['#1f4e78', '#00a896']), use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 📊 Comparativo Histórico e Indexadores")
        
        ativos_disponiveis = sorted(df_perf_final['Ativo'].unique().tolist())
        c_f_g1, c_f_g2 = st.columns(2)
        
        atv_sel = c_f_g1.multiselect("Comparar Ativos Específicos:", options=ativos_disponiveis, default=ativos_disponiveis[:5] if len(ativos_disponiveis) >= 5 else ativos_disponiveis)
        ind_sel = c_f_g2.multiselect("Comparar com os Indexadores:", ['CDI', 'IPCA'], default=['CDI', 'IPCA'])
        
        janela = st.radio("Período de Análise:", ["Desde a Data de Compra (Automático)", "Definir Período Customizado (Manual)"], horizontal=True)
        
        if janela == "Desde a Data de Compra (Automático)":
            if atv_sel:
                df_comp = df_perf_final[df_perf_final['Ativo'].isin(atv_sel)][['Ativo', 'Evolução c/ Div (%)', 'CDI Acum. (%)', 'IPCA Acum. (%)']].copy()
                df_comp = df_comp.rename(columns={'Evolução c/ Div (%)': 'Carteira (c/ Div)', 'CDI Acum. (%)': 'CDI', 'IPCA Acum. (%)': 'IPCA'})
                
                colunas_manter = ['Ativo', 'Carteira (c/ Div)'] + [ind for ind in ind_sel]
                df_comp = df_comp[colunas_manter]
                df_melt = df_comp.melt(id_vars='Ativo', var_name='Indicador', value_name='Rentabilidade (%)')
                
                fig_comp = px.bar(
                    df_melt, x='Ativo', y='Rentabilidade (%)', color='Indicador', barmode='group',
                    color_discrete_map={'Carteira (c/ Div)': '#003f5c', 'CDI': '#00a896', 'IPCA': '#f4a261'},
                    title="Rentabilidade Acumulada no Tempo de Posse"
                )
                st.plotly_chart(fig_comp, use_container_width=True)
            else:
                st.warning("Selecione ao menos um ativo para visualizar o gráfico.")
        else:
            c_dt1, c_dt2 = st.columns(2)
            dt_ini = c_dt1.date_input("De:", pd.Timestamp.now().date() - pd.Timedelta(days=365))
            dt_fim = c_dt2.date_input("Até:", pd.Timestamp.now().date())
            
            if st.button("Gerar Gráfico Comparativo", use_container_width=True):
                if atv_sel:
                    with st.spinner("Calculando série histórica..."):
                        cdi_m, ipca_m = 0.0, 0.0
                        if not df_macro.empty:
                            try:
                                f_m = df_macro.loc[dt_ini:dt_fim]
                                cdi_m = ((1 + f_m['CDI'].dropna()).prod() - 1) * 100
                                ipca_m = ((1 + f_m['IPCA'].dropna()).prod() - 1) * 100
                            except: pass
                            
                        l_res = []
                        for t in atv_sel:
                            r_atv = 0.0
                            try:
                                ht = yf.Ticker(f"{t}.SA").history(start=dt_ini, end=dt_fim)
                                if not ht.empty and len(ht) >= 2:
                                    p_i = ht['Close'].iloc[0]
                                    p_f = ht['Close'].iloc[-1]
                                    d_p = 0.0
                                    try:
                                        al_d = yf.Ticker(f"{t}.SA").dividends
                                        if not al_d.empty:
                                            if al_d.index.tz is not None: al_d.index = al_d.index.tz_localize(None)
                                            d_p = al_d[(al_d.index >= pd.Timestamp(dt_ini)) & (al_d.index <= pd.Timestamp(dt_fim))].sum()
                                    except: pass
                                    r_atv = ((p_f + d_p) / p_i - 1) * 100
                            except: pass
                            
                            it_m = {'Ativo': t, 'Carteira (c/ Div)': r_atv}
                            if 'CDI' in ind_sel: it_m['CDI'] = cdi_m
                            if 'IPCA' in ind_sel: it_m['IPCA'] = ipca_m
                            l_res.append(it_m)
                        
                        df_m_plot = pd.DataFrame(l_res)
                        if not df_m_plot.empty:
                            df_melt_m = df_m_plot.melt(id_vars='Ativo', var_name='Indicador', value_name='Rentabilidade (%)')
                            fig_comp_m = px.bar(
                                df_melt_m, x='Ativo', y='Rentabilidade (%)', color='Indicador', barmode='group',
                                color_discrete_map={'Carteira (c/ Div)': '#003f5c', 'CDI': '#00a896', 'IPCA': '#f4a261'},
                                title=f"Desempenho Customizado de {dt_ini.strftime('%d/%m/%Y')} até {dt_fim.strftime('%d/%m/%Y')}"
                            )
                            st.plotly_chart(fig_comp_m, use_container_width=True)
                else:
                    st.warning("Selecione ao menos um ativo para visualizar o gráfico.")

    with tab_prov:
        st.markdown("### 💸 Proventos Mensais e Status de Pagamento B3")
        cf1, cf2, c_btn = st.columns([2, 2, 2])
        m_map = {1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio", 6:"Junho", 7:"Julho", 8:"Agosto", 9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"}
        
        m_hoje = pd.Timestamp.now().month
        a_hoje = pd.Timestamp.now().year
        
        m_sel = cf1.selectbox("Mês do Provento:", options=list(m_map.keys()), format_func=lambda x: m_map[x], index=m_hoje-1)
        a_sel = cf2.selectbox("Ano de Referência:", options=[a_hoje, a_hoje-1, a_hoje-2])
        
        if c_btn.button("🔄 Processar Renda Mensal", use_container_width=True):
            with st.spinner("Buscando agenda de pagamentos na B3..."):
                la, lf = [], []
                for t_tk, dm in st.session_state.dados_mercado.items():
                    val = 0.0
                    try:
                        divs = yf.Ticker(f"{t_tk}.SA").dividends
                        if not divs.empty:
                            if divs.index.tz is not None: 
                                divs.index = divs.index.tz_localize(None)
                            val = float(divs[(divs.index.month == m_sel) & (divs.index.year == a_sel)].sum())
                    except: pass
                    
                    rec = val * dm['Qtd']
                    yoc = (rec / (dm['Qtd'] * dm['PM'])) * 100 if dm['PM'] > 0 else 0
                    
                    if dm['Tipo'] == 'FII':
                        lf.append({
                            "Fundo (FII)": t_tk, 
                            "Unitário (R$)": val, 
                            "Recebido (R$)": rec, 
                            "Yield on Cost (%)": yoc, 
                            "Status": "Divulgado / Pago 🟢" if val > 0 else "Aguardando 🟡"
                        })
                    else:
                        if val > 0: 
                            la.append({
                                "Ação": t_tk, 
                                "Unitário (R$)": val, 
                                "Recebido (R$)": rec, 
                                "Yield on Cost (%)": yoc, 
                                "Status": "Pago 🟢"
                            })
                
                st.session_state.divs_a = pd.DataFrame(la)
                st.session_state.divs_f = pd.DataFrame(lf)
                st.session_state.divs_m = m_sel
                st.session_state.divs_ano = a_sel
        
        if ('divs_a' in st.session_state and not st.session_state.divs_a.empty) or ('divs_f' in st.session_state and not st.session_state.divs_f.empty):
            tot_mes = 0.0
            
            if 'divs_f' in st.session_state and not st.session_state.divs_f.empty:
                st.markdown("#### 🏢 Status dos Fundos Imobiliários (FIIs)")
                st.dataframe(st.session_state.divs_f.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct}), use_container_width=True, hide_index=True)
                tot_mes += st.session_state.divs_f['Recebido (R$)'].sum()
                
            if 'divs_a' in st.session_state and not st.session_state.divs_a.empty:
                st.markdown("#### 📈 Ações Pagadoras")
                st.dataframe(st.session_state.divs_a.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct}), use_container_width=True, hide_index=True)
                tot_mes += st.session_state.divs_a['Recebido (R$)'].sum()
                
            st.success(f"**💰 Total Estimado de Proventos no Período ({m_map[st.session_state.divs_m]}/{st.session_state.divs_ano}):** {f_brl(tot_mes)}")
            
        st.markdown("---")
        st.markdown("### 🏛️ Histórico Analítico de Proventos (Tempo de Posse)")
        l_hist = []
        for t_hist, dm_hist in st.session_state.dados_mercado.items():
            try:
                divs_h = yf.Ticker(f"{t_hist}.SA").dividends
                if not divs_h.empty:
                    if divs_h.index.tz is not None: 
                        divs_h.index = divs_h.index.tz_localize(None)
                    divs_fil = divs_h[divs_h.index >= pd.Timestamp(dm_hist['Data'])]
                    for d_idx, val_h in divs_fil.items():
                        t_rec = val_h * dm_hist['Qtd']
                        inv_h = dm_hist['Qtd'] * dm_hist['PM']
                        yoc_h = (t_rec / inv_h) * 100 if inv_h > 0 else 0
                        dy_h = (val_h / dm_hist['Preço Atual']) * 100 if dm_hist['Preço Atual'] > 0 else 0
                        
                        l_hist.append({
                            "Data Ex": d_idx.date(), 
                            "Ativo": t_hist, 
                            "Unitário (R$)": float(val_h), 
                            "Quantidade": int(dm_hist['Qtd']), 
                            "Recebido (R$)": float(t_rec), 
                            "Yield on Cost (%)": float(yoc_h), 
                            "DY Atual (%)": float(dy_h)
                        })
            except: pass
            
        if l_hist:
            df_hist_tot = pd.DataFrame(l_hist).sort_values("Data Ex", ascending=False)
            c_h1, c_h2 = st.columns(2)
            atvs_disp = sorted(df_hist_tot['Ativo'].unique().tolist())
            
            atvs_sel = c_h1.multiselect("Filtrar Tabela por Ativo:", options=atvs_disp, default=atvs_disp)
            r_hist = c_h2.date_input("Filtrar Tabela por Período:", value=(min(df_hist_tot['Data Ex']), max(df_hist_tot['Data Ex'])))
            
            df_hist_f = df_hist_tot[df_hist_tot['Ativo'].isin(atvs_sel)]
            if isinstance(r_hist, tuple) and len(r_hist) == 2:
                df_hist_f = df_hist_f[(df_hist_f['Data Ex'] >= r_hist[0]) & (df_hist_f['Data Ex'] <= r_hist[1])]
                
            if not df_hist_f.empty:
                st.dataframe(df_hist_f.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct, "DY Atual (%)": f_pct}), use_container_width=True, hide_index=True)
                
                xls_buffer = to_excel(df_hist_f, sheet_name="Historico_Proventos")
                st.download_button(label="📥 Baixar Histórico de Proventos em Planilha (Excel)", data=xls_buffer, file_name=f"Historico_Proventos_{st.session_state.username}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

else:
    for tb in [tab_visao, tab_val, tab_radar, tab_graf, tab_prov]:
        with tb: 
            st.info("ℹ️ Adicione ativos na tabela de Controle Manual ou via upload de planilha na barra lateral. Depois, clique em **Conectar ao Mercado Vivo** para preencher essas abas.")

# ==========================================
# 8. ABAS ISOLADAS (SEMPRE ATIVAS)
# ==========================================
with tab_tesouro:
    st.markdown("### 🏛️ Simulador e Controle de Tesouro Direto")
    st.info("Faça o upload da sua planilha de controle do Tesouro ou insira os títulos manualmente. O sistema calculará o valor futuro baseado na inflação real vigente.")
    
    arq_tesouro = st.file_uploader("Upload da Planilha do Tesouro Direto (Excel/CSV)", type=["xlsx", "csv"], key="up_tesouro")
    
    if arq_tesouro:
        try:
            if arq_tesouro.name.endswith('.csv'):
                txt_tes = arq_tesouro.getvalue().decode('utf-8-sig', errors='ignore')
                sep_tes = '\t' if txt_tes.count('\t') > txt_tes.count(';') else ';'
                df_up_tes = pd.read_csv(io.StringIO(txt_tes), sep=sep_tes)
            else:
                df_up_tes = pd.read_excel(arq_tesouro)
            
            mapeamento_tes = {}
            for col in df_up_tes.columns:
                c_up = str(col).strip().upper()
                if 'TÍTULO' in c_up or 'TITULO' in c_up: mapeamento_tes[col] = 'Título'
                elif 'DATA' in c_up or 'COMPRA' in c_up: mapeamento_tes[col] = 'Data Compra'
                elif 'VALOR' in c_up or 'INVEST' in c_up: mapeamento_tes[col] = 'Valor Investido (R$)'
                elif 'TIPO' in c_up or 'PRÉ' in c_up or 'PÓS' in c_up or 'INDEX' in c_up: mapeamento_tes[col] = 'Tipo Taxa'
                elif 'TAXA' in c_up: mapeamento_tes[col] = 'Taxa Contratada (%)'
                elif 'VENC' in c_up or 'ANO' in c_up: mapeamento_tes[col] = 'Ano Vencimento'
            
            df_up_tes = df_up_tes.rename(columns=mapeamento_tes)
            colunas_padrao = ["Título", "Data Compra", "Tipo Taxa", "Valor Investido (R$)", "Taxa Contratada (%)", "Ano Vencimento"]
            
            for c in colunas_padrao:
                if c not in df_up_tes.columns: df_up_tes[c] = None
            
            df_up_tes['Data Compra'] = pd.to_datetime(df_up_tes['Data Compra'], errors='coerce').dt.date
            df_up_tes['Valor Investido (R$)'] = df_up_tes['Valor Investido (R$)'].apply(limpar_numero)
            df_up_tes['Taxa Contratada (%)'] = df_up_tes['Taxa Contratada (%)'].apply(limpar_numero)
            df_up_tes['Ano Vencimento'] = df_up_tes['Ano Vencimento'].apply(limpar_numero).astype(int)
            df_up_tes['Tipo Taxa'] = df_up_tes['Tipo Taxa'].astype(str).apply(lambda x: "Pós-fixado (IPCA+)" if "IPCA" in x.upper() or "PÓS" in x.upper() else "Pré-fixado")
            
            st.session_state.df_tesouro = df_up_tes[colunas_padrao]
            st.success("Planilha do Tesouro integrada com sucesso!")
        except Exception as e:
            st.error(f"Erro ao processar planilha: {e}")

    if st.session_state.df_tesouro.empty: 
        st.session_state.df_tesouro = pd.DataFrame([{
            "Título": "Tesouro IPCA+ 2029", "Data Compra": pd.Timestamp.now().date(), "Tipo Taxa": "Pós-fixado (IPCA+)",
            "Valor Investido (R$)": 1000.0, "Taxa Contratada (%)": 6.0, "Ano Vencimento": 2029
        }])
        
    st.markdown("#### 📝 Lançamento de Ativos (Campos Editáveis)")
    df_t = st.data_editor(
        st.session_state.df_tesouro, 
        num_rows="dynamic", 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "Tipo Taxa": st.column_config.SelectboxColumn("Tipo Taxa", options=["Pré-fixado", "Pós-fixado (IPCA+)"], required=True),
            "Data Compra": st.column_config.DateColumn("Data Compra"),
            "Ano Vencimento": st.column_config.NumberColumn("Ano Vencimento", format="%d")
        }
    )
    st.session_state.df_tesouro = df_t
    
    # Geração de Relatório com Cálculo Inline Automático
    if not df_t.empty:
        res_t = []
        ipca_projetado = ipca_12m_hoje if ipca_12m_hoje > 0 else 4.0
        
        for _, rt in df_t.iterrows():
            if pd.isna(rt.get('Título')) or str(rt.get('Título')).strip() == '':
                continue
            try:
                ano_venc = int(limpar_numero(rt.get('Ano Vencimento', pd.Timestamp.now().year + 5)))
                anos = max(1, ano_venc - pd.Timestamp.now().year)
                investido = float(limpar_numero(rt.get('Valor Investido (R$)', 0)))
                taxa_contrato = float(limpar_numero(rt.get('Taxa Contratada (%)', 0)))
                
                if str(rt.get('Tipo Taxa')).strip() == "Pós-fixado (IPCA+)":
                    taxa_real_aplicada = taxa_contrato + ipca_projetado
                else:
                    taxa_real_aplicada = taxa_contrato
                    
                v_final = investido * ((1 + (taxa_real_aplicada/100)) ** anos)
                
                res_t.append({
                    "Título": rt.get('Título'),
                    "Data Compra": rt.get('Data Compra'),
                    "Tipo Taxa": rt.get('Tipo Taxa'),
                    "Valor Investido (R$)": investido,
                    "Taxa Contratada (%)": taxa_contrato,
                    "Ano Vencimento": ano_venc,
                    "Taxa Equivalente (a.a.)": f"{taxa_real_aplicada:.2f}%",
                    "Valor Futuro no Vencimento": v_final,
                    "Projeção de Lucro Bruto": v_final - investido
                })
            except:
                continue
        
        if res_t:
            st.markdown("#### 🚀 Painel de Projeções Analíticas")
            df_res_t = pd.DataFrame(res_t)
            st.dataframe(df_res_t.style.format({
                "Valor Investido (R$)": f_brl, 
                "Valor Futuro no Vencimento": f_brl, 
                "Projeção de Lucro Bruto": f_brl,
                "Taxa Contratada (%)": lambda x: f"{x:.2f}%"
            }), use_container_width=True, hide_index=True)

with tab_ia:
    st.markdown("### 💬 Comitê de IA Sênior")
    cb1, cb2 = st.columns(2)
    
    if cb1.button("🗑️ Limpar Chat e Iniciar Nova Sessão", use_container_width=True):
        st.session_state.historico_chat = [{"role": "assistant", "content": f"Saudações, {st.session_state.username}. O terminal foi limpo."}]
        st.rerun()
        
    if HAS_DOCX and len(st.session_state.historico_chat) > 1:
        doc_buffer = export_docx(st.session_state.historico_chat)
        cb2.download_button("📄 Exportar Conversa de IA em Documento (Word)", data=doc_buffer, file_name=f"Parecer_IA_{st.session_state.username}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    elif not HAS_DOCX: 
        cb2.caption("⚠️ Instale a biblioteca 'python-docx' no seu arquivo requirements.txt para habilitar o botão de exportação em Word.")

    api_key_secreta = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key_secreta: 
        api_key_secreta = st.text_input("Insira sua Gemini API Key (Apenas uma vez):", type="password")
        
    for m in st.session_state.historico_chat:
        with st.chat_message(m["role"]): 
            st.write(m["content"])
        
    if prompt := st.chat_input("Ex: 'A BBAS3 está sendo negociada com desconto?' ou 'Qual sua avaliação da minha carteira?'"):
        with st.chat_message("user"): 
            st.write(prompt)
            
        st.session_state.historico_chat.append({"role": "user", "content": prompt})
        
        with st.chat_message("assistant"):
            with st.spinner("Analisando cruzamentos de dados e o cenário macroeconômico..."):
                ctx_c = str(st.session_state.dados_mercado) if st.session_state.dados_mercado else "O usuário não conectou a carteira no momento."
                ctx_m = f"Selic Vigente: {f_pct(selic_hoje)}|IPCA Atual 12m: {f_pct(ipca_12m_hoje)}. Projeções Focus para {ano_atual}: Selic {f_pct(proj_focus.get(f'Selic_{ano_atual}'))}/IPCA {f_pct(proj_focus.get(f'IPCA_{ano_atual}'))}"
                
                h_txt = "\n".join([f"{'Usuário' if h['role']=='user' else 'Gestora IA'}: {h['content']}" for h in st.session_state.historico_chat[-6:-1]])
                
                sys_p = (
                    f"Você é um Analista CNPI Sênior de alta performance. [Carteira do Cliente]: {ctx_c}. [Cenário Macro]: {ctx_m}.\n"
                    f"DIRETRIZ DE CONTINUIDADE: Use o HISTÓRICO RECENTE abaixo para não perder o fio da conversa com o cliente.\n"
                    f"REGRA ESTRITA E IMUTÁVEL DE ESCOPO: 1) Se o usuário usar a palavra 'carteira', 'meus ativos' ou nomes explícitos contidos na [Carteira do Cliente], analise profundamente o portfólio dele de forma consultiva e executiva. 2) Se a pergunta for neutra ou não citar ativos dele, VOCÊ DEVE IGNORAR TOTALMENTE a carteira e responder apenas tecnicamente sobre a dúvida.\n"
                    f"=== HISTÓRICO DA CONVERSA ===\n{h_txt}"
                )
                
                resposta_ia = "⚠️ Chave API da IA ausente ou incorreta."
                if api_key_secreta:
                    try:
                        import google.generativeai as genai
                        genai.configure(api_key=api_key_secreta)
                        sucesso_ia = False
                        erro_log = ""
                        
                        for mdl in ['gemini-2.5-flash', 'gemini-1.5-flash']:
                            try:
                                resposta_ia = genai.GenerativeModel(mdl).generate_content([sys_p, prompt]).text
                                sucesso_ia = True
                                break 
                            except Exception as e_ia: 
                                erro_log = str(e_ia)
                                continue
                                
                        if not sucesso_ia: 
                            resposta_ia = f"⚠️ Falha de comunicação de rede com os servidores da IA: {erro_log}"
                    except Exception as e_motor: 
                        resposta_ia = f"⚠️ Falta de dependência no servidor (motor genai): {e_motor}"
                        
                st.write(resposta_ia)
                
        st.session_state.historico_chat.append({"role": "assistant", "content": resposta_ia})
        
        if st.session_state.username:
            salvar_dados_completos_db(st.session_state.username) 
            
        st.rerun()
