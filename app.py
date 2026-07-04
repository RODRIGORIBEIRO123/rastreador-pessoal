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
            cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")
            for row_num in range(2, worksheet.max_row + 1):
                data_cell = worksheet.cell(row=row_num, column=col_num)
                data_cell.font, data_cell.border = Font(name="Arial", size=10), thin_border
                data_cell.alignment = Alignment(horizontal="right" if isinstance(data_cell.value, (int, float)) else "center")
            worksheet.column_dimensions[get_column_letter(col_num)].width = max(max(len(str(c.value or '')) for c in column) + 4, 13)
    return output.getvalue()

def export_docx(historico):
    if not HAS_DOCX: return None
    doc = docx.Document()
    doc.add_heading('Relatório - Gestora IA CNPI', 0)
    for msg in historico:
        doc.add_heading("Analista:" if msg["role"] == "user" else "Gestora IA:", level=2)
        doc.add_paragraph(msg["content"])
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

MAPEAMENTO_TICKERS = {"GALG11": "GARE11", "SOMA3": "ALOS3", "ARZZ3": "ALOS3", "VVAR3": "BHIA3", "VIIA3": "BHIA3", "BRML3": "ALSO3", "BBRK11": "BRCR11", "HCTR11": "TRXD11", "TORD11": "TRXD11"}
UNITS_ACOES = ['SANB11', 'TAEE11', 'KLBN11', 'BPAC11', 'ALUP11', 'ENGI11', 'BIDI11', 'CPLE11', 'SAPR11', 'RNEW11']

for k in ['df_base', 'df_tesouro', 'dados_mercado', 'df_simul']:
    if k not in st.session_state: st.session_state[k] = pd.DataFrame() if 'df' in k else {}
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'username' not in st.session_state: st.session_state.username = ""
if 'historico_chat' not in st.session_state: st.session_state.historico_chat = []

# ==========================================
# 2. MOTOR DE BANCO DE DADOS HÍBRIDO (BLINDADO)
# ==========================================
IS_POSTGRES = "POSTGRES_URL" in st.secrets
PARAM = "%s" if IS_POSTGRES else "?"

def get_db_connection():
    if IS_POSTGRES:
        import psycopg2
        return psycopg2.connect(st.secrets["POSTGRES_URL"])
    return sqlite3.connect("terminal_cnpi.db")

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (username TEXT PRIMARY KEY, password TEXT)''')
    if IS_POSTGRES:
        c.execute('''CREATE TABLE IF NOT EXISTS carteiras (username TEXT, ativo TEXT, quantidade DOUBLE PRECISION, preco_medio DOUBLE PRECISION, data_media TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS tesouro (username TEXT, titulo TEXT, investido DOUBLE PRECISION, taxa DOUBLE PRECISION, vencimento INTEGER)''')
    else:
        c.execute('''CREATE TABLE IF NOT EXISTS carteiras (username TEXT, Ativo TEXT, Quantidade REAL, Preco_Medio REAL, Data_Media TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS tesouro (username TEXT, titulo TEXT, investido REAL, taxa REAL, vencimento INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS chat_ia (username TEXT, role TEXT, content TEXT)''')
    conn.commit()
    conn.close()

def hash_pw(pwd): return hashlib.sha256(pwd.encode()).hexdigest()

def auth_user(u, p):
    conn = get_db_connection(); c = conn.cursor()
    c.execute(f"SELECT * FROM usuarios WHERE username={PARAM} AND password={PARAM}", (u, hash_pw(p)))
    res = c.fetchone(); conn.close()
    return res is not None

def reg_user(u, p):
    conn = get_db_connection()
    try:
        c = conn.cursor()
        c.execute(f"INSERT INTO usuarios (username, password) VALUES ({PARAM}, {PARAM})", (u, hash_pw(p)))
        conn.commit(); conn.close(); return True
    except:
        conn.close(); return False

def atualizar_senha(u, nova_senha):
    conn = get_db_connection(); c = conn.cursor()
    c.execute(f"SELECT * FROM usuarios WHERE username={PARAM}", (u,))
    if c.fetchone() is None: conn.close(); return False
    c.execute(f"UPDATE usuarios SET password={PARAM} WHERE username={PARAM}", (hash_pw(nova_senha), u))
    conn.commit(); conn.close()
    return True

def salvar_dados_completos_db(username):
    conn = get_db_connection(); c = conn.cursor()
    
    c.execute(f"DELETE FROM carteiras WHERE username={PARAM}", (username,))
    if not st.session_state.df_base.empty:
        for _, r in st.session_state.df_base.iterrows():
            c.execute(f"INSERT INTO carteiras (username, ativo, quantidade, preco_medio, data_media) VALUES ({PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM})",
                      (username, r['Ativo'], float(r['Quantidade']), float(r['Preço Médio']), str(r['Data Média'])))
            
    c.execute(f"DELETE FROM tesouro WHERE username={PARAM}", (username,))
    if not st.session_state.df_tesouro.empty:
        for _, r in st.session_state.df_tesouro.iterrows():
            c.execute(f"INSERT INTO tesouro (username, titulo, investido, taxa, vencimento) VALUES ({PARAM}, {PARAM}, {PARAM}, {PARAM}, {PARAM})",
                      (username, r['Título'], float(r['Valor Investido (R$)']), float(r['Taxa Anual (%)']), int(r['Ano Venc.'])))
            
    c.execute(f"DELETE FROM chat_ia WHERE username={PARAM}", (username,))
    for msg in st.session_state.historico_chat[-30:]:
        c.execute(f"INSERT INTO chat_ia (username, role, content) VALUES ({PARAM}, {PARAM}, {PARAM})", (username, msg['role'], msg['content']))
        
    conn.commit(); conn.close()

def carregar_dados_completos_db(username):
    conn = get_db_connection()
    
    query_cart = f"SELECT Ativo, Quantidade, Preco_Medio, Data_Media FROM carteiras WHERE username={PARAM}"
    df_cart = pd.read_sql_query(query_cart, conn, params=(username,))
    df_cart = df_cart.rename(columns={"Preco_Medio": "Preço Médio", "preco_medio": "Preço Médio", "Data_Media": "Data Média", "data_media": "Data Média", "ativo": "Ativo", "quantidade": "Quantidade"})
    if not df_cart.empty: df_cart['Data Média'] = pd.to_datetime(df_cart['Data Média']).dt.date
    st.session_state.df_base = df_cart
    
    query_tes = f"SELECT titulo, investido, taxa, vencimento FROM tesouro WHERE username={PARAM}"
    df_tes = pd.read_sql_query(query_tes, conn, params=(username,))
    df_tes = df_tes.rename(columns={"titulo": "Título", "investido": "Valor Investido (R$)", "taxa": "Taxa Anual (%)", "vencimento": "Ano Venc."})
    st.session_state.df_tesouro = df_tes
    
    query_chat = f"SELECT role, content FROM chat_ia WHERE username={PARAM}"
    df_chat = pd.read_sql_query(query_chat, conn, params=(username,))
    if not df_chat.empty: st.session_state.historico_chat = df_chat.to_dict('records')
    else: st.session_state.historico_chat = [{"role": "assistant", "content": f"Saudações, {username}. O terminal está mapeado e online. Como posso ajudar?"}]
        
    conn.close()

init_db()

# ==========================================
# 3. TELA DE AUTENTICAÇÃO
# ==========================================
if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center;'>🔐 Terminal de Gestão Profissional</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Acesso restrito. Identifique-se para carregar seu portfólio.</p>", unsafe_allow_html=True)
    
    col_log1, col_log2, col_log3 = st.columns([1, 1, 1])
    with col_log2:
        tab_login, tab_register, tab_forgot = st.tabs(["Acesso", "Novo Registro", "Recuperar Senha"])
        with tab_login:
            lu = st.text_input("Usuário", key="log_user")
            lp = st.text_input("Senha", type="password", key="log_pass")
            if st.button("Entrar", use_container_width=True):
                if auth_user(lu, lp):
                    st.session_state.logged_in, st.session_state.username = True, lu
                    carregar_dados_completos_db(lu)
                    st.rerun()
                else: st.error("Credenciais inválidas.")
        with tab_register:
            ru = st.text_input("Novo Usuário", key="reg_user")
            rp = st.text_input("Nova Senha", type="password", key="reg_pass")
            if st.button("Registrar", use_container_width=True):
                if ru and rp:
                    if reg_user(ru, rp): st.success("Conta criada! Pode fazer o login.")
                    else: st.error("Usuário já existe.")
                else: st.warning("Preencha ambos os campos.")
        with tab_forgot:
            fu = st.text_input("Usuário Cadastrado", key="for_user")
            fp = st.text_input("Nova Senha", type="password", key="for_pass")
            if st.button("Redefinir Senha", use_container_width=True):
                if atualizar_senha(fu, fp): st.success("Senha redefinida com sucesso.")
                else: st.error("Usuário não encontrado.")
    st.stop()

# ==========================================
# 4. FUNÇÕES DE DADOS E PROCESSAMENTO B3
# ==========================================
st.markdown(f"""<div style="text-align: center; margin-bottom: 20px;"><h3 style="font-weight: 400; margin-bottom: 0;">💼 Terminal de Gestão</h3><h6 style="color: #666; font-weight: 300;">Analista: {st.session_state.username.upper()}</h6></div>""", unsafe_allow_html=True)
st.write("---")

@st.cache_data(ttl=86400)
def carregar_dados_mercado():
    macro, fundamentos, selic, ipca_12m = pd.DataFrame(), {}, 10.50, 4.00
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'], macro['IPCA'] = macro['CDI'] / 100, macro['IPCA'] / 100
        df = pd.read_html(io.StringIO(requests.get('https://www.fundamentus.com.br/resultado.php', headers={'User-Agent': 'Mozilla/5.0'}).text), decimal=',', thousands='.')[0]
        for _, r in df.iterrows():
            t, c, pl, pvp = str(r['Papel']).strip().upper(), float(r['Cotação']), float(r['P/L']), float(r['P/VP'])
            fundamentos[t] = {'vpa': c/pvp if pvp>0 else 0.0, 'lpa': c/pl if pl>0 else 0.0}
        selic = float(requests.get("https://brasilapi.com.br/api/taxas/v1", timeout=5).json()[0]['valor'])
        ipca_df = sgs.get({'IPCA_12M': 13522}, last=1)
        if not ipca_df.empty: ipca_12m = float(ipca_df['IPCA_12M'].iloc[-1])
    except: pass
    
    ano_at, proj = pd.Timestamp.now().year, {}
    try:
        url = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$top=300&$filter=Indicador%20eq%20'IPCA'%20or%20Indicador%20eq%20'Selic'&$orderby=Data%20desc&$format=json"
        df_p = pd.DataFrame(requests.get(url, timeout=5).json().get('value', []))
        if not df_p.empty:
            df_p = df_p[df_p['Data'] == df_p['Data'].max()]
            for ao in [0, 1, 2]:
                aa = str(ano_at + ao)
                df_a = df_p[df_p['DataReferencia'] == aa]
                if not df_a[df_a['Indicador'] == 'IPCA'].empty: proj[f"IPCA_{aa}"] = float(df_a[df_a['Indicador'] == 'IPCA']['Mediana'].values[0])
                if not df_a[df_a['Indicador'] == 'Selic'].empty: proj[f"Selic_{aa}"] = float(df_a[df_a['Indicador'] == 'Selic']['Mediana'].values[0])
    except: pass
    return macro, fundamentos, selic, ipca_12m, proj, ano_at

df_macro, fundamentos_br, selic_hoje, ipca_12m_hoje, proj_focus, ano_atual = carregar_dados_mercado()

def limpar_numero(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float, np.number)): return float(x)
    try: return float(str(x).replace('R$', '').replace('.', '').replace(',', '.').strip())
    except: return 0.0

def traduzir_setor(setor_en):
    return {"Banks": "Bancos", "Utilities - Regulated Electric": "Energia", "Real Estate - Retail": "Shoppings/Varejo", "REIT - Retail": "Shoppings/Varejo", "Real Estate - Industrial": "Logística", "REIT - Industrial": "Logística", "REIT - Office": "Lajes Corporativas", "REIT - Diversified": "Fundo Híbrido", "Financial Data & Stock Exchanges": "Bolsa de Valores", "Insurance": "Seguradoras", "Oil & Gas Integrated": "Petróleo e Gás"}.get(setor_en, "Outros Setores")

def consolidar_carteira(df):
    if df.empty: return df
    df['Ativo'] = df['Ativo'].astype(str).str.strip().str.upper().apply(lambda x: MAPEAMENTO_TICKERS.get(x, x))
    linhas = []
    for ativo, group in df.groupby('Ativo'):
        qtd = float(group['Quantidade'].sum())
        if qtd > 0:
            pm = (group['Quantidade'] * group['Preço Médio']).sum() / qtd
            ts = sum((pd.Timestamp(r['Data Média']).timestamp() * r['Quantidade']) for _, r in group.iterrows() if pd.notna(r['Data Média']))
            linhas.append({"Ativo": ativo, "Quantidade": qtd, "Preço Médio": float(pm), "Data Média": pd.to_datetime(ts/qtd, unit='s').date()})
    return pd.DataFrame(linhas)

def corrigir_cabecalho_b3(df):
    if df.empty: return df
    if 'Data do Negócio' in df.columns or 'Data Média' in df.columns: return df
        
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
                elif 'PRE' in c_up: mapeamento[col] = 'Preço Unitário'
                elif 'ENTRADA' in c_up or 'SAÍDA' in c_up or 'SAIDA' in c_up: mapeamento[col] = 'Entrada/Saída'
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
    if df.empty: return pd.DataFrame()
    df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
    df['Quantidade'] = df['Quantidade'].apply(limpar_numero)
    if 'Valor' in df.columns: df['Valor'] = df['Valor'].apply(limpar_numero)
    elif 'Preço Unitário' in df.columns: df['Valor'] = df['Quantidade'] * df['Preço Unitário'].apply(limpar_numero)
    else: df['Valor'] = 0.0

    pos = {}
    for _, r in df.sort_values('Data do Negócio').iterrows():
        tk = MAPEAMENTO_TICKERS.get(str(r.get('Código de Negociação', '')).strip().upper().split(" ")[0].replace('F',''), str(r.get('Código de Negociação', '')).strip().upper().split(" ")[0].replace('F',''))
        if not re.match(r'^[A-Z]{4}\d{1,2}$', tk): continue
        if tk not in pos: pos[tk] = {'qtd': 0.0, 'valor': 0.0, 'ts': 0.0}
        q, v, mov, io_dir = float(r['Quantidade']), float(r['Valor']), str(r.get('Tipo de Movimentação', '')).upper(), str(r.get('Entrada/Saída', '')).upper()
        if 'COMPRA' in mov or 'CRED' in mov or 'ENT' in io_dir:
            pos[tk]['ts'] = pd.Timestamp(r['Data do Negócio']).timestamp() if pos[tk]['qtd']==0 else ((pos[tk]['ts']*pos[tk]['qtd'])+(pd.Timestamp(r['Data do Negócio']).timestamp()*q))/(pos[tk]['qtd']+q)
            pos[tk]['qtd'] += q; pos[tk]['valor'] += v
        elif 'VENDA' in mov or 'DEB' in mov or 'SAI' in io_dir:
            pm = pos[tk]['valor']/pos[tk]['qtd'] if pos[tk]['qtd']>0 else 0
            pos[tk]['qtd'] = max(0, pos[tk]['qtd'] - q)
            pos[tk]['valor'] = max(0, pos[tk]['valor'] - (q*pm))
            
    return consolidar_carteira(pd.DataFrame([{"Ativo": k, "Quantidade": v['qtd'], "Preço Médio": v['valor']/v['qtd'] if v['qtd']>0 else 0, "Data Média": pd.to_datetime(v['ts'], unit='s').date()} for k, v in pos.items() if v['qtd']>0]))

# ==========================================
# 5. SIDEBAR: UPLOAD E DB
# ==========================================
with st.sidebar:
    st.markdown("### 👤 ANALISTA OPERACIONAL")
    if st.button("🚪 Sair do Terminal", use_container_width=True):
        st.session_state.logged_in, st.session_state.username = False, ""
        st.rerun()
    st.divider()
    
    st.markdown("### 💾 Backup e Segurança")
    if st.button("Sincronizar no Banco de Dados", type="primary", use_container_width=True):
        salvar_dados_completos_db(st.session_state.username)
        st.success("Dados blindados no banco!")
    st.divider()
    
    st.markdown("### 1. Integrar Notas B3")
    st.info("Faça upload de planilhas para substituir ou somar operações à sua carteira.")
    arq_p = st.file_uploader("Substituir Base Completa (B3 / CSV)", type=["xlsx", "csv"])
    arq_n = st.file_uploader("Apenas Novas Operações", type=["xlsx", "csv"])
    dt_corte = st.date_input("Filtrar Novas a partir de:", pd.Timestamp.now().date() - pd.Timedelta(days=15)) if arq_n else None

    if st.button("🚀 Processar Arquivos", use_container_width=True):
        b_atual = st.session_state.df_base.copy()
        if arq_p:
            txt = arq_p.getvalue().decode('utf-8-sig', errors='ignore') if arq_p.name.endswith('.csv') else None
            df_p = pd.read_csv(io.StringIO(txt), sep=('\t' if txt and txt.count('\t')>txt.count(';') else ';')) if txt else pd.read_excel(arq_p)
            df_p = corrigir_cabecalho_b3(df_p)
            if 'Data Média' in df_p.columns: b_atual = consolidar_carteira(df_p)
            elif 'Data do Negócio' in df_p.columns: b_atual = processar_planilha_b3(df_p)
            else: st.error("Formato inválido."); st.stop()
                
        if arq_n and not b_atual.empty:
            txt_n = arq_n.getvalue().decode('utf-8-sig', errors='ignore') if arq_n.name.endswith('.csv') else None
            df_n = pd.read_csv(io.StringIO(txt_n), sep=('\t' if txt_n and txt_n.count('\t')>txt_n.count(';') else ';')) if txt_n else pd.read_excel(arq_n)
            df_n = corrigir_cabecalho_b3(df_n)
            if not df_n.empty and 'Data do Negócio' in df_n.columns:
                df_n['Data do Negócio'] = pd.to_datetime(df_n['Data do Negócio'], dayfirst=True, errors='coerce')
                df_n = df_n[df_n['Data do Negócio'].dt.date >= dt_corte]
                linhas_b = [{"Código de Negociação": r['Ativo'], "Tipo de Movimentação": "Compra", "Data do Negócio": pd.to_datetime(r['Data Média']), "Quantidade": r['Quantidade'], "Valor": r['Quantidade']*r['Preço Médio']} for _, r in b_atual.iterrows()]
                b_atual = processar_planilha_b3(pd.concat([pd.DataFrame(linhas_b), df_n], ignore_index=True))
                
        st.session_state.df_base = b_atual
        st.warning("Memória atualizada. Lembre-se de salvar no DB.")
        st.rerun()

# ==========================================
# 6. PAINEL MACRO E CONTROLE MANUAL
# ==========================================
st.markdown("### 👑 Conjuntura Macroeconômica")
c_m1, c_m2 = st.columns([1, 2])
c_m1.success(f"🎯 **Cenário Atual (Vigente)**\n\nSelic: **{f_pct(selic_hoje)} a.a.**\n\nIPCA 12m: **{f_pct(ipca_12m_hoje)}**")
c_m2.info(f"🔮 **Projeções Focus (Mercado)**\n\n**Selic:** {ano_atual}: **{f_pct(proj_focus.get(f'Selic_{ano_atual}', 0))}** |  {ano_atual+1}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+1}', 0))}** |  {ano_atual+2}: **{f_pct(proj_focus.get(f'Selic_{ano_atual+2}', 0))}**\n\n**IPCA:** {ano_atual}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual}', 0))}** |  {ano_atual+1}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+1}', 0))}** |  {ano_atual+2}: **{f_pct(proj_focus.get(f'IPCA_{ano_atual+2}', 0))}**")
st.write("---")

st.markdown("### 2. Controle Operacional da Carteira (Manual)")
c_op1, c_op2, c_op3 = st.columns([1, 1.5, 1])
with c_op1:
    tdel = st.selectbox("Excluir Ativo da Memória:", [""] + sorted(st.session_state.df_base["Ativo"].tolist()) if not st.session_state.df_base.empty else [""])
    if st.button("Remover Ativo Selecionado", use_container_width=True) and tdel:
        st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != tdel]
        st.rerun()
with c_op2:
    nt = st.text_input("Novo Aporte (Ticker)")
    cq, cp = st.columns(2)
    nq, np_v = cq.number_input("Qtd", min_value=1), cp.number_input("PM (R$)", min_value=0.01)
    if st.button("Adicionar à Carteira", use_container_width=True) and nt:
        st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, pd.DataFrame([{"Ativo": nt.upper(), "Quantidade": float(nq), "Preço Médio": float(np_v), "Data Média": pd.Timestamp.now().date()}])], ignore_index=True))
        st.rerun()
with c_op3:
    st.download_button("📥 Baixar Carteira Ajustada (CSV)", data=st.session_state.df_base.to_csv(index=False, sep=';', encoding='utf-8-sig'), file_name=f"Carteira_Ajustada_{st.session_state.username}.csv", use_container_width=True)

st.markdown("#### 📝 Tabela Editável (Ajuste Fino)")
st.info("Ajuste quantidades ou preços médios diretamente nas células abaixo antes de conectar ao mercado.")
df_editado = st.data_editor(st.session_state.df_base, use_container_width=True, hide_index=True)

if st.button("🚀 Conectar ao Mercado Vivo", type="primary", use_container_width=True):
    if not df_editado.empty:
        st.session_state.df_base = consolidar_carteira(df_editado)
        prg, tot, dm, sim = st.progress(0), len(st.session_state.df_base), {}, []
        for i, r in st.session_state.df_base.iterrows():
            t, dc = str(r['Ativo']).upper(), pd.to_datetime(r['Data Média']) if pd.notna(r['Data Média']) else pd.Timestamp.now()
            p_at, d_tot, d_12m, is_fii = float(r['Preço Médio']), 0.0, 0.0, t.endswith('11') and t not in UNITS_ACOES
            try:
                tk = yf.Ticker(f"{t}.SA")
                h = tk.history(period="1d")
                if not h.empty: p_at = float(h['Close'].iloc[-1])
                dvs = tk.dividends
                if not dvs.empty:
                    if dvs.index.tz is not None: dvs.index = dvs.index.tz_localize(None)
                    d_tot = float(dvs[dvs.index >= dc].sum() * r['Quantidade'])
                    d_12m = float(dvs[dvs.index >= (pd.Timestamp.now() - pd.DateOffset(years=1))].sum())
            except: pass
            
            cdi_ac, ipca_ac = 0.0, 0.0
            try: 
                f_m = df_macro.loc[dc:]
                cdi_ac, ipca_ac = ((1+f_m['CDI'].dropna()).prod()-1)*100, ((1+f_m['IPCA'].dropna()).prod()-1)*100
            except: pass
            
            dm[t] = {"Qtd": float(r['Quantidade']), "PM": float(r['Preço Médio']), "Data": dc, "Preço Atual": p_at, "Div_Total": d_tot, "CDI": cdi_ac, "IPCA": ipca_ac, "Tipo": "FII" if is_fii else "Ação", "Setor": traduzir_setor(tk.info.get('industry', '')) if 'tk' in locals() else "Desconhecido"}
            sim.append({"Ativo": t, "Cotação Atual": p_at, "VPA (Contábil)": fundamentos_br.get(t, {}).get('vpa', 0), "LPA Projetado": fundamentos_br.get(t, {}).get('lpa', 0), "Div. Projetado (R$)": d_12m})
            prg.progress((i+1)/tot)
        st.session_state.dados_mercado, st.session_state.df_simul = dm, pd.DataFrame(sim)
        st.success("Mercado Sincronizado! As abas de análise estão ativas.")
    else: st.warning("A carteira está vazia.")
st.write("---")

# ==========================================
# 7. DASHBOARDS E RELATÓRIOS (TABS)
# ==========================================
t1, t2, t3, t4, t5, t_tes, t6 = st.tabs(["📊 Visão Geral", "💰 Valuation", "🎯 Radar & Projeção", "📈 Gráficos", "💸 Proventos B3", "🏛️ Tesouro Direto", "💬 Gestora IA (CNPI)"])

if st.session_state.dados_mercado:
    l_pf = []
    for t, dm in st.session_state.dados_mercado.items():
        inv, sld = dm['Qtd'] * dm['PM'], dm['Qtd'] * dm['Preço Atual']
        l_pf.append({
            "Ativo": t, "Tipo": dm["Tipo"], "Setor": dm.get("Setor", "Desconhecido"), "Qtd": int(dm['Qtd']), 
            "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
            "Total Investido": inv, "Saldo Atual": sld, "Saldo C/ Dividendos": sld + dm['Div_Total'],
            "Resultado (R$)": sld - inv, "Resultado C/ Dividendos": (sld - inv) + dm['Div_Total'],
            "Data Média": dm['Data'].strftime('%d/%m/%Y'), "Total Div. (R$)": dm['Div_Total'], 
            "DY on Cost (%)": (dm['Div_Total'] / inv)*100 if inv>0 else 0, 
            "Evolução c/ Div (%)": (((sld + dm['Div_Total']) / inv)-1)*100 if inv>0 else 0,
            "IPCA Acum. (%)": dm['IPCA'], "CDI Acum. (%)": dm['CDI']
        })
    df_perf_final = pd.DataFrame(l_pf)

    with t1:
        st.markdown("### 🏆 Visão Global da Carteira")
        df_a, df_f = df_pf[df_pf['Tipo'] == 'Ação'], df_pf[df_pf['Tipo'] == 'FII']
        
        ev_a = (df_a['Saldo Atual'].sum() / df_a['Total Investido'].sum() - 1)*100 if df_a['Total Investido'].sum()>0 else 0
        ev_f = (df_f['Saldo Atual'].sum() / df_f['Total Investido'].sum() - 1)*100 if df_f['Total Investido'].sum()>0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📈 Patrimônio Ações", f_brl(df_a['Saldo Atual'].sum()), f_pct(ev_a))
        m2.metric("🏢 Patrimônio FIIs", f_brl(df_f['Saldo Atual'].sum()), f_pct(ev_f))
        m3.metric("💸 Renda Histórica Ações", f_brl(df_a['Total Div. (R$)'].sum()))
        m4.metric("💸 Renda Histórica FIIs", f_brl(df_f['Total Div. (R$)'].sum()))

        st.dataframe(df_perf_final.drop(columns=['Tipo', 'Setor']).style.format({c: f_brl for c in ["Preço Médio", "Preço Atual", "Total Investido", "Saldo Atual", "Saldo C/ Dividendos", "Resultado (R$)", "Resultado C/ Dividendos", "Total Div. (R$)"]}|{c: f_pct for c in ["DY on Cost (%)", "Evolução c/ Div (%)", "IPCA Acum. (%)", "CDI Acum. (%)"]}), use_container_width=True, hide_index=True)

    with t2:
        st.markdown("#### Métodos Certificados de Valuation")
        st.markdown("""* **Preço Teto Decio Bazin:** Calcula o preço máximo ideal para compra focado em dividendos.\n* **Preço Justo Benjamin Graham:** Avalia o valor de fábrica (patrimônio/lucro). Não aplicável a FIIs.""")
        
        yd = st.number_input("Taxa de Retorno Mínima Exigida Bazin (%):", value=6.0, step=0.5) / 100.0
        df_sim = st.data_editor(st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]], use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"])
        st.session_state.df_simul[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]] = df_sim[["Div. Projetado (R$)", "VPA (Contábil)", "LPA Projetado"]]
        
        rv = []
        for _, r in df_sim.iterrows():
            is_fii = r['Ativo'].endswith('11') and r['Ativo'] not in UNITS_ACOES
            bz = (float(r["Div. Projetado (R$)"]) / yd) if float(r["Div. Projetado (R$)"]) > 0 else 0.0
            mbz = ((bz / float(r["Cotação Atual"])) - 1) * 100 if bz > 0 else 0.0
            if not is_fii:
                gh = (22.5 * float(r["LPA Projetado"]) * float(r["VPA (Contábil)"]))**0.5 if float(r["LPA Projetado"]) > 0 and float(r["VPA (Contábil)"]) > 0 else 0.0
                mgh = ((gh / float(r["Cotação Atual"])) - 1) * 100 if gh > 0 else 0.0
            else:
                gh, mgh = np.nan, np.nan
            rv.append({"Ativo": r['Ativo'], "Teto Bazin": bz, "Margem Bazin (%)": mbz, "Justo Graham": gh, "Margem Graham (%)": mgh})
            
        st.session_state.df_recs_val = pd.DataFrame(rv)
        st.dataframe(st.session_state.df_recs_val.style.format({"Teto Bazin": lambda x: f_brl(x) if x > 0 else "-", "Justo Graham": lambda x: f_brl(x) if pd.notna(x) and x > 0 else "-", "Margem Bazin (%)": lambda x: f_pct(x) if x != 0 else "-", "Margem Graham (%)": lambda x: f_pct(x) if pd.notna(x) and x != 0 else "-"}), use_container_width=True, hide_index=True)

    with t3: 
        st.markdown("##### Parametrização do Radar Operacional")
        c_p1, c_p2, c_p3, c_p4 = st.columns(4)
        patr_fora = c_p1.number_input("Patrimônio Externo (R$):", value=0.0, step=1000.0)
        aporte = c_p2.number_input("Aporte Mensal Previsto (R$):", value=2000.0, step=500.0)
        rent = c_p3.number_input("Rentabilidade Mensal Alvo (%):", value=0.8, step=0.1) / 100.0
        cresc_div = c_p4.number_input("Crescimento Anual de Dividendos (%):", value=5.0, step=1.0) / 100.0

        st.markdown("##### 🎯 Triagem Estratégica Corporativa")
        c_m1, c_m2 = st.columns(2)
        mb_ex = c_m1.number_input("Margem Mínima Bazin Exigida (%):", value=5.0)
        mg_ex = c_m2.number_input("Margem Mínima Graham Exigida (%):", value=15.0)
        
        df_radar = pd.merge(df_pf[['Ativo', 'Tipo', 'Preço Atual']], st.session_state.df_recs_val, on='Ativo')
        df_radar['Status Bazin'] = df_radar.apply(lambda r: "COMPRA 🟢" if r['Teto Bazin']>0 and r['Margem Bazin (%)'] >= mb_ex else ("MANTER 🟡" if r['Teto Bazin']>0 and r['Margem Bazin (%)'] >= -5 else "VENDA 🔴"), axis=1)
        df_radar['Status Graham'] = df_radar.apply(lambda r: "COMPRA 🟢" if r['Tipo']=='Ação' and pd.notna(r['Justo Graham']) and r['Margem Graham (%)'] >= mg_ex else ("MANTER 🟡" if r['Tipo']=='Ação' and pd.notna(r['Justo Graham']) and r['Margem Graham (%)'] >= 0 else ("VENDA 🔴" if r['Tipo']=='Ação' else "-")), axis=1)
        st.dataframe(df_radar[['Ativo', 'Tipo', 'Preço Atual', 'Teto Bazin', 'Margem Bazin (%)', 'Status Bazin', 'Justo Graham', 'Margem Graham (%)', 'Status Graham']].style.format({"Preço Atual": f_brl, "Teto Bazin": lambda x: f_brl(x) if x > 0 else "-", "Justo Graham": lambda x: f_brl(x) if pd.notna(x) and x > 0 else "-", "Margem Bazin (%)": lambda x: f_pct(x) if x != 0 else "-", "Margem Graham (%)": lambda x: f_pct(x) if pd.notna(x) and x != 0 else "-"}), use_container_width=True, hide_index=True)

        st.markdown("##### ❄️ Projeção Bola de Neve (1 Ano)")
        saldo_dinamico, b_div, ac_ap, ac_jd, lp = df_pf['Saldo Atual'].sum() + patr_fora, st.session_state.df_simul['Div. Projetado (R$)'].sum() / 12 if not st.session_state.df_simul.empty else 0.0, 0.0, 0.0, []
        for m in range(13):
            if m > 0:
                gc, div_m = saldo_dinamico * rent, b_div * ((1 + cresc_div) ** (m/12))
                ac_jd += (gc + div_m); ac_ap += aporte; saldo_dinamico += (gc + div_m + aporte)
            lp.append({"Mês": f"Mês {m}", "Capital Inicial": df_pf['Saldo Atual'].sum() + patr_fora, "Aportes Acumulados": ac_ap, "Juros/Divs Acumulados": ac_jd})
        st.plotly_chart(px.bar(pd.DataFrame(lp).melt(id_vars=["Mês"], value_vars=["Capital Inicial", "Aportes Acumulados", "Juros/Divs Acumulados"], var_name="Componente", value_name="Valor (R$)"), x="Mês", y="Valor (R$)", color="Componente", title="Evolução Patrimonial Controlada", color_discrete_sequence=['#1f4e78', '#00a896', '#f4a261']), use_container_width=True)

    with t4:
        st.markdown("#### Gráficos de Distribuição Patrimonial")
        c_g1, c_g2 = st.columns(2)
        cores_institucionais = ['#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087', '#f95d6a', '#ff7c43', '#ffa600']
        c_g1.plotly_chart(px.pie(df_pf, values='Saldo Atual', names='Ativo', title="Por Ativo", color_discrete_sequence=cores_institucionais), use_container_width=True)
        c_g2.plotly_chart(px.pie(df_pf, values='Saldo Atual', names='Tipo', title="Por Classe (Ação vs FII)", color_discrete_sequence=['#1f4e78', '#00a896']), use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 📊 Gráfico Dinâmico Comparativo (Desempenho Real)")
        ativos_disponiveis = sorted(df_pf['Ativo'].unique().tolist())
        c_f_g1, c_f_g2 = st.columns(2)
        atv_sel = c_f_g1.multiselect("Selecione os Ativos:", options=ativos_disponiveis, default=ativos_disponiveis[:5] if len(ativos_disponiveis) >= 5 else ativos_disponiveis)
        ind_sel = c_f_g2.multiselect("Comparar com:", ['CDI', 'IPCA'], default=['CDI', 'IPCA'])
        
        janela = st.radio("Período de Leitura:", ["Desde a Data de Compra (Automático)", "Definir Período Customizado (Manual)"], horizontal=True)
        
        if janela == "Desde a Data de Compra (Automático)":
            if atv_sel:
                df_comp = df_pf[df_pf['Ativo'].isin(atv_sel)][['Ativo', 'Evolução c/ Div (%)', 'CDI Acum. (%)', 'IPCA Acum. (%)']].rename(columns={'Evolução c/ Div (%)': 'Carteira (c/ Div)', 'CDI Acum. (%)': 'CDI', 'IPCA Acum. (%)': 'IPCA'})
                df_melt = df_comp[['Ativo', 'Carteira (c/ Div)'] + ind_sel].melt(id_vars='Ativo', var_name='Indicador', value_name='Rentabilidade (%)')
                st.plotly_chart(px.bar(df_melt, x='Ativo', y='Rentabilidade (%)', color='Indicador', barmode='group', color_discrete_map={'Carteira (c/ Div)': '#003f5c', 'CDI': '#00a896', 'IPCA': '#f4a261'}, title="Rentabilidade Acumulada no Tempo de Posse"), use_container_width=True)
        else:
            c_dt1, c_dt2 = st.columns(2)
            dt_ini, dt_fim = c_dt1.date_input("De:", pd.Timestamp.now().date() - pd.Timedelta(days=365)), c_dt2.date_input("Até:", pd.Timestamp.now().date())
            
            if st.button("Gerar Gráfico Comparativo", use_container_width=True) and atv_sel:
                with st.spinner("Calculando histórico..."):
                    df_m_hist = carregar_macro()
                    cdi_m, ipca_m = 0.0, 0.0
                    if not df_m_hist.empty:
                        try:
                            f_m = df_m_hist.loc[dt_ini:dt_fim]
                            cdi_m, ipca_m = ((1 + f_m['CDI'].dropna()).prod() - 1) * 100, ((1 + f_m['IPCA'].dropna()).prod() - 1) * 100
                        except: pass
                    l_res = []
                    for t in atv_sel:
                        r_atv = 0.0
                        try:
                            ht = yf.Ticker(f"{t}.SA").history(start=dt_ini, end=dt_fim)
                            if not ht.empty and len(ht) >= 2:
                                p_i, p_f, d_p = ht['Close'].iloc[0], ht['Close'].iloc[-1], 0.0
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
                        st.plotly_chart(px.bar(df_m_plot.melt(id_vars='Ativo', var_name='Indicador', value_name='Rentabilidade (%)'), x='Ativo', y='Rentabilidade (%)', color='Indicador', barmode='group', color_discrete_map={'Carteira (c/ Div)': '#003f5c', 'CDI': '#00a896', 'IPCA': '#f4a261'}, title=f"Desempenho de {dt_ini.strftime('%d/%m/%Y')} até {dt_fim.strftime('%d/%m/%Y')}"), use_container_width=True)

    with t5:
        st.markdown("### 💸 Proventos Mensais e Status de Pagamento")
        c_f1, c_f2, c_btn = st.columns([2, 2, 2])
        m_map = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
        m_hoje, a_hoje = pd.Timestamp.now().month, pd.Timestamp.now().year
        
        m_sel = c_f1.selectbox("Mês:", options=list(m_map.keys()), format_func=lambda x: m_map[x], index=m_hoje-1)
        a_sel = c_f2.selectbox("Ano:", options=[a_hoje, a_hoje-1, a_hoje-2])
        
        if c_btn.button("🔄 Processar Proventos do Mês", use_container_width=True):
            with st.spinner("Buscando agenda de pagamentos na B3..."):
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
                    yoc = (rec / (dm['Qtd'] * dm['PM'])) * 100 if dm['PM'] > 0 else 0
                    
                    if dm['Tipo'] == 'FII':
                        lf.append({"Fundo (FII)": t_tk, "Unitário (R$)": val, "Recebido (R$)": rec, "Yield on Cost (%)": yoc, "Status": "Divulgado / Pago 🟢" if val > 0 else "Aguardando 🟡"})
                    else:
                        if val > 0: la.append({"Ação": t_tk, "Unitário (R$)": val, "Recebido (R$)": rec, "Yield on Cost (%)": yoc, "Status": "Pago 🟢"})
                
                st.session_state.divs_a, st.session_state.divs_f = pd.DataFrame(la), pd.DataFrame(lf)
                st.session_state.divs_m, st.session_state.divs_ano = m_sel, a_sel
        
        if ('divs_a' in st.session_state and not st.session_state.divs_a.empty) or ('divs_f' in st.session_state and not st.session_state.divs_f.empty):
            tot_mes = 0
            if 'divs_f' in st.session_state and not st.session_state.divs_f.empty:
                st.markdown("#### 🏢 Status dos FIIs")
                st.dataframe(st.session_state.divs_f.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct}), use_container_width=True, hide_index=True)
                tot_mes += st.session_state.divs_f['Recebido (R$)'].sum()
            if 'divs_a' in st.session_state and not st.session_state.divs_a.empty:
                st.markdown("#### 📈 Ações Pagadoras")
                st.dataframe(st.session_state.divs_a.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct}), use_container_width=True, hide_index=True)
                tot_mes += st.session_state.divs_a['Recebido (R$)'].sum()
            st.success(f"**💰 Total Estimado ({m_map[st.session_state.divs_m]}/{st.session_state.divs_ano}):** {f_brl(tot_mes)}")
            
        st.markdown("---")
        st.markdown("### 🏛️ Histórico Analítico de Proventos")
        l_hist = []
        for t_hist, dm_hist in st.session_state.dados_mercado.items():
            try:
                divs_h = yf.Ticker(f"{t_hist}.SA").dividends
                if not divs_h.empty:
                    if divs_h.index.tz is not None: divs_h.index = divs_h.index.tz_localize(None)
                    divs_fil = divs_h[divs_h.index >= pd.Timestamp(dm_hist['Data'])]
                    for d_idx, val_h in divs_fil.items():
                        t_rec = val_h * dm_hist['Qtd']
                        inv_h = dm_hist['Qtd'] * dm_hist['PM']
                        l_hist.append({"Data Ex": d_idx.date(), "Ativo": t_hist, "Unitário (R$)": float(val_h), "Quantidade": int(dm_hist['Qtd']), "Recebido (R$)": float(t_rec), "Yield on Cost (%)": float((t_rec / inv_h)*100 if inv_h>0 else 0), "DY Atual (%)": float((val_h / dm_hist['Preço Atual'])*100 if dm_hist['Preço Atual']>0 else 0)})
            except: pass
            
        if l_hist:
            df_hist_tot = pd.DataFrame(l_hist).sort_values("Data Ex", ascending=False)
            c_h1, c_h2 = st.columns(2)
            atvs_disp = sorted(df_hist_tot['Ativo'].unique().tolist())
            atvs_sel = c_h1.multiselect("Filtrar por Ativo:", options=atvs_disp, default=atvs_disp)
            r_hist = c_h2.date_input("Filtrar por Período:", value=(min(df_hist_tot['Data Ex']), max(df_hist_tot['Data Ex'])))
            
            df_hist_f = df_hist_tot[df_hist_tot['Ativo'].isin(atvs_sel)]
            if isinstance(r_hist, tuple) and len(r_hist) == 2:
                df_hist_f = df_hist_f[(df_hist_f['Data Ex'] >= r_hist[0]) & (df_hist_f['Data Ex'] <= r_hist[1])]
                
            if not df_hist_f.empty:
                st.dataframe(df_hist_f.style.format({"Unitário (R$)": f_brl_4, "Recebido (R$)": f_brl, "Yield on Cost (%)": f_pct, "DY Atual (%)": f_pct}), use_container_width=True, hide_index=True)
                st.download_button(label="📥 Baixar Histórico (Excel)", data=to_excel(df_hist_f, sheet_name="Historico_Proventos"), file_name=f"Historico_Proventos_{st.session_state.username}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

else:
    for tb in [t1, t2, t3, t4, t5]:
        with tb: st.info("ℹ️ Adicione ativos na tabela manual ou via planilha e clique em **Conectar ao Mercado Vivo**.")

# ==========================================
# 8. ABAS ISOLADAS (TESOURO E IA)
# ==========================================
with t_tes:
    st.markdown("### 🏛️ Simulador Tesouro Direto")
    st.info("Insira títulos de renda fixa manualmente para projetar juros compostos.")
    if st.session_state.df_tesouro.empty: st.session_state.df_tesouro = pd.DataFrame([{"Título": "Tesouro IPCA+ 2029", "Valor Investido (R$)": 1000.0, "Taxa Anual (%)": 6.0, "Ano Venc.": 2029}])
    df_t = st.data_editor(st.session_state.df_tesouro, num_rows="dynamic", use_container_width=True, hide_index=True)
    st.session_state.df_tesouro = df_t
    if st.button("Projetar Juros até Vencimento", type="primary"):
        res_t = []
        for _, rt in df_t.iterrows():
            anos = max(1, int(rt['Ano Venc.']) - pd.Timestamp.now().year)
            v_final = float(rt['Valor Investido (R$)']) * ((1 + (float(rt['Taxa Anual (%)'])/100)) ** anos)
            res_t.append({"Título": rt['Título'], "Anos P/ Vencer": anos, "Investido": float(rt['Valor Investido (R$)']), "Valor Bruto Vencimento": v_final, "Lucro Bruto Projetado": v_final - float(rt['Valor Investido (R$)'])})
        st.dataframe(pd.DataFrame(res_t).style.format({"Investido": f_brl, "Valor Bruto Vencimento": f_brl, "Lucro Bruto Projetado": f_brl}), use_container_width=True, hide_index=True)

with t6:
    st.markdown("### 💬 Comitê de IA Sênior")
    cb1, cb2 = st.columns(2)
    if cb1.button("🗑️ Limpar Chat", use_container_width=True):
        st.session_state.historico_chat = [{"role": "assistant", "content": f"Saudações, {st.session_state.username}. Como posso ajudar?"}]
        st.rerun()
        
    if HAS_DOCX and len(st.session_state.historico_chat) > 1:
        cb2.download_button("📄 Exportar Parecer (DOCX)", data=export_docx(st.session_state.historico_chat), file_name=f"Analise_IA_{st.session_state.username}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    elif not HAS_DOCX: cb2.caption("⚠️ Instale 'python-docx' para exportação em Word.")

    api_key_secreta = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key_secreta: api_key_secreta = st.text_input("Insira sua Gemini API Key:", type="password")
        
    for m in st.session_state.historico_chat:
        with st.chat_message(m["role"]): st.write(m["content"])
        
    if prompt := st.chat_input("Ex: 'A BBAS3 está barata?'"):
        with st.chat_message("user"): st.write(prompt)
        st.session_state.historico_chat.append({"role": "user", "content": prompt})
        
        with st.chat_message("assistant"):
            with st.spinner("Analisando mercado e histórico..."):
                ctx_c = str(st.session_state.dados_mercado) if st.session_state.dados_mercado else "Usuário sem carteira cadastrada."
                ctx_m = f"Selic: {f_pct(selic_hoje)}|IPCA: {f_pct(ipca_12m_hoje)}. Focus {ano_atual}: Selic {f_pct(proj_focus.get(f'Selic_{ano_atual}'))}/IPCA {f_pct(proj_focus.get(f'IPCA_{ano_atual}'))}"
                h_txt = "\n".join([f"{'Usuário' if h['role']=='user' else 'Gestora IA'}: {h['content']}" for h in st.session_state.historico_chat[-6:-1]])
                sys_p = (f"Analista CNPI Sênior. [Carteira Atual]: {ctx_c}. [Cenário Macro]: {ctx_m}.\nDIRETRIZ DE CONTINUIDADE: Use o HISTÓRICO abaixo.\nREGRA ESTRITA: 1) Se o usuário citar 'carteira' ou ativos da carteira, analise o portfólio dele. 2) Se pergunta for geral, IGNORE a carteira e responda neutro.\n=== HISTÓRICO ===\n{h_txt}")
                
                resposta_ia = "⚠️ Chave API ausente."
                if api_key_secreta:
                    try:
                        import google.generativeai as genai
                        genai.configure(api_key=api_key_secreta)
                        sucesso_ia, erro_log = False, ""
                        for mdl in ['gemini-2.5-flash', 'gemini-1.5-flash']:
                            try:
                                resposta_ia = genai.GenerativeModel(mdl).generate_content([sys_p, prompt]).text
                                sucesso_ia = True; break 
                            except Exception as e_ia: erro_log = str(e_ia); continue
                        if not sucesso_ia: resposta_ia = f"⚠️ Falha de rede com a IA: {erro_log}"
                    except Exception as e_m: resposta_ia = f"⚠️ Erro de importação/motor: {e_m}"
                st.write(resposta_ia)
                
        st.session_state.historico_chat.append({"role": "assistant", "content": resposta_ia})
        salvar_dados_completos_db(st.session_state.username) 
        st.rerun()
        
