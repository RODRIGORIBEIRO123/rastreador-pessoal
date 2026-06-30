import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re
import io
import requests
import plotly.express as px

# ==========================================
# 1. CONFIGURAÇÃO E FORMATADORES DE DADOS
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")

data_hoje = pd.Timestamp.now().strftime('%d/%m/%Y')
st.title(f"📊 Terminal de Gestão Profissional - {data_hoje}")

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

if 'df_base' not in st.session_state: st.session_state.df_base = pd.DataFrame()
if 'dados_mercado' not in st.session_state: st.session_state.dados_mercado = {}
if 'df_simul' not in st.session_state: st.session_state.df_simul = pd.DataFrame()

if 'historico_chat' not in st.session_state:
    st.session_state.historico_chat = [{"role": "assistant", "content": "Saudações. O terminal está mapeado em tempo real. Pode fazer a sua pergunta sobre o cenário macroeconômico ou a sua carteira."}]

# ==========================================
# 2. FUNÇÕES MACRO E FUNDAMENTOS
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
    fallback = {
        f"IPCA_{ano_atual}": 3.80, f"Selic_{ano_atual}": selic_atual, 
        f"IPCA_{ano_atual+1}": 3.70, f"Selic_{ano_atual+1}": selic_atual - 1.0,
        f"IPCA_{ano_atual+2}": 3.50, f"Selic_{ano_atual+2}": selic_atual - 1.5
    }
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
    except: 
        return fallback, ano_atual

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
    mapa = {
        "Banks": "Bancos", "Utilities - Regulated Electric": "Energia", 
        "Real Estate - Retail": "Shoppings/Varejo", "REIT - Retail": "Shoppings/Varejo",
        "Real Estate - Industrial": "Logística", "REIT - Industrial": "Logística",
        "REIT - Office": "Lajes Corporativas", "REIT - Diversified": "Fundo Híbrido",
        "Financial Data & Stock Exchanges": "Bolsa de Valores", "Insurance": "Seguradoras",
        "Oil & Gas Integrated": "Petróleo e Gás"
    }
    return mapa.get(setor_en, "Outros Setores")

# ==========================================
# 3. LÓGICA DE PROCESSAMENTO DE ATIVOS
# ==========================================
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
            if qtd >= (posicoes[ticker]['qtd'] - 0.001): 
                posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'ts_medio': 0.0}
            else:
                pm = posicoes[ticker]['valor'] / posicoes[ticker]['qtd'] if posicoes[ticker]['qtd'] > 0 else 0
                posicoes[ticker]['qtd'] -= qtd
                posicoes[ticker]['valor'] -= (qtd * pm)

    ativos_limpos = [{"Ativo": t, "Quantidade": d['qtd'], "Preço Médio": d['valor']/d['qtd'] if d['qtd'] > 0 else 0, "Data Média": pd.to_datetime(d['ts_medio'], unit='s').date()} for t, d in posicoes.items() if d['qtd'] > 0]
    return consolidar_carteira(pd.DataFrame(ativos_limpos))

# ==========================================
# 4. SISTEMA DE UPLOAD DE DADOS
# ==========================================
st.sidebar.header("1. Upload de Arquivos")
arquivo_principal = st.sidebar.file_uploader("1️⃣ Planilha Principal (Opcional)", type=["xlsx", "csv"])
arquivo_novo = st.sidebar.file_uploader("2️⃣ Novas Negociações B3 (Opcional)", type=["xlsx", "csv"])

data_corte = None
if arquivo_novo:
    data_corte = st.sidebar.date_input("Filtrar Novas Transações a partir de:", pd.Timestamp.now().date() - pd.Timedelta(days=15))

if st.sidebar.button("🚀 Carregar e Rodar Aplicativo", use_container_width=True):
    if arquivo_principal:
        with st.spinner("Processando e cruzando dados..."):
            df_prin = ler_arquivo_universal(arquivo_principal)
            if 'Data Média' in df_prin.columns:
                df_prin['Data Média'] = pd.to_datetime(df_prin['Data Média'], errors='coerce').dt.date
                base_atual = consolidar_carteira(df_prin)
            else:
                base_atual = processar_planilha_b3(df_prin)
                
            if arquivo_novo and not base_atual.empty:
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
            st.rerun()

# ==========================================
# 5. PAINEL MACRO (SEMPRE VISÍVEL)
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
# 6. BANCO DE DADOS E CONEXÃO B3
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Controle do Banco de Dados")
    c_a, c_b, c_c = st.columns([1, 1, 1])
    with c_a:
        tdel = st.selectbox("Excluir Ativo:", [""] + sorted(st.session_state.df_base["Ativo"].tolist()))
        if st.button("Remover", use_container_width=True) and tdel != "":
            st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != tdel]
            st.rerun()
    with c_b:
        nt = st.text_input("Nova Compra (Ticker)")
        ncq, ncp = st.columns(2)
        nq, np = ncq.number_input("Qtd", min_value=1), ncp.number_input("PM (R$)", min_value=0.01)
        if st.button("Adicionar", use_container_width=True) and nt != "":
            nl = pd.DataFrame([{"Ativo": nt.upper(), "Quantidade": float(nq), "Preço Médio": float(np), "Data Média": pd.Timestamp.now().date()}])
            st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, nl], ignore_index=True))
            st.rerun()
    with c_c:
        st.download_button("📥 Salvar Backup Diário (.csv)", data=st.session_state.df_base.to_csv(index=False, sep=';', encoding='utf-8-sig'), file_name="Banco_de_Dados_Carteira.csv", use_container_width=True)

    df_editado = st.data_editor(st.session_state.df_base, use_container_width=True, hide_index=True, column_config={"Ativo": st.column_config.TextColumn("Ativo", disabled=False), "Quantidade": st.column_config.NumberColumn(min_value=0.0), "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0), "Data Média": st.column_config.DateColumn("Data Média Ponderada", format="DD/MM/YYYY")})

    if st.button("🚀 Processar Conexão com o Mercado", type="primary"):
        st.session_state.df_base = consolidar_carteira(df_editado) 
        df_macro, fundamentos_br = carregar_macro(), obter_fundamentos_brasil()
        progresso, total = st.progress(0), len(st.session_state.df_base)
        dados_mercado, linhas_simul_iniciais = {}, []

        for i, row in st.session_state.df_base.iterrows():
            ticker = str(row['Ativo']).strip().upper()
            data_compra = pd.to_datetime(row['Data Média']) if pd.notna(row['Data Média']) else pd.Timestamp.now()
            preco_atual = float(row['Preço Médio'])
            divs_total, divs_12m, lpa, vpa, setor = 0.0, 0.0, 0.0, 0.0, "Desconhecido"
            tipo_ativo = "FII" if ticker.endswith('11') and ticker not in UNITS_ACOES else "Ação"
            
            try:
                acao = yf.Ticker(f"{ticker}.SA")
                try: 
                    hist = acao.history(period="1d")
                    if not hist.empty: preco_atual = float(hist['Close'].iloc[-1])
                except: pass
                try: 
                    divs = acao.dividends
                    divs_total = float(divs[divs.index.tz_localize(None) >= data_compra].sum() * row['Quantidade'])
                    divs_12m = float(divs[divs.index.tz_localize(None) >= (pd.Timestamp.now() - pd.DateOffset(years=1))].sum())
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
        st.success("Conexão Estabelecida com Sucesso!")

# ==========================================
# 7. DASHBOARD E PAINEL DE RELATÓRIOS
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
                "Data Média": dm['Data'].strftime('%d/%m/%Y'), "Meses (Média)": int(calcular_meses(dm['Data'])),
                "Total Div. (R$)": dm['Div_Total'], "DY on Cost (%)": yoc, "Evolução c/ Div (%)": var_c_div, "IPCA Acum. (%)": dm['IPCA'], "CDI Acum. (%)": dm['CDI']
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

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Visão Geral", "💰 Valuation (Bazin & Graham)", "⚖️ Pesos e Setores", "📈 Gráficos", "💸 Proventos (Filtro Mensal)", "💬 IA & Recomendações"
        ])
        
        with tab1:
            styled_perf = df_perf_final.drop(columns=['Tipo', 'Setor']).style.format({
                "Preço Médio": f_brl, "Preço Atual": f_brl, "Total Investido": f_brl, 
                "Saldo Atual": f_brl, "Resultado (R$)": f_brl, "Total Div. (R$)": f_brl,
                "DY on Cost (%)": f_pct, "Evolução c/ Div (%)": f_pct, "IPCA Acum. (%)": f_pct, "CDI Acum. (%)": f_pct
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
            
            st.session_state.df_rec_bazin = pd.DataFrame(linhas_bazin)
            styled_bazin = st.session_state.df_rec_bazin.style.format({"Preço Teto (Bazin)": f_brl, "Margem Segurança (%)": f_pct})
            st.dataframe(styled_bazin, use_container_width=True, hide_index=True)
            
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
            
            st.session_state.df_rec_graham = pd.DataFrame(linhas_graham)
            styled_graham = st.session_state.df_rec_graham.style.format({"Preço Justo (Graham)": f_brl, "Margem Segurança (%)": f_pct})
            st.dataframe(styled_graham, use_container_width=True, hide_index=True)

        with tab3:
            c_g1, c_g2, c_g3 = st.columns(3)
            c_g1.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Tipo', hole=0.4, title="Por Tipo"), use_container_width=True)
            c_g2.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Ativo', title="Por Ativo"), use_container_width=True)
            c_g3.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Setor', title="Por Setor"), use_container_width=True)

        with tab4:
            todos_ativos = df_perf_final['Ativo'].tolist()
            c_sel, c_ind = st.columns([2, 1])
            with c_sel: ativos_selecionados = st.multiselect("Selecione os ativos:", todos_ativos, default=todos_ativos[:6], key="ms_g")
            with c_ind: ind_selecionados = st.multiselect("Indicadores:", ["Evolução c/ Div (%)", "CDI Acum. (%)", "IPCA Acum. (%)"], default=["Evolução c/ Div (%)", "CDI Acum. (%)", "IPCA Acum. (%)"], key="ind_g")
            
            if ativos_selecionados and ind_selecionados:
                df_grafico = df_perf_final[df_perf_final['Ativo'].isin(ativos_selecionados)].copy()
                df_grafico['Período'] = df_grafico['Data Média'].astype(str) + " até Hoje"
                df_melt = df_grafico.melt(id_vars=["Ativo", "Período"], value_vars=ind_selecionados, var_name="Indicador", value_name="Rentabilidade")
                fig1 = px.bar(df_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group")
                fig1.update_layout(yaxis_ticksuffix=" %", margin=dict(t=40))
                st.plotly_chart(fig1, use_container_width=True)

        with tab5:
            st.markdown("### 💸 Filtro e Histórico de Proventos")
            
            c_f1, c_f2, c_btn = st.columns([2, 2, 2])
            meses_map = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
            mes_atual = pd.Timestamp.now().month
            ano_atual_data = pd.Timestamp.now().year
            
            mes_selecionado = c_f1.selectbox("Mês de Referência:", options=list(meses_map.keys()), format_func=lambda x: meses_map[x], index=mes_atual-1)
            ano_selecionado = c_f2.selectbox("Ano de Referência:", options=[ano_atual_data, ano_atual_data-1, ano_atual_data-2, ano_atual_data-3, ano_atual_data-4])
            
            if 'divs_calculados' not in st.session_state:
                st.session_state.divs_calculados = None
                st.session_state.divs_mes = mes_atual
                st.session_state.divs_ano = ano_atual_data

            if c_btn.button("🔄 Processar Dividendos do Período", use_container_width=True) or st.session_state.divs_calculados is None:
                st.session_state.divs_mes = mes_selecionado
                st.session_state.divs_ano = ano_selecionado
                
                with st.spinner(f"Varrendo histórico de pagamentos de {meses_map[mes_selecionado]}/{ano_selecionado}..."):
                    linhas_div = []
                    for t, dm in st.session_state.dados_mercado.items():
                        try:
                            acao = yf.Ticker(f"{t}.SA")
                            divs = acao.dividends
                            if not divs.empty:
                                if divs.index.tz is not None: divs.index = divs.index.tz_localize(None)
                                divs_mes = divs[(divs.index.month == st.session_state.divs_mes) & (divs.index.year == st.session_state.divs_ano)].sum()
                                
                                if divs_mes > 0:
                                    yoc = ((divs_mes * dm['Qtd']) / (dm['Qtd'] * dm['PM'])) * 100 if dm['PM'] > 0 else 0
                                    dy_spot = (divs_mes / dm['Preço Atual']) * 100 if dm['Preço Atual'] > 0 else 0
                                    
                                    linhas_div.append({
                                        "Ativo": t, 
                                        "Provento Unit. (R$)": float(divs_mes),
                                        "Qtd Detida": int(dm['Qtd']),
                                        "Total Recebido (R$)": float(divs_mes * dm['Qtd']),
                                        "Yield on Cost (%)": float(yoc),
                                        "DY Atual (%)": float(dy_spot)
                                    })
                        except: continue
                    st.session_state.divs_calculados = pd.DataFrame(linhas_div)
            
            df_divs = st.session_state.divs_calculados
            if df_divs is not None and not df_divs.empty:
                styled_df = df_divs.style.format({
                    "Provento Unit. (R$)": f_brl_4,
                    "Total Recebido (R$)": f_brl,
                    "Yield on Cost (%)": f_pct,
                    "DY Atual (%)": f_pct
                })
                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                total_br = f_brl(df_divs['Total Recebido (R$)'].sum())
                st.success(f"**Total Recebido em {meses_map[st.session_state.divs_mes]}/{st.session_state.divs_ano}:** {total_br}")
            elif df_divs is not None and df_divs.empty:
                st.info(f"A sua base de ativos não regista proventos para {meses_map[st.session_state.divs_mes]}/{st.session_state.divs_ano}.")

        with tab6:
            st.markdown("### 🤖 Radar e Recomendações")
            st.info("Utilize a aba inferior (Comitê de IA) para solicitar análises avançadas de balanço.")

# ==========================================
# 8. COMITÊ DE IA E CHAT (SEMPRE VISÍVEL)
# ==========================================
st.write("---")
st.markdown("### 💬 Comitê de Alocação IA - Visão CNPI Sênior")

api_key = ""
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except:
    api_key = st.text_input("Chave API do Google Gemini:", type="password")

for msg in st.session_state.historico_chat:
    with st.chat_message(msg["role"]): st.write(msg["content"])
    
if prompt := st.chat_input("Ex: Com a Selic atual, devo aportar em Renda Fixa ou FIIs?"):
    st.session_state.historico_chat.append({"role": "user", "content": prompt})
    with st.chat_message("user"): st.write(prompt)
    
    ctx_carteira = "Nenhuma carteira carregada."
    if not st.session_state.df_base.empty and 'df_perf_final' in locals():
        ctx_carteira = df_perf_final[['Ativo', 'Qtd', 'Preço Médio', 'Preço Atual', 'Total Investido', 'Saldo Atual', 'Resultado (R$)', 'Total Div. (R$)', 'Evolução c/ Div (%)']].to_csv(index=False, sep='|')
        
    ctx_macro = (f"Selic Vigente Hoje: {f_pct(selic_hoje)} | IPCA Acum. 12 meses: {f_pct(ipca_12m_hoje)}\n"
                 f"Projeções Selic Fim de Ano: {ano_atual}: {f_pct(proj_focus.get(f'Selic_{ano_atual}'))} | {ano_atual+1}: {f_pct(proj_focus.get(f'Selic_{ano_atual+1}'))} | {ano_atual+2}: {f_pct(proj_focus.get(f'Selic_{ano_atual+2}'))}\n"
                 f"Projeções IPCA Fim de Ano: {ano_atual}: {f_pct(proj_focus.get(f'IPCA_{ano_atual}'))} | {ano_atual+1}: {f_pct(proj_focus.get(f'IPCA_{ano_atual+1}'))} | {ano_atual+2}: {f_pct(proj_focus.get(f'IPCA_{ano_atual+2}'))}")
    
    sys_prompt = f"""
    Você é uma Analista de Investimentos Sênior (CNPI) e Gestora de Portfólio.
    Sua linguagem é corporativa, executiva, fria, baseada em dados e voltada a relatórios de alta governança.
    
    [MATRIZ DE DADOS EM TEMPO REAL]
    {ctx_carteira}
    
    [CENÁRIO MACRO SPOT E FORWARD 3 ANOS]
    {ctx_macro}
    
    [DIRETRIZES DE RESPOSTA]
    - Não use linguagem de iniciante. Use termos técnicos apropriados.
    - Se a carteira estiver vazia, foque inteiramente na análise do cenário Macro.
    - Cruze os dados disponíveis de forma exata e matemática.
    """
    
    if api_key:
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            modelos_para_tentar = ['gemini-3.5-flash', 'gemini-3.1-flash-lite', 'gemini-2.5-flash']
            resposta_sucesso = False
            erros_tecnicos = []
            
            for nome_modelo in modelos_para_tentar:
                try:
                    model = genai.GenerativeModel(nome_modelo)
                    response = model.generate_content([sys_prompt, prompt])
                    resposta = response.text
                    resposta_sucesso = True
                    break 
                except Exception as e:
                    erros_tecnicos.append(f"{nome_modelo}: {str(e)}")
                    continue
                        
            if not resposta_sucesso:
                erro_formatado = "\n".join(erros_tecnicos)
                resposta = f"⚠️ Falha de comunicação com a IA.\n\nLogs:\n`{erro_formatado}`"
        except Exception as e:
            resposta = f"⚠️ Erro estrutural grave: {str(e)}"
    else:
        resposta = "⚠️ Operação bloqueada. Insira a chave da API do Gemini."
    
    st.session_state.historico_chat.append({"role": "assistant", "content": resposta})
    st.rerun()
