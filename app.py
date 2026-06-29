import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re
import io
import requests
from openpyxl.chart import BarChart, Reference
import plotly.express as px

# ==========================================
# 1. CONFIGURAÇÃO, MEMÓRIA E DICIONÁRIOS
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")
st.title("📊 Terminal de Gestão Profissional")

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
    st.session_state.historico_chat = [{"role": "assistant", "content": "Saudações. Sou a sua analista sênior integrada. O terminal foi reestruturado para mapear seus ativos e os indicadores do Focus em tempo real. Insira sua chave API corporativa para habilitar a inteligência profunda."}]

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
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}
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
def obter_projecoes_focus():
    ano_atual = pd.Timestamp.now().year
    fallback = {f"IPCA_{ano_atual}": 5.33, f"Selic_{ano_atual}": 14.00, f"IPCA_{ano_atual+1}": 4.15, f"Selic_{ano_atual+1}": 12.00}
    try:
        url = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$top=100&$filter=Indicador%20eq%20'IPCA'%20or%20Indicador%20eq%20'Selic'&$orderby=Data%20desc&$format=json"
        res = requests.get(url, timeout=5).json()
        if 'value' in res and len(res['value']) > 0:
            df = pd.DataFrame(res['value'])
            data_recente = df['Data'].max()
            df = df[df['Data'] == data_recente]
            df_atual = df[df['DataReferencia'] == str(ano_atual)]
            df_prox = df[df['DataReferencia'] == str(ano_atual+1)]
            if not df_atual[df_atual['Indicador'] == 'IPCA'].empty: fallback[f"IPCA_{ano_atual}"] = float(df_atual[df_atual['Indicador'] == 'IPCA']['Mediana'].values[0])
            if not df_atual[df_atual['Indicador'] == 'Selic'].empty: fallback[f"Selic_{ano_atual}"] = float(df_atual[df_atual['Indicador'] == 'Selic']['Mediana'].values[0])
            if not df_prox[df_prox['Indicador'] == 'IPCA'].empty: fallback[f"IPCA_{ano_atual+1}"] = float(df_prox[df_prox['Indicador'] == 'IPCA']['Mediana'].values[0])
            if not df_prox[df_prox['Indicador'] == 'Selic'].empty: fallback[f"Selic_{ano_atual+1}"] = float(df_prox[df_prox['Indicador'] == 'Selic']['Mediana'].values[0])
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
    mapa = {
        "Banks": "Bancos", "Utilities - Regulated Electric": "Energia", 
        "Real Estate - Retail": "Shoppings/Varejo", "REIT - Retail": "Shoppings/Varejo",
        "Real Estate - Industrial": "Logística", "REIT - Industrial": "Logística",
        "REIT - Office": "Lajes Corporativas", "REIT - Diversified": "Fundo Híbrido",
        "Financial Data & Stock Exchanges": "Bolsa de Valores", "Insurance": "Seguradoras",
        "Oil & Gas Integrated": "Petróleo e Gás"
    }
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

def gerar_excel_premium(df_perf, df_val):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_perf.fillna(0).to_excel(writer, sheet_name='Rentabilidade', index=False)
        df_val.fillna(0).to_excel(writer, sheet_name='Valuation', index=False)
    return output.getvalue()

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
# 2. SISTEMA DE UPLOAD COM BOTÃO EXPLÍCITO
# ==========================================
st.sidebar.header("1. Upload de Arquivos")
arquivo_principal = st.sidebar.file_uploader("1️⃣ Planilha Principal (B3 ou Backup)", type=["xlsx", "csv"])
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
# 3. INTERFACE DE CONTROLE
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
# 4. PAINEL DE RELATÓRIOS E DASHBOARD
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
                "Total Div. (R$)": dm['Div_Total'], "DY on Cost": yoc, "Evolução c/ Div": var_c_div, "IPCA Acum.": dm['IPCA'], "CDI Acum.": dm['CDI']
            })
        df_perf_final = pd.DataFrame(linhas_perf)

        st.write("---")
        st.markdown("### 🏆 Visão Global do Portfólio")
        data_mais_antiga = st.session_state.df_base['Data Média'].min()
        data_formatada = data_mais_antiga.strftime('%d/%m/%Y') if pd.notna(data_mais_antiga) else "Início"
        st.caption(f"🗓️ **Período de Análise Base:** Desde {data_formatada} até Hoje.")

        df_acoes = df_perf_final[df_perf_final['Tipo'] == 'Ação']
        df_fiis = df_perf_final[df_perf_final['Tipo'] == 'FII']
        evolucao_acoes = (df_acoes['Saldo Atual'].sum() / df_acoes['Total Investido'].sum() - 1) * 100 if df_acoes['Total Investido'].sum() > 0 else 0
        evolucao_fiis = (df_fiis['Saldo Atual'].sum() / df_fiis['Total Investido'].sum() - 1) * 100 if df_fiis['Total Investido'].sum() > 0 else 0
        tooltip_msg = f"Período considerado: Desde {data_formatada} até Hoje."

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📈 Patrimônio (Ações)", f"R$ {df_acoes['Saldo Atual'].sum():,.2f}", f"{evolucao_acoes:.2f}% (R$ {df_acoes['Resultado (R$)'].sum():,.2f})", help=tooltip_msg)
        m2.metric("🏢 Patrimônio (FIIs)", f"R$ {df_fiis['Saldo Atual'].sum():,.2f}", f"{evolucao_fiis:.2f}% (R$ {df_fiis['Resultado (R$)'].sum():,.2f})", help=tooltip_msg)
        m3.metric("💸 Renda Acumulada (Ações)", f"R$ {df_acoes['Total Div. (R$)'].sum():,.2f}", help=tooltip_msg)
        m4.metric("💸 Renda Acumulada (FIIs)", f"R$ {df_fiis['Total Div. (R$)'].sum():,.2f}", help=tooltip_msg)

        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📊 Visão Geral", "💰 Bazin (Renda)", "🏢 Graham (Valor)", 
            "⚖️ Pesos e Setores", "🎯 Recomendações e Projeções", "📈 Gráficos Interativos", "💬 Assistente IA"
        ])
        
        with tab1:
            st.dataframe(df_perf_final.drop(columns=['Tipo', 'Setor']), use_container_width=True, hide_index=True)

        with tab2:
            yield_desejado = st.number_input("Taxa de Risco Exigida (%):", value=6.0, step=0.5) / 100.0
            df_bazin_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Proj. (R$)"] if "Div. Proj. (R$)" in st.session_state.df_simul.columns else ["Ativo", "Cotação Atual", "Div. Projetado (R$)"]].copy()
            df_bazin_editado = st.data_editor(df_bazin_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"], key="edit_bazin")
            st.session_state.df_simul["Div. Projetado (R$)"] = df_bazin_editado.iloc[:, 2]
            
            linhas_bazin = []
            for _, row in df_bazin_editado.iterrows():
                bazin = (row.iloc[2] / yield_desejado) if (row.iloc[2] > 0 and yield_desejado > 0) else 0.0
                margem_b = (((bazin / row['Cotação Atual']) - 1) * 100) if (bazin > 0 and row['Cotação Atual'] > 0) else 0.0
                linhas_bazin.append({"Ativo": row['Ativo'], "Preço Teto (Bazin)": bazin, "Margem Segurança": margem_b})
            st.session_state.df_rec_bazin = pd.DataFrame(linhas_bazin)
            st.dataframe(st.session_state.df_rec_bazin, use_container_width=True, hide_index=True)

        with tab3:
            df_graham_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "VPA (Contábil)", "LPA Projetado"]].copy()
            df_graham_editado = st.data_editor(df_graham_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"], key="edit_graham")
            st.session_state.df_simul["VPA (Contábil)"] = df_graham_editado["VPA (Contábil)"]
            st.session_state.df_simul["LPA Projetado"] = df_graham_editado["LPA Projetado"]
            
            linhas_graham = []
            for _, row in df_graham_editado.iterrows():
                graham = (22.5 * row['LPA Projetado'] * row['VPA (Contábil)']) ** 0.5 if (row['LPA Projetado'] > 0 and row['VPA (Contábil)'] > 0) else 0.0
                margem_g = (((graham / row['Cotação Atual']) - 1) * 100) if (graham > 0 and row['Cotação Atual'] > 0) else 0.0
                linhas_graham.append({"Ativo": row['Ativo'], "Preço Justo (Graham)": graham, "Margem Segurança": margem_g})
            st.session_state.df_rec_graham = pd.DataFrame(linhas_graham)
            st.dataframe(st.session_state.df_rec_graham, use_container_width=True, hide_index=True)

        with tab4:
            c_g1, c_g2, c_g3 = st.columns(3)
            c_g1.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Tipo', hole=0.4, title="Por Tipo"), use_container_width=True)
            c_g2.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Ativo', title="Por Ativo"), use_container_width=True)
            c_g3.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Setor', title="Por Classe/Setor"), use_container_width=True)
            
            def gerar_tabela_peso(coluna):
                df_peso = df_perf_final.groupby(coluna)['Saldo Atual'].sum().reset_index()
                df_peso['Peso (%)'] = (df_peso['Saldo Atual'] / df_peso['Saldo Atual'].sum()) * 100
                return df_peso.sort_values('Peso (%)', ascending=False)
            c_t1, c_t2, c_t3 = st.columns(3)
            with c_t1: st.dataframe(gerar_tabela_peso('Tipo'), use_container_width=True, hide_index=True)
            with c_t2: st.dataframe(gerar_tabela_peso('Ativo'), use_container_width=True, hide_index=True)
            with c_t3: st.dataframe(gerar_tabela_peso('Setor'), use_container_width=True, hide_index=True)

        with tab5:
            proj_focus, ano_atual = obter_projecoes_focus()
            st.markdown(f"### 🇧🇷 Projeções Macroeconômicas (Boletim Focus Bacen)")
            st.info(f"**IPCA {ano_atual}:** {proj_focus.get(f'IPCA_{ano_atual}', 5.33)}%  |  **Selic {ano_atual}:** {proj_focus.get(f'Selic_{ano_atual}', 14.00)}%  ||  **IPCA {ano_atual+1}:** {proj_focus.get(f'IPCA_{ano_atual+1}', 4.15)}%  |  **Selic {ano_atual+1}:** {proj_focus.get(f'Selic_{ano_atual+1}', 12.00)}%")
            
            st.markdown("### 🤖 Radar de Oportunidades do Especialista")
            
            c_p1, c_p2, c_p3, c_p4 = st.columns(4)
            patrimonio_fora = c_p1.number_input("Patrimônio Fora da Bolsa (R$):", value=0.0, step=10000.0)
            aporte_mensal_planejado = c_p2.number_input("Aporte Mensal (R$):", value=2000.0, step=500.0)
            rentabilidade_ganho = c_p3.number_input("Rentab. Mensal Estimada (%):", value=0.8, step=0.1) / 100.0
            cresc_dividendos_anual = c_p4.number_input("Crescimento Anual de Dividendos (%):", value=5.0, step=1.0) / 100.0

            df_recs = pd.merge(df_perf_final[['Ativo', 'Tipo']], st.session_state.df_rec_bazin[['Ativo', 'Margem Segurança']], on='Ativo', how='left').rename(columns={'Margem Segurança': 'Margem Bazin (%)'})
            df_recs = pd.merge(df_recs, st.session_state.df_rec_graham[['Ativo', 'Margem Segurança']], on='Ativo', how='left').rename(columns={'Margem Segurança': 'Margem Graham (%)'})
            
            st.markdown("##### ⚖️ Parametrização Exigida do Prêmio de Risco")
            c_m1, c_m2 = st.columns(2)
            margem_graham_exigida = c_m1.number_input("Margem de Segurança Mínima de Graham (Ações %):", value=15.0, step=1.0)
            margem_bazin_exigida = c_m2.number_input("Margem de Segurança Mínima de Bazin (FIIs/Renda %):", value=5.0, step=1.0)

            recomendacoes = []
            for _, r in df_recs.iterrows():
                if r['Tipo'] == 'Ação':
                    if r['Margem Graham (%)'] > margem_graham_exigida and r['Margem Bazin (%)'] > margem_bazin_exigida: recomendacoes.append("COMPRA FORTE 🟢")
                    elif r['Margem Graham (%)'] > 0 or r['Margem Bazin (%)'] > 0: recomendacoes.append("MANTER / COMPRA 🟡")
                    else: recomendacoes.append("AVALIAR VENDA 🔴")
                else: 
                    if r['Margem Bazin (%)'] > margem_bazin_exigida: recomendacoes.append("COMPRA FORTE 🟢")
                    elif r['Margem Bazin (%)'] > -5: recomendacoes.append("MANTER 🟡")
                    else: recomendacoes.append("AVALIAR VENDA 🔴")
            df_recs['Status Recomendações'] = recomendacoes
            
            st.dataframe(df_recs, use_container_width=True, hide_index=True, column_config={
                "Status Recomendações": st.column_config.TextColumn("Status Recomendações ❓", help=f"COMPRA FORTE 🟢: Desconto patrimonial Graham > {margem_graham_exigida}% E Margem Bazin > {margem_bazin_exigida}%.\nMANTER 🟡: Margem equilibrada.\nAVALIAR VENDA 🔴: Ativo caro e dividendos abaixo da taxa exigida.")
            })

            st.divider()
            st.markdown("### ❄️ Projeção Bola de Neve Completa (Juros Compostos Reais)")
            
            saldo_total_atual = df_perf_final['Saldo Atual'].sum() + patrimonio_fora
            div_total_12m = st.session_state.df_simul['Div. Projetado (R$)'].sum()
            base_div_mensal = div_total_12m / 12 if div_total_12m > 0 else 0
            
            meses_lista, patr_base_lista, aportes_lista, compostos_lista = [], [], [], []
            saldo_corrente, acum_aportes, acum_juros_divs = saldo_total_atual, 0.0, 0.0
            
            for m in range(13):
                meses_lista.append(f"Mês {m}")
                patr_base_lista.append(saldo_total_atual)
                aportes_lista.append(acum_aportes)
                compostos_lista.append(acum_juros_divs)
                
                ganho_capital = saldo_corrente * rentabilidade_ganho
                fator_cresc_mensal = (1 + cresc_dividendos_anual) ** (m / 12)
                dividendo_mes = base_div_mensal * fator_cresc_mensal
                
                acum_juros_divs += (ganho_capital + dividendo_mes)
                acum_aportes += aporte_mensal_planejado
                saldo_corrente += (ganho_capital + dividendo_mes + aporte_mensal_planejado)
                
            df_proj_real = pd.DataFrame({"Mês": meses_lista, "Patrimônio Base": patr_base_lista, "Aportes Acumulados": aportes_lista, "Juros + Divs. Reinvestidos": compostos_lista})
            df_proj_real_melt = df_proj_real.melt(id_vars=["Mês"], value_vars=["Patrimônio Base", "Aportes Acumulados", "Juros + Divs. Reinvestidos"], var_name="Componente", value_name="Valor")
            
            fig_proj = px.area(df_proj_real_melt, x="Mês", y="Valor", color="Componente", title="Evolução Patrimonial Composta (Ganho de Capital + Dividendos)", color_discrete_sequence=["#34495e", "#2980b9", "#27ae60"])
            fig_proj.update_traces(hovertemplate='%{x}<br>%{data.name}: R$ %{y:,.2f}<extra></extra>')
            fig_proj.update_layout(xaxis_title="Linha do Tempo", yaxis_title="Montante Projetado (R$)", margin=dict(t=40))
            fig_proj.update_yaxes(range=[saldo_total_atual * 0.98, saldo_corrente * 1.02])
            st.plotly_chart(fig_proj, use_container_width=True)

        with tab6:
            st.markdown("### 1. Performance Global (Individualizada por Ativo)")
            todos_ativos = df_perf_final['Ativo'].tolist()
            c_sel, c_ind = st.columns([2, 1])
            with c_sel: ativos_selecionados = st.multiselect("Selecione os ativos:", todos_ativos, default=todos_ativos[:6], key="ms_g")
            with c_ind: ind_selecionados = st.multiselect("Indicadores:", ["Evolução c/ Div", "CDI Acum.", "IPCA Acum."], default=["Evolução c/ Div", "CDI Acum.", "IPCA Acum."], key="ind_g")
            
            if ativos_selecionados and ind_selecionados:
                df_grafico = df_perf_final[df_perf_final['Ativo'].isin(ativos_selecionados)].copy()
                df_grafico['Período'] = df_grafico['Data Média'].astype(str) + " até Hoje"
                df_melt = df_grafico.melt(id_vars=["Ativo", "Período"], value_vars=ind_selecionados, var_name="Indicador", value_name="Rentabilidade")
                
                titulo_graf1 = f"Performance Desde: {df_grafico.iloc[0]['Data Média']} até Hoje" if len(ativos_selecionados) == 1 else "Performance Baseada nas Datas Médias de Cada Ativo"
                fig1 = px.bar(df_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group", hover_data=["Período"], title=titulo_graf1)
                fig1.update_traces(hovertemplate='<b>%{x}</b> (%{data.name})<br>Período: %{customdata[0]}<br>Rentabilidade: %{y:.2f}%<extra></extra>')
                fig1.update_layout(yaxis_ticksuffix=" %", margin=dict(t=40))
                st.plotly_chart(fig1, use_container_width=True)
                st.dataframe(df_grafico[['Ativo', 'Período', 'Evolução c/ Div', 'CDI Acum.', 'IPCA Acum.']], use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("### 2. Análise de Período Específico (Janela Tática)")
            c_dt1, c_dt2, c_btn = st.columns([1, 1, 1])
            with c_dt1: dt_inicio_custom = st.date_input("Data de Início", pd.Timestamp.now().date() - pd.DateOffset(years=1))
            with c_dt2: dt_fim_custom = st.date_input("Data de Fim", pd.Timestamp.now().date())
            with c_btn: ind_custom = st.multiselect("Indicadores:", ["Retorno Total (%)", "CDI Período", "IPCA Período"], default=["Retorno Total (%)", "CDI Período", "IPCA Período"], key="ind_custom")
            
            if st.button("Gerar Análise do Período", use_container_width=True):
                if not ativos_selecionados: st.warning("Selecione os ativos na caixa acima.")
                else:
                    with st.spinner("Conectando à bolsa para o período específico..."):
                        linhas_custom = []
                        df_macro = carregar_macro()
                        dt_fim_yf = dt_fim_custom + pd.Timedelta(days=1)
                        for t in ativos_selecionados:
                            try:
                                acao = yf.Ticker(f"{t}.SA")
                                hist = acao.history(start=dt_inicio_custom.strftime('%Y-%m-%d'), end=dt_fim_yf.strftime('%Y-%m-%d'))
                                if not hist.empty:
                                    preco_ini, preco_fim = float(hist['Close'].iloc[0]), float(hist['Close'].iloc[-1])
                                    divs = acao.dividends
                                    divs_periodo = float(divs[(divs.index.tz_localize(None) >= pd.to_datetime(dt_inicio_custom)) & (divs.index.tz_localize(None) <= pd.to_datetime(dt_fim_custom))].sum())
                                    evolucao_custom = (((preco_fim + divs_periodo) / preco_ini) - 1) * 100
                                    cdi_custom, ipca_custom = calcular_macro_acumulado(df_macro, pd.to_datetime(dt_inicio_custom), pd.to_datetime(dt_fim_custom))
                                    linhas_custom.append({"Ativo": t, "Retorno Total (%)": evolucao_custom, "CDI Período": cdi_custom, "IPCA Período": ipca_custom})
                            except: pass
                        if linhas_custom:
                            df_custom = pd.DataFrame(linhas_custom)
                            df_custom['Período'] = f"{dt_inicio_custom.strftime('%d/%m/%Y')} a {dt_fim_custom.strftime('%d/%m/%Y')}"
                            df_custom_melt = df_custom.melt(id_vars=["Ativo", "Período"], value_vars=ind_custom, var_name="Indicador", value_name="Rentabilidade")
                            fig_custom = px.bar(df_custom_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group", hover_data=["Período"], title=f"Performance no Período: {dt_inicio_custom.strftime('%d/%m/%Y')} a {dt_fim_custom.strftime('%d/%m/%Y')}")
                            fig_custom.update_traces(hovertemplate='<b>%{x}</b> (%{data.name})<br>Período: %{customdata[0]}<br>Rentabilidade: %{y:.2f}%<extra></extra>')
                            fig_custom.update_layout(yaxis_ticksuffix=" %", margin=dict(t=40))
                            st.plotly_chart(fig_custom, use_container_width=True)
                            st.dataframe(df_custom, use_container_width=True, hide_index=True)

        # ==========================================
        # ABA 7: TERMINAL DE IA PARAMETRIZADO NATIVO COM AUTO-DISCOVERY + FALLBACK 429
        # ==========================================
        with tab7:
            st.markdown("### 🏢 Comitê de Alocação IA - Visão CNPI Sênior")
            st.markdown("Insira sua chave de API corporativa para habilitar o processamento analítico profundo. O motor fará o rastreamento automático do modelo ideal para a sua credencial, com contingência ativada.")
            
            api_key = st.text_input("Chave API do Google Gemini (Opcional):", type="password")
            
            st.write("---")
            for msg in st.session_state.historico_chat:
                with st.chat_message(msg["role"]): st.write(msg["content"])
                
            if prompt := st.chat_input("Ex: Baseado no Focus atual, quais FIIs eu devo reavaliar?"):
                st.session_state.historico_chat.append({"role": "user", "content": prompt})
                with st.chat_message("user"): st.write(prompt)
                
                contexto_carteira = df_perf_final[['Ativo', 'Qtd', 'Preço Médio', 'Preço Atual', 'Total Investido', 'Saldo Atual', 'Resultado (R$)', 'Total Div. (R$)', 'Evolução c/ Div']].to_csv(index=False, sep='|')
                contexto_macro = f"Selic Corrente/Projetada: {proj_focus.get(f'Selic_{ano_atual}', 14.0)}% | IPCA Esperado: {proj_focus.get(f'IPCA_{ano_atual}', 5.33)}%"
                
                sys_prompt = f"""
                Você é uma Analista de Investimentos Sênior (CNPI) e Gestora de Portfólio.
                Sua linguagem é corporativa, executiva, fria, baseada em dados e voltada a relatórios de alta governança.
                Você tem acesso irrestrito ao banco de dados da carteira do usuário e ao cenário macroeconômico do Banco Central.
                
                [MATRIZ DE DADOS EM TEMPO REAL]
                {contexto_carteira}
                
                [CENÁRIO MACRO FOCUS]
                {contexto_macro}
                
                [DIRETRIZES DE RESPOSTA]
                - Não use linguagem de iniciante. Use termos técnicos apropriados.
                - Cruze os dados para fornecer números exatos na sua resposta.
                """
                
                if api_key:
                    try:
                        url_models = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
                        res_models = requests.get(url_models, timeout=10)
                        
                        if res_models.status_code == 200:
                            modelos_disponiveis = res_models.json().get('models', [])
                            modelos_geracao = [m['name'] for m in modelos_disponiveis if 'generateContent' in m.get('supportedGenerationMethods', [])]
                            
                            # IGNORAMOS OS MODELOS '3.1' OU EXPERIMENTAIS PAGOS - FOCO TOTAL NOS FREE-TIER MAIS ROBUSTOS
                            modelos_prioridade = ['gemini-1.5-flash', 'gemini-1.0-pro', 'gemini-1.5-pro']
                            modelos_candidatos = []
                            for pref in modelos_prioridade:
                                for m_disp in modelos_geracao:
                                    if pref in m_disp and m_disp not in modelos_candidatos:
                                        modelos_candidatos.append(m_disp)
                                        
                            if not modelos_candidatos:
                                modelos_candidatos = [m for m in modelos_geracao if 'vision' not in m.lower()]
                            
                            sucesso_api = False
                            for modelo_alvo in modelos_candidatos:
                                url_api = f"https://generativelanguage.googleapis.com/v1beta/{modelo_alvo}:generateContent?key={api_key}"
                                payload = {"contents": [{"parts": [{"text": f"{sys_prompt}\n\nPergunta do Gestor: {prompt}"}]}]}
                                headers = {"Content-Type": "application/json"}
                                
                                res_api = requests.post(url_api, json=payload, headers=headers, timeout=20)
                                
                                if res_api.status_code == 200:
                                    resposta = res_api.json()['contents'][0]['parts'][0]['text']
                                    sucesso_api = True
                                    break
                                elif res_api.status_code == 429:
                                    # Erro 429 (Cota Excedida / Zero Limit): Pula silenciosamente para o próximo modelo da lista
                                    continue
                                else:
                                    resposta = f"⚠️ Erro ao gerar conteúdo (Status {res_api.status_code}): {res_api.text}"
                                    sucesso_api = True # Para quebrar o loop em caso de erro de sintaxe/chave
                                    break
                                    
                            if not sucesso_api and res_api.status_code == 429:
                                resposta = "⚠️ Análise de Infraestrutura: Todos os modelos disponíveis para a sua chave atingiram o limite de cota gratuita (Error 429 - Resource Exhausted). O Google exige que aguarde alguns minutos para fazer novas requisições ou habilite o plano de faturação no AI Studio. A IA local quantitativa assumirá as requisições até o desbloqueio."
                                
                        else:
                            resposta = f"⚠️ Chave inválida ou bloqueada pelo Google. Erro de Autenticação: {res_models.text}"
                    except Exception as e:
                        resposta = f"⚠️ Falha na infraestrutura de rede. Detalhe: {e}"
                else:
                    p_u = prompt.upper()
                    if "CONCENTRAÇÃO" in p_u or "PESO" in p_u or "RISCO" in p_u or "SETOR" in p_u:
                        maior_ativo = df_perf_final.sort_values('Saldo Atual', ascending=False).iloc[0]
                        maior_setor = df_perf_final.groupby('Setor')['Saldo Atual'].sum().idxmax()
                        peso_setor = (df_perf_final.groupby('Setor')['Saldo Atual'].sum().max() / df_perf_final['Saldo Atual'].sum()) * 100
                        resposta = f"**[Diagnóstico Quantitativo Sênior]** O seu risco de cauda setorial concentra-se em **{maior_setor}** ({peso_setor:.2f}% do capital). Sob uma Selic projetada de {proj_focus.get(f'Selic_{ano_atual}', 14.0)}%, este nível de exposição exige monitorização estrita do custo de oportunidade."
                    elif "BAZIN" in p_u or "TETO" in p_u or "COMPRA" in p_u:
                        ativo_bazin = st.session_state.df_rec_bazin.sort_values('Margem Segurança', ascending=False).iloc[0]
                        resposta = f"**[Auditoria de Valuation - Renda]** De acordo com a base histórica de dividendos, o ativo com maior assimetria positiva de preço (desconto tático) é **{ativo_bazin['Ativo']}**, apresentando Margem de Segurança de **{ativo_bazin['Margem Segurança']:.2f}%**."
                    else:
                        resposta = f"**[Modo Local]** Chave de API ausente. Sem o token, estou limitada a processar relatórios quantitativos pré-programados sobre 'concentração setorial' ou 'assimetria Bazin'. Por favor, insira a chave para liberar a inferência profunda."
                
                st.session_state.historico_chat.append({"role": "assistant", "content": resposta})
                st.rerun()
