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

# 📡 NOVO MOTOR BACEN: Mais robusto para capturar os dados corretos
@st.cache_data(ttl=86400)
def obter_projecoes_focus():
    ano_atual = pd.Timestamp.now().year
    try:
        # Acesso via API OData Oficial filtrando para garantir que pegamos os dados mais recentes (DataReferencia)
        url = f"https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?$top=100&$filter=Indicador%20eq%20'IPCA'%20or%20Indicador%20eq%20'Selic'&$orderby=Data%20desc&$format=json"
        res = requests.get(url, timeout=5).json()
        
        if 'value' in res and len(res['value']) > 0:
            df = pd.DataFrame(res['value'])
            
            # Puxamos as projeções para Dezembro do ano atual e Dezembro do próximo ano
            df_atual = df[(df['DataReferencia'].str.startswith(str(ano_atual))) & (df['DataReferencia'].str.endswith('12'))]
            df_prox = df[(df['DataReferencia'].str.startswith(str(ano_atual+1))) & (df['DataReferencia'].str.endswith('12'))]
            
            ipca_atual = df_atual[df_atual['Indicador'] == 'IPCA']['Mediana'].values[0] if not df_atual[df_atual['Indicador'] == 'IPCA'].empty else 3.80
            selic_atual = df_atual[df_atual['Indicador'] == 'Selic']['Mediana'].values[0] if not df_atual[df_atual['Indicador'] == 'Selic'].empty else 10.50
            
            ipca_prox = df_prox[df_prox['Indicador'] == 'IPCA']['Mediana'].values[0] if not df_prox[df_prox['Indicador'] == 'IPCA'].empty else 3.90
            selic_prox = df_prox[df_prox['Indicador'] == 'Selic']['Mediana'].values[0] if not df_prox[df_prox['Indicador'] == 'Selic'].empty else 9.50
            
            return {f"IPCA_{ano_atual}": float(ipca_atual), f"Selic_{ano_atual}": float(selic_atual), f"IPCA_{ano_atual+1}": float(ipca_prox), f"Selic_{ano_atual+1}": float(selic_prox)}, ano_atual
        else:
            raise ValueError("JSON vazio")
    except: 
        # Fallback Seguro com os dados do Focus atualizados para meados de 2024 (Mercado Real)
        return {f"IPCA_{ano_atual}": 3.90, f"Selic_{ano_atual}": 10.50, f"IPCA_{ano_atual+1}": 3.78, f"Selic_{ano_atual+1}": 9.50}, ano_atual

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
st.sidebar.markdown("Para iniciar, forneça o seu banco de dados ou a planilha da B3.")

arquivo_principal = st.sidebar.file_uploader("1️⃣ Planilha Principal (B3 ou Backup)", type=["xlsx", "csv"])

st.sidebar.markdown("---")
st.sidebar.markdown("**Atualização Incremental (Opcional)**")
arquivo_novo = st.sidebar.file_uploader("2️⃣ Novas Negociações B3", type=["xlsx", "csv"])

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
                    linhas_base.append({
                        "Código de Negociação": row['Ativo'],
                        "Tipo de Movimentação": "Compra",
                        "Data do Negócio": pd.to_datetime(row['Data Média']),
                        "Quantidade": row['Quantidade'],
                        "Valor": row['Quantidade'] * row['Preço Médio']
                    })
                df_historico_falso = pd.DataFrame(linhas_base)
                df_mesclado = pd.concat([df_historico_falso, df_novos], ignore_index=True)
                base_atual = processar_planilha_b3(df_mesclado)
                
            st.session_state.df_base = base_atual
            st.rerun()
    else:
        st.sidebar.error("⚠️ Insira o arquivo principal para iniciar.")

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

            if ticker in fundamentos_br:
                vpa, lpa = fundamentos_br[ticker]['vpa'], fundamentos_br[ticker]['lpa']

            cdi, ipca = calcular_macro_acumulado(df_macro, data_compra)
            
            dados_mercado[ticker] = {
                "Qtd": float(row['Quantidade']), "PM": float(row['Preço Médio']), "Data": data_compra,
                "Preço Atual": preco_atual, "Div_Total": divs_total, "CDI": cdi, "IPCA": ipca, "Setor": setor, "Tipo": tipo_ativo
            }
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
                "Ativo": t, "Tipo": dm["Tipo"], "Setor": dm["Setor"], "Qtd": int(dm['Qtd']), 
                "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
                "Total Investido": investido, "Saldo Atual": saldo, "Resultado (R$)": resultado,
                "Data Média": dm['Data'].strftime('%d/%m/%Y'), "Meses (Média)": int(calcular_meses(dm['Data'])),
                "Total Div. (R$)": dm['Div_Total'], "DY on Cost": yoc,
                "Evolução c/ Div": var_c_div, "IPCA Acum.": dm['IPCA'], "CDI Acum.": dm['CDI']
            })
        df_perf_final = pd.DataFrame(linhas_perf)

        st.write("---")
        # ==========================================
        # DASHBOARD EXECUTIVO (MÉTRICAS)
        # ==========================================
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

        st.download_button(label="📥 Exportar Relatório (Excel)", data=gerar_excel_premium(df_perf_final, st.session_state.df_simul), file_name="Relatorio_Carteira.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Visão Geral", "💰 Bazin (Renda)", "🏢 Graham (Valor)", 
            "⚖️ Pesos e Setores", "🎯 Recomendações e Projeções", "📈 Gráficos Interativos"
        ])
        
        with tab1:
            st.dataframe(df_perf_final.drop(columns=['Tipo', 'Setor']), use_container_width=True, hide_index=True, column_config={
                "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f"), "Preço Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Total Investido": st.column_config.NumberColumn(format="R$ %.2f"), "Saldo Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Resultado (R$)": st.column_config.NumberColumn(format="R$ %.2f"), "Total Div. (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                "DY on Cost": st.column_config.NumberColumn(format="%.2f %%"), "Evolução c/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "IPCA Acum.": st.column_config.NumberColumn(format="%.2f %%"), "CDI Acum.": st.column_config.NumberColumn(format="%.2f %%")
            })

        with tab2:
            yield_desejado = st.number_input("Taxa de Risco Exigida (%):", value=6.0, step=0.5) / 100.0
            df_bazin_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)"]].copy()
            df_bazin_editado = st.data_editor(df_bazin_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"], key="edit_bazin", column_config={"Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"), "Div. Projetado (R$)": st.column_config.NumberColumn(format="R$ %.2f")})
            st.session_state.df_simul["Div. Projetado (R$)"] = df_bazin_editado["Div. Projetado (R$)"]
            
            linhas_bazin = []
            for _, row in df_bazin_editado.iterrows():
                bazin = (row['Div. Projetado (R$)'] / yield_desejado) if (row['Div. Projetado (R$)'] > 0 and yield_desejado > 0) else 0.0
                margem_b = (((bazin / row['Cotação Atual']) - 1) * 100) if (bazin > 0 and row['Cotação Atual'] > 0) else 0.0
                linhas_bazin.append({"Ativo": row['Ativo'], "Preço Teto (Bazin)": bazin, "Margem Segurança": margem_b})
            
            st.session_state.df_rec_bazin = pd.DataFrame(linhas_bazin)
            st.dataframe(st.session_state.df_rec_bazin, use_container_width=True, hide_index=True, column_config={"Preço Teto (Bazin)": st.column_config.NumberColumn(format="R$ %.2f"), "Margem Segurança": st.column_config.NumberColumn(format="%.2f %%")})

        with tab3:
            df_graham_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "VPA (Contábil)", "LPA Projetado"]].copy()
            df_graham_editado = st.data_editor(df_graham_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"], key="edit_graham", column_config={"Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"), "VPA (Contábil)": st.column_config.NumberColumn(format="R$ %.2f"), "LPA Projetado": st.column_config.NumberColumn(format="R$ %.2f")})
            st.session_state.df_simul["VPA (Contábil)"] = df_graham_editado["VPA (Contábil)"]
            st.session_state.df_simul["LPA Projetado"] = df_graham_editado["LPA Projetado"]
            
            linhas_graham = []
            for _, row in df_graham_editado.iterrows():
                graham = (22.5 * row['LPA Projetado'] * row['VPA (Contábil)']) ** 0.5 if (row['LPA Projetado'] > 0 and row['VPA (Contábil)'] > 0) else 0.0
                margem_g = (((graham / row['Cotação Atual']) - 1) * 100) if (graham > 0 and row['Cotação Atual'] > 0) else 0.0
                linhas_graham.append({"Ativo": row['Ativo'], "Preço Justo (Graham)": graham, "Margem Segurança": margem_g})
                
            st.session_state.df_rec_graham = pd.DataFrame(linhas_graham)
            st.dataframe(st.session_state.df_rec_graham, use_container_width=True, hide_index=True, column_config={"Preço Justo (Graham)": st.column_config.NumberColumn(format="R$ %.2f"), "Margem Segurança": st.column_config.NumberColumn(format="%.2f %%")})

        with tab4:
            st.markdown("### ⚖️ Distribuição do Portfólio")
            c_g1, c_g2, c_g3 = st.columns(3)
            
            fig_tipo = px.pie(df_perf_final, values='Saldo Atual', names='Tipo', hole=0.4)
            fig_tipo.update_layout(title_text="Por Tipo", title_x=0.2)
            c_g1.plotly_chart(fig_tipo, use_container_width=True)
            
            fig_ativo = px.pie(df_perf_final, values='Saldo Atual', names='Ativo')
            fig_ativo.update_layout(title_text="Por Ativo", title_x=0.2)
            c_g2.plotly_chart(fig_ativo, use_container_width=True)
            
            fig_setor = px.pie(df_perf_final, values='Saldo Atual', names='Setor')
            fig_setor.update_layout(title_text="Por Setores/Classe", title_x=0.2)
            c_g3.plotly_chart(fig_setor, use_container_width=True)

            st.divider()
            st.markdown("### 📋 Tabelas de Alocação Exata")
            def gerar_tabela_peso(coluna):
                df_peso = df_perf_final.groupby(coluna)['Saldo Atual'].sum().reset_index()
                df_peso['Peso (%)'] = (df_peso['Saldo Atual'] / df_peso['Saldo Atual'].sum()) * 100
                return df_peso.sort_values('Peso (%)', ascending=False)

            c_t1, c_t2, c_t3 = st.columns(3)
            with c_t1: st.dataframe(gerar_tabela_peso('Tipo'), use_container_width=True, hide_index=True, column_config={"Saldo Atual": st.column_config.NumberColumn(format="R$ %.2f"), "Peso (%)": st.column_config.NumberColumn(format="%.2f %%")})
            with c_t2: st.dataframe(gerar_tabela_peso('Ativo'), use_container_width=True, hide_index=True, column_config={"Saldo Atual": st.column_config.NumberColumn(format="R$ %.2f"), "Peso (%)": st.column_config.NumberColumn(format="%.2f %%")})
            with c_t3: st.dataframe(gerar_tabela_peso('Setor'), use_container_width=True, hide_index=True, column_config={"Saldo Atual": st.column_config.NumberColumn(format="R$ %.2f"), "Peso (%)": st.column_config.NumberColumn(format="%.2f %%")})

        with tab5:
            # Integração Boletim Focus BCB
            proj_focus, ano_atual = obter_projecoes_focus()
            st.markdown(f"### 🇧🇷 Expectativas Macroeconômicas (Boletim Focus Bacen)")
            st.info(f"**IPCA {ano_atual}:** {proj_focus.get(f'IPCA_{ano_atual}', '--')}%  |  **Selic {ano_atual}:** {proj_focus.get(f'Selic_{ano_atual}', '--')}%  ||  **IPCA {ano_atual+1}:** {proj_focus.get(f'IPCA_{ano_atual+1}', '--')}%  |  **Selic {ano_atual+1}:** {proj_focus.get(f'Selic_{ano_atual+1}', '--')}%")
            
            st.markdown("### 🤖 Radar de Oportunidades do Especialista")
            df_recs = pd.merge(df_perf_final[['Ativo', 'Tipo']], st.session_state.df_rec_bazin[['Ativo', 'Margem Segurança']], on='Ativo', how='left').rename(columns={'Margem Segurança': 'Margem Bazin (%)'})
            df_recs = pd.merge(df_recs, st.session_state.df_rec_graham[['Ativo', 'Margem Segurança']], on='Ativo', how='left').rename(columns={'Margem Segurança': 'Margem Graham (%)'})
            
            recomendacoes = []
            for _, r in df_recs.iterrows():
                if r['Tipo'] == 'Ação':
                    if r['Margem Graham (%)'] > 15 and r['Margem Bazin (%)'] > 5: recomendacoes.append("COMPRA FORTE 🟢")
                    elif r['Margem Graham (%)'] > 0 or r['Margem Bazin (%)'] > 0: recomendacoes.append("MANTER / COMPRA 🟡")
                    else: recomendacoes.append("AVALIAR VENDA 🔴")
                else: 
                    if r['Margem Bazin (%)'] > 5: recomendacoes.append("COMPRA FORTE 🟢")
                    elif r['Margem Bazin (%)'] > -5: recomendacoes.append("MANTER 🟡")
                    else: recomendacoes.append("AVALIAR VENDA 🔴")
                    
            df_recs['Status Recomendações'] = recomendacoes
            
            st.dataframe(df_recs, use_container_width=True, hide_index=True, column_config={
                "Margem Bazin (%)": st.column_config.NumberColumn(format="%.2f %%"), 
                "Margem Graham (%)": st.column_config.NumberColumn(format="%.2f %%"),
                "Status Recomendações": st.column_config.TextColumn(
                    "Status Recomendações ❓",
                    help="COMPRA FORTE 🟢: Desconto patrimonial e boa renda. | MANTER / COMPRA 🟡: Pelo menos uma das margens positiva. | AVALIAR VENDA 🔴: Ativo caro e dividendos abaixo do Yield Exigido."
                )
            })

            st.divider()
            
            st.markdown("### ❄️ Projeção Bola de Neve Completa (Juros Compostos Reais)")
            st.markdown("Ajuste os parâmetros abaixo para simular o crescimento do seu patrimônio com reinvestimento de dividendos e novos aportes.")
            
            # PARÂMETROS FINANCEIROS DE ALTO NÍVEL
            c_p1, c_p2, c_p3, c_p4 = st.columns(4)
            patrimonio_fora = c_p1.number_input("Patrimônio Fora da Bolsa (R$):", value=0.0, step=10000.0)
            aporte_mensal_planejado = c_p2.number_input("Aporte Mensal (R$):", value=2000.0, step=500.0)
            rentabilidade_ganho = c_p3.number_input("Rentab. Mensal Estimada (%):", value=0.8, step=0.1) / 100.0
            cresc_dividendos_anual = c_p4.number_input("Crescimento Anual de Divs (%):", value=5.0, step=1.0) / 100.0
            
            saldo_total_atual = df_perf_final['Saldo Atual'].sum() + patrimonio_fora
            div_total_12m = st.session_state.df_simul['Div. Projetado (R$)'].sum()
            base_div_mensal = div_total_12m / 12 if div_total_12m > 0 else 0
            
            meses_lista, patr_base_lista, aportes_lista, compostos_lista = [], [], [], []
            saldo_corrente = saldo_total_atual
            acum_aportes = 0.0
            acum_juros_divs = 0.0
            
            # Motor de Juros Compostos Puro Mês a Mês (CORREÇÃO DE TIPOGRAFIA: 'Patrimônio Base')
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
                
            df_proj_real = pd.DataFrame({
                "Mês": meses_lista, 
                "Patrimônio Base": patr_base_lista, # Tipografia corrigida
                "Aportes Acumulados": aportes_lista, 
                "Juros + Divs. Reinvestidos": compostos_lista
            })
            
            # Melt sem o erro de acentuação
            df_proj_real_melt = df_proj_real.melt(
                id_vars=["Mês"], 
                value_vars=["Patrimônio Base", "Aportes Acumulados", "Juros + Divs. Reinvestidos"], 
                var_name="Componente", value_name="Valor"
            )
            
            fig_proj = px.area(df_proj_real_melt, x="Mês", y="Valor", color="Componente", title="Evolução Patrimonial Composta (Ganho de Capital + Dividendos)", color_discrete_sequence=["#34495e", "#2980b9", "#27ae60"])
            fig_proj.update_traces(hovertemplate='%{x}<br>%{data.name}: R$ %{y:,.2f}<extra></extra>')
            fig_proj.update_layout(xaxis_title="Linha do Tempo", yaxis_title="Montante Projetado (R$)", margin=dict(t=40))
            
            min_y = saldo_total_atual * 0.98
            max_y = saldo_corrente * 1.02
            fig_proj.update_yaxes(range=[min_y, max_y])
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
                
                st.markdown("#### 📋 Dados Tabulares Globais")
                st.dataframe(df_grafico[['Ativo', 'Período', 'Evolução c/ Div', 'CDI Acum.', 'IPCA Acum.']], use_container_width=True, hide_index=True, column_config={"Evolução c/ Div": st.column_config.NumberColumn(format="%.2f %%"), "CDI Acum.": st.column_config.NumberColumn(format="%.2f %%"), "IPCA Acum.": st.column_config.NumberColumn(format="%.2f %%")})
            elif not ind_selecionados:
                st.info("Selecione pelo menos um indicador para exibir.")

            st.divider()
            
            st.markdown("### 2. Análise de Período Específico (Janela Tática)")
            c_dt1, c_dt2, c_btn = st.columns([1, 1, 1])
            with c_dt1: dt_inicio_custom = st.date_input("Data de Início", pd.Timestamp.now().date() - pd.DateOffset(years=1))
            with c_dt2: dt_fim_custom = st.date_input("Data de Fim", pd.Timestamp.now().date())
            with c_btn:
                ind_custom = st.multiselect("Indicadores:", ["Retorno Total (%)", "CDI Período", "IPCA Período"], default=["Retorno Total (%)", "CDI Período", "IPCA Período"], key="ind_custom")
            
            if st.button("Gerar Análise do Período", use_container_width=True):
                if not ativos_selecionados: st.warning("Selecione os ativos na caixa acima.")
                elif not ind_custom: st.warning("Selecione pelo menos um indicador.")
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
                            periodo_str = f"{dt_inicio_custom.strftime('%d/%m/%Y')} a {dt_fim_custom.strftime('%d/%m/%Y')}"
                            df_custom['Período'] = periodo_str
                            
                            df_custom_melt = df_custom.melt(id_vars=["Ativo", "Período"], value_vars=ind_custom, var_name="Indicador", value_name="Rentabilidade")
                            
                            titulo_graf2 = f"Performance no Período: {dt_inicio_custom.strftime('%d/%m/%Y')} a {dt_fim_custom.strftime('%d/%m/%Y')}"
                            fig_custom = px.bar(df_custom_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group", hover_data=["Período"], title=titulo_graf2)
                            
                            fig_custom.update_traces(hovertemplate='<b>%{x}</b> (%{data.name})<br>Período: %{customdata[0]}<br>Rentabilidade: %{y:.2f}%<extra></extra>')
                            fig_custom.update_layout(yaxis_ticksuffix=" %", margin=dict(t=40))
                            st.plotly_chart(fig_custom, use_container_width=True)
                            
                            st.markdown("#### 📋 Dados Tabulares do Período Específico")
                            st.dataframe(df_custom, use_container_width=True, hide_index=True, column_config={"Retorno Total (%)": st.column_config.NumberColumn(format="%.2f %%"), "CDI Período": st.column_config.NumberColumn(format="%.2f %%"), "IPCA Período": st.column_config.NumberColumn(format="%.2f %%")})
                        else: st.error("Sem histórico para o período.")
