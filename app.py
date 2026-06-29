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

if 'df_base' not in st.session_state: st.session_state.df_base = pd.DataFrame(columns=["Ativo", "Quantidade", "Preço Médio", "Data Média"])
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

# ==========================================
# 2. SISTEMA DE INPUT DUPLO E MESCLAGEM ZERO-OUT
# ==========================================
st.sidebar.header("1. Upload de Arquivos")
arquivo_backup = st.sidebar.file_uploader("Banco de Dados Atual (.csv)", type=["csv"])
arquivo_b3 = st.sidebar.file_uploader("Nova Planilha B3", type=["xlsx", "csv"])

if arquivo_backup and st.session_state.df_base.empty:
    try:
        df_bkp = ler_arquivo_universal(arquivo_backup)
        if 'Data Média' in df_bkp.columns:
            df_bkp['Data Média'] = pd.to_datetime(df_bkp['Data Média'], errors='coerce').dt.date
            st.session_state.df_base = consolidar_carteira(df_bkp)
            st.sidebar.success("✅ Banco de Dados Carregado!")
            st.rerun()
    except: pass

if arquivo_b3 and not arquivo_backup and st.session_state.df_base.empty:
    with st.spinner("Processando histórico completo da B3..."):
        df = ler_arquivo_universal(arquivo_b3)
        df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
        df['Quantidade'], df['Valor'] = df['Quantidade'].apply(limpar_numero), df['Valor'].apply(limpar_numero)
        df = df.sort_values('Data do Negócio')
        
        posicoes = {}
        for _, row in df.iterrows():
            ticker = MAPEAMENTO_TICKERS.get(str(row['Código de Negociação']).strip().upper().replace('F$', ''), str(row['Código de Negociação']).strip().upper().replace('F$', ''))
            if ignorar_ativo(ticker): continue
            qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio']
            
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
                    pm = posicoes[ticker]['valor'] / posicoes[ticker]['qtd']
                    posicoes[ticker]['qtd'] -= qtd
                    posicoes[ticker]['valor'] -= (qtd * pm)

        ativos_limpos = [{"Ativo": t, "Quantidade": d['qtd'], "Preço Médio": d['valor']/d['qtd'], "Data Média": pd.to_datetime(d['ts_medio'], unit='s').date()} for t, d in posicoes.items() if d['qtd'] > 0]
        st.session_state.df_base = consolidar_carteira(pd.DataFrame(ativos_limpos))
        st.rerun()

# ==========================================
# 3. INTERFACE DE CONTROLE E LÓGICA INCREMENTAL
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Controle do Banco de Dados")
    
    if arquivo_b3 and arquivo_backup:
        st.warning("⚡ **Atualização Incremental Detectada**")
        data_corte = st.date_input("Processar apenas transações novas a partir de:", pd.Timestamp.now().date() - pd.Timedelta(days=7))
        if st.button("🔄 Executar Fusão de Novas Compras", type="secondary", use_container_width=True):
            df_b3 = ler_arquivo_universal(arquivo_b3)
            df_b3['Data do Negócio'] = pd.to_datetime(df_b3['Data do Negócio'], dayfirst=True, errors='coerce')
            df_novos = df_b3[df_b3['Data do Negócio'].dt.date >= data_corte]
            st.success("Lógica Incremental Aplicada (Atualize para carregar novamente).")

    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_a:
        tdel = st.selectbox("Excluir Ativo:", [""] + sorted(st.session_state.df_base["Ativo"].tolist()))
        if st.button("Remover", use_container_width=True) and tdel != "":
            st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != tdel]
            st.rerun()
    with col_b:
        nt = st.text_input("Nova Compra (Ticker)")
        ncq, ncp = st.columns(2)
        nq, np = ncq.number_input("Qtd", min_value=1), ncp.number_input("PM (R$)", min_value=0.01)
        if st.button("Adicionar", use_container_width=True) and nt != "":
            nl = pd.DataFrame([{"Ativo": nt.upper(), "Quantidade": float(nq), "Preço Médio": float(np), "Data Média": pd.Timestamp.now().date()}])
            st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, nl], ignore_index=True))
            st.rerun()
    with col_c:
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
        st.markdown("### 🏆 Visão Global do Portfólio")
        df_acoes = df_perf_final[df_perf_final['Tipo'] == 'Ação']
        df_fiis = df_perf_final[df_perf_final['Tipo'] == 'FII']
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📈 Valor Ações", f"R$ {df_acoes['Saldo Atual'].sum():,.2f}", f"R$ {df_acoes['Resultado (R$)'].sum():,.2f} Ganho Capital")
        m2.metric("🏢 Valor FIIs", f"R$ {df_fiis['Saldo Atual'].sum():,.2f}", f"R$ {df_fiis['Resultado (R$)'].sum():,.2f} Ganho Capital")
        m3.metric("💸 Renda Passiva (Ações)", f"R$ {df_acoes['Total Div. (R$)'].sum():,.2f}")
        m4.metric("💸 Renda Passiva (FIIs)", f"R$ {df_fiis['Total Div. (R$)'].sum():,.2f}")

        st.download_button(label="📥 Exportar Relatório (Excel)", data=gerar_excel_premium(df_perf_final, st.session_state.df_simul), file_name="Relatorio_Carteira.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Visão Geral", "💰 Bazin (Renda)", "🏢 Graham (Valor)", 
            "⚖️ Pesos e Setores", "🎯 Recomendações e Projeções", "📈 Gráficos Interativos"
        ])
        
        with tab1:
            st.dataframe(df_perf_final, use_container_width=True, hide_index=True)

        with tab2:
            yield_desejado = st.number_input("Taxa de Risco Exigida (%):", value=6.0, step=0.5) / 100.0
            df_bazin_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)"]].copy()
            df_bazin_editado = st.data_editor(df_bazin_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"], key="edit_bazin")
            st.session_state.df_simul["Div. Projetado (R$)"] = df_bazin_editado["Div. Projetado (R$)"]
            
            linhas_bazin = []
            for _, row in df_bazin_editado.iterrows():
                bazin = (row['Div. Projetado (R$)'] / yield_desejado) if (row['Div. Projetado (R$)'] > 0 and yield_desejado > 0) else 0.0
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
            c_g1.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Tipo', title='Alocação por Tipo', hole=0.4), use_container_width=True)
            c_g2.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Ativo', title='Alocação por Ativo'), use_container_width=True)
            c_g3.plotly_chart(px.pie(df_perf_final, values='Saldo Atual', names='Setor', title='Alocação por Setores/Classe'), use_container_width=True)

        # ==========================================
        # ABA 5: RECOMENDAÇÕES E PROJEÇÃO REALINHADA
        # ==========================================
        with tab5:
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
            
            with st.expander("❓ Entender os Critérios das Recomendações (Cérebro da Analista)"):
                st.markdown("""
                **Modelo de Decisão CNPI Interativo:**
                * **COMPRA FORTE 🟢:** O ativo possui Margem de Segurança Graham acima de 15% (subavaliado patrimonialmente) E Margem Bazin acima de 5% (ótimo fluxo de dividendos).
                * **MANTER / COMPRA 🟡:** O preço de mercado está equilibrado com o valor intrínseco contábil ou teto de proventos.
                * **AVALIAR VENDA 🔴:** O ativo rompeu o Preço Teto de Bazin ou está cotado substancialmente acima do Valor Justo de Graham.
                """)
            st.dataframe(df_recs, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("### ❄️ Projeção Bola de Neve (Próximos 12 Meses)")
            st.markdown("Insira o valor do seu aporte mensal recorrente abaixo. O simulador calculará o crescimento do seu patrimônio somando o capital novo e o reinvestimento de dividendos compostos.")
            
            # 🔥 Entrada do Aporte Mensal conforme solicitado
            aporte_mensal_planejado = st.number_input("Aporte Mensal Planejado (R$):", value=2000.0, step=500.0, min_value=0.0, key="aporte_bola_neve")
            
            saldo_total_atual = df_perf_final['Saldo Atual'].sum()
            div_total_12m = st.session_state.df_simul['Div. Projetado (R$)'].sum()
            yield_mensal = (div_total_12m / saldo_total_atual) / 12 if saldo_total_atual > 0 else 0
            
            meses_lista = []
            patr_inicial_acum = []
            aportes_acumulados = []
            dividendos_reinvestidos = []
            
            saldo_corrente = saldo_total_atual
            acum_aportes = 0.0
            acum_divs = 0.0
            
            for i in range(13):
                meses_lista.append(f"Mês {i}")
                patr_inicial_acum.append(saldo_total_atual)
                aportes_acumulados.append(acum_aportes)
                dividendos_reinvestidos.append(acum_divs)
                
                # Juros compostos aplicados à base
                rendimento_mes = saldo_corrente * yield_mensal
                acum_divs += rendimento_mes
                acum_aportes += aporte_mensal_planejado
                saldo_corrente += (rendimento_mes + aporte_mensal_planejado)
                
            df_proj_real = pd.DataFrame({
                "Mês": meses_lista,
                "Patrimônio Inicial": patr_inicial_acum,
                "Aportes Acumulados": aportes_acumulados,
                "Rendimento de Dividendos": dividendos_reinvestidos
            })
            
            # 🔥 Gráfico reformulado para Área Empilhada (Evita o efeito de barra plana)
            fig_proj = px.area(
                df_proj_real, x="Mês", 
                y=["Patrimônio Inicial", "Aportes Acumulados", "Rendimento de Dividendos"],
                title="Evolução Patrimonial Composta com Aportes e Reinvestimento",
                color_discrete_sequence=["#2c3e50", "#2980b9", "#27ae60"]
            )
            
            # Janela vertical inteligente para destacar a inclinação da curva
            min_y = saldo_total_atual * 0.98
            max_y = (saldo_total_atual + acum_aportes + acum_divs) * 1.02
            fig_proj.update_yaxes(range=[min_y, max_y], title="Montante Projetado (R$)")
            fig_proj.update_layout(xaxis_title="Linha do Tempo", legend_title="Componentes", margin=dict(t=40))
            st.plotly_chart(fig_proj, use_container_width=True)

        with tab6:
            todos_ativos = df_perf_final['Ativo'].tolist()
            c_sel, c_ind = st.columns([2, 1])
            with c_sel: ativos_selecionados = st.multiselect("Selecione os ativos:", todos_ativos, default=todos_ativos[:6], key="ms_g")
            with c_ind: ind_selecionados = st.multiselect("Indicadores:", ["Evolução c/ Div", "CDI Acum.", "IPCA Acum."], default=["Evolução c/ Div", "CDI Acum.", "IPCA Acum."], key="ind_g")
            
            if ativos_selecionados and ind_selecionados:
                df_grafico = df_perf_final[df_perf_final['Ativo'].isin(ativos_selecionados)].copy()
                df_grafico['Período'] = df_grafico['Data Média'].astype(str) + " até Hoje"
                df_melt = df_grafico.melt(id_vars=["Ativo", "Período"], value_vars=ind_selecionados, var_name="Indicador", value_name="Rentabilidade")
                
                fig1 = px.bar(df_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group", hover_data=["Período"])
                fig1.update_traces(hovertemplate='<b>%{x}</b> (%{data.name})<br>Período: %{customdata[0]}<br>Rentabilidade: %{y:.2f}%<extra></extra>')
                fig1.update_layout(yaxis_ticksuffix=" %", margin=dict(t=40))
                st.plotly_chart(fig1, use_container_width=True)
