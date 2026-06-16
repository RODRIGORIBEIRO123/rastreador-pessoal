import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re

# ==========================================
# 1. CONFIGURAÇÃO E CACHE DE DADOS
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")
st.title("📊 Terminal de Gestão Institucional")

# Inicializa o cache na memória para evitar recarregamento desnecessário do Yahoo Finance
if 'dados_mercado' not in st.session_state:
    st.session_state.dados_mercado = {}
    st.session_state.processado = False
    st.session_state.df_base = pd.DataFrame()

@st.cache_data(ttl=86400)
def carregar_macro():
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'], macro['IPCA'] = macro['CDI'] / 100, macro['IPCA'] / 100
        return macro
    except:
        return pd.DataFrame()

def calcular_macro_acumulado(df_macro, data_inicio):
    if df_macro is None or df_macro.empty or pd.isna(data_inicio): 
        return 0.0, 0.0
    try:
        filtro = df_macro.loc[data_inicio:]
        cdi_acum = ((1 + filtro['CDI'].dropna()).prod() - 1) * 100
        ipca_acum = ((1 + filtro['IPCA'].dropna()).prod() - 1) * 100
        return cdi_acum, ipca_acum
    except:
        return 0.0, 0.0

def limpar_ticker(ticker):
    t = str(ticker).strip().upper()
    return t[:-1] if t.endswith('F') else t

def ignorar_ativo(ticker):
    t = str(ticker).strip().upper()
    if t.startswith(('WIN', 'WDO', 'IND', 'DOL')) or pd.isna(ticker): return True
    t_limpo = t[:-1] if t.endswith('F') else t
    if re.match(r'^[A-Z]{4}[A-Z]\d+', t_limpo) and not t_limpo.endswith(('11', '34', '39')):
        if len(t_limpo) >= 6: return True
    return False

# ==========================================
# 2. MOTOR DE EXTRAÇÃO DA B3
# ==========================================
st.sidebar.header("📁 Entrada de Dados")
arquivo = st.sidebar.file_uploader("Upload da planilha da B3", type=["xlsx", "csv"])

if arquivo and not st.session_state.processado:
    with st.spinner("Estruturando banco de dados da B3..."):
        df = pd.read_csv(arquivo, sep=';', encoding='latin1') if arquivo.name.endswith('.csv') else pd.read_excel(arquivo)
        df.columns = df.columns.str.strip()
        df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], format='%d/%m/%Y', errors='coerce')
        df['Preço'] = df['Preço'].astype(str).replace({r'R\$': '', r'\.': '', ',': '.'}, regex=True).astype(float)
        df['Valor'] = df['Valor'].astype(str).replace({r'R\$': '', r'\.': '', ',': '.'}, regex=True).astype(float)
        df = df.sort_values('Data do Negócio')
        
        posicoes = {}
        for _, row in df.iterrows():
            if ignorar_ativo(row['Código de Negociação']): continue
            ticker = limpar_ticker(row['Código de Negociação'])
            qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio']
            
            if ticker not in posicoes:
                posicoes[ticker] = {'qtd': 0.0, 'valor_investido': 0.0, 'primeiro_aporte': data}
                
            if row['Tipo de Movimentação'] == 'Compra':
                posicoes[ticker]['qtd'] += qtd
                posicoes[ticker]['valor_investido'] += valor
                if pd.notna(data) and data < posicoes[ticker]['primeiro_aporte']: 
                    posicoes[ticker]['primeiro_aporte'] = data
            elif row['Tipo de Movimentação'] == 'Venda' and posicoes[ticker]['qtd'] > 0:
                qtd_venda = min(qtd, posicoes[ticker]['qtd'])
                pm_atual = posicoes[ticker]['valor_investido'] / posicoes[ticker]['qtd']
                posicoes[ticker]['qtd'] -= qtd_venda
                posicoes[ticker]['valor_investido'] -= (qtd_venda * pm_atual)
                if posicoes[ticker]['qtd'] <= 0.001: posicoes[ticker]['qtd'] = 0

        ativos_limpos = []
        for t, d in posicoes.items():
            if d['qtd'] > 0:
                pm = d['valor_investido'] / d['qtd']
                ativos_limpos.append({
                    "Ativo": t, 
                    "Quantidade": float(d['qtd']), 
                    "Preço Médio": float(pm), 
                    "1º Aporte": d['primeiro_aporte']
                })
        
        st.session_state.df_base = pd.DataFrame(ativos_limpos)
        st.session_state.processado = True

# ==========================================
# 3. INTERFACE DE GESTÃO HÍBRIDA
# ==========================================
if st.session_state.processado:
    st.markdown("### 1. Auditoria e Correção de Posição")
    st.markdown("Edite as quantidades e preços médios diretamente na tabela abaixo para alinhar com sua corretora.")
    
    # Data Editor nativo (Garante que os tipos de dados não quebrem)
    df_editado = st.data_editor(
        st.session_state.df_base, 
        use_container_width=True, 
        num_rows="dynamic",
        column_config={
            "1º Aporte": st.column_config.DateColumn("1º Aporte", format="DD/MM/YYYY"),
            "Preço Médio": st.column_config.NumberColumn("Preço Médio", format="R$ %.2f", min_value=0.0)
        }
    )

    if st.button("🔄 Executar Integração com Mercado (Real-Time)", type="primary"):
        df_macro = carregar_macro()
        progresso = st.progress(0)
        total = len(df_editado)
        data_12m = pd.Timestamp.now() - pd.DateOffset(years=1)
        
        dados_mercado = {}
        for i, row in df_editado.iterrows():
            ticker = str(row['Ativo']).strip().upper()
            if not ticker or pd.isna(row['Quantidade']) or row['Quantidade'] <= 0: continue
            
            # Fetching YFinance Data
            try:
                acao = yf.Ticker(f"{ticker}.SA")
                hist = acao.history(period="1d")
                preco_atual = hist['Close'].iloc[-1] if not hist.empty else float(row['Preço Médio'])
                
                divs = acao.dividends
                data_compra = pd.to_datetime(row['1º Aporte']) if pd.notna(row['1º Aporte']) else pd.Timestamp.now()
                
                divs_total = divs[divs.index.tz_localize(None) >= data_compra].sum() * float(row['Quantidade'])
                divs_12m = divs[divs.index.tz_localize(None) >= data_12m].sum()
                
                info = acao.info
                lpa = info.get('trailingEps', 0)
                vpa = info.get('bookValue', 0)
                
            except:
                preco_atual, divs_total, divs_12m, lpa, vpa = float(row['Preço Médio']), 0.0, 0.0, 0.0, 0.0

            cdi, ipca = calcular_macro_acumulado(df_macro, data_compra)
            
            dados_mercado[ticker] = {
                "Preço Atual": preco_atual,
                "Div_Total": divs_total,
                "Div_12m": divs_12m,
                "LPA": lpa if lpa is not None else 0.0,
                "VPA": vpa if vpa is not None else 0.0,
                "CDI": cdi,
                "IPCA": ipca
            }
            progresso.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dados_mercado
        st.success("Integração concluída com sucesso!")

    # ==========================================
    # 4. PAINEL DE RELATÓRIOS (ABAS)
    # ==========================================
    if st.session_state.dados_mercado:
        tab1, tab2 = st.tabs(["📈 Evolução de Rentabilidade", "🔎 Simulador Valuation (Graham & Bazin)"])
        
        # --- ABA 1: RENTABILIDADE ---
        with tab1:
            linhas_perf = []
            for _, row in df_editado.iterrows():
                t = str(row['Ativo']).strip().upper()
                if t not in st.session_state.dados_mercado: continue
                
                qtd = float(row['Quantidade'])
                pm = float(row['Preço Médio'])
                dm = st.session_state.dados_mercado[t]
                
                investido = qtd * pm
                saldo = qtd * dm['Preço Atual']
                var_s_div = ((saldo / investido) - 1) * 100 if investido > 0 else 0
                var_c_div = (((saldo + dm['Div_Total']) / investido) - 1) * 100 if investido > 0 else 0
                
                linhas_perf.append({
                    "Ativo": t,
                    "Qtd": int(qtd),
                    "Preço Médio": pm,
                    "Preço Atual": dm['Preço Atual'],
                    "Evolução (S/ Div)": var_s_div,
                    "Evolução (C/ Div)": var_c_div,
                    "IPCA Acum.": dm['IPCA'],
                    "CDI Acum.": dm['CDI']
                })
            
            df_perf = pd.DataFrame(linhas_perf)
            
            # Configuração nativa de colunas para formatação + ordenação
            cfg_perf = {
                "Preço Médio": st.column_config.NumberColumn("Preço Médio", format="R$ %.2f"),
                "Preço Atual": st.column_config.NumberColumn("Preço Atual", format="R$ %.2f"),
                "Evolução (S/ Div)": st.column_config.NumberColumn("Evol. s/ Div", format="%.2f %%"),
                "Evolução (C/ Div)": st.column_config.NumberColumn("Evol. c/ Div", format="%.2f %%"),
                "IPCA Acum.": st.column_config.NumberColumn("IPCA", format="%.2f %%"),
                "CDI Acum.": st.column_config.NumberColumn("CDI", format="%.2f %%")
            }
            st.dataframe(df_perf, use_container_width=True, hide_index=True, column_config=cfg_perf)

        # --- ABA 2: SIMULADOR DE VALUATION ---
        with tab2:
            st.markdown("### Simulador Projetivo")
            st.markdown("Altere os campos **LPA Projetado** e **Div. Projetado** para simular novos cenários. O cálculo de Preço Teto e Margem atualizará automaticamente na tabela abaixo.")
            
            linhas_simul = []
            for _, row in df_editado.iterrows():
                t = str(row['Ativo']).strip().upper()
                if t not in st.session_state.dados_mercado: continue
                dm = st.session_state.dados_mercado[t]
                
                linhas_simul.append({
                    "Ativo": t,
                    "VPA (Constante)": float(dm['VPA']),
                    "LPA Projetado": float(dm['LPA']),
                    "Div. Projetado (12m)": float(dm['Div_12m'])
                })
            
            df_simul_input = pd.DataFrame(linhas_simul)
            
            # Tabela de input para o usuário digitar os lucros projetados
            df_simul_output = st.data_editor(
                df_simul_input,
                use_container_width=True,
                hide_index=True,
                disabled=["Ativo", "VPA (Constante)"],
                column_config={
                    "VPA (Constante)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "LPA Projetado": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0),
                    "Div. Projetado (12m)": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0)
                }
            )
            
            # Executa a matemática em tempo real baseada nos inputs do usuário
            linhas_valuation = []
            for _, row in df_simul_output.iterrows():
                t = row['Ativo']
                vpa = row['VPA (Constante)']
                lpa_proj = row['LPA Projetado']
                div_proj = row['Div. Projetado (12m)']
                preco_atual = st.session_state.dados_mercado[t]['Preço Atual']
                
                # Fórmulas
                graham = (22.5 * lpa_proj * vpa) ** 0.5 if (lpa_proj > 0 and vpa > 0) else np.nan
                margem_g = ((graham / preco_atual) - 1) * 100 if (pd.notna(graham) and preco_atual > 0) else np.nan
                
                bazin = div_proj / 0.06 if div_proj > 0 else np.nan
                margem_b = ((bazin / preco_atual) - 1) * 100 if (pd.notna(bazin) and preco_atual > 0) else np.nan
                
                linhas_valuation.append({
                    "Ativo": t,
                    "Cotação Atual": preco_atual,
                    "Preço Justo (Graham)": graham,
                    "Margem Graham": margem_g,
                    "Preço Teto (Bazin)": bazin,
                    "Margem Bazin": margem_b
                })
                
            cfg_val = {
                "Cotação Atual": st.column_config.NumberColumn("Cotação Atual", format="R$ %.2f"),
                "Preço Justo (Graham)": st.column_config.NumberColumn("Preço Justo (Graham)", format="R$ %.2f"),
                "Margem Graham": st.column_config.NumberColumn("Margem Graham", format="%.2f %%"),
                "Preço Teto (Bazin)": st.column_config.NumberColumn("Preço Teto (Bazin)", format="R$ %.2f"),
                "Margem Bazin": st.column_config.NumberColumn("Margem Bazin", format="%.2f %%")
            }
            
            st.markdown("### Resultado do Valuation (Ordenável)")
            st.dataframe(pd.DataFrame(linhas_valuation), use_container_width=True, hide_index=True, column_config=cfg_val)
