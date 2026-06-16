import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re

# ==========================================
# 1. CONFIGURAÇÃO E MEMÓRIA
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")
st.title("📊 Terminal de Gestão Profissional")

if 'df_base' not in st.session_state: st.session_state.df_base = pd.DataFrame(columns=["Ativo", "Quantidade", "Preço Médio", "Data 1º Aporte"])
if 'dados_mercado' not in st.session_state: st.session_state.dados_mercado = {}
if 'df_simul' not in st.session_state: st.session_state.df_simul = pd.DataFrame()

@st.cache_data(ttl=86400)
def carregar_macro():
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'], macro['IPCA'] = macro['CDI'] / 100, macro['IPCA'] / 100
        return macro
    except: return pd.DataFrame()

def calcular_macro_acumulado(df_macro, data_inicio):
    if df_macro is None or df_macro.empty or pd.isna(data_inicio): return 0.0, 0.0
    try:
        filtro = df_macro.loc[data_inicio:]
        cdi_acum = ((1 + filtro['CDI'].dropna()).prod() - 1) * 100
        ipca_acum = ((1 + filtro['IPCA'].dropna()).prod() - 1) * 100
        return cdi_acum, ipca_acum
    except: return 0.0, 0.0

def ignorar_ativo(ticker):
    t = str(ticker).strip().upper()
    if t.startswith(('WIN', 'WDO', 'IND', 'DOL')) or pd.isna(ticker): return True
    t_limpo = t[:-1] if t.endswith('F') else t
    if re.match(r'^[A-Z]{4}[A-Z]\d+', t_limpo) and not t_limpo.endswith(('11', '34', '39')):
        if len(t_limpo) >= 6: return True
    return False

# ==========================================
# 2. UPLOAD E LEITURA (MOTOR BLINDADO)
# ==========================================
st.sidebar.header("1. Upload da B3")
arquivo = st.sidebar.file_uploader("Arquivo Negociação B3", type=["xlsx", "csv"])

if arquivo and st.session_state.df_base.empty:
    with st.spinner("Processando histórico..."):
        # Leitor inteligente que aceita CSV com Vírgula ou Ponto-e-Vírgula
        if arquivo.name.endswith('.csv'):
            try:
                df = pd.read_csv(arquivo, sep=';', encoding='latin1')
                if 'Data do Negócio' not in df.columns:
                    arquivo.seek(0)
                    df = pd.read_csv(arquivo, sep=',', encoding='utf-8')
            except:
                arquivo.seek(0)
                df = pd.read_csv(arquivo, sep=',', encoding='utf-8')
        else:
            df = pd.read_excel(arquivo)
            
        df.columns = df.columns.str.strip()
        df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], format='%d/%m/%Y', errors='coerce')
        df['Preço'] = df['Preço'].astype(str).replace({r'R\$': '', r'\.': '', ',': '.'}, regex=True).astype(float)
        df['Valor'] = df['Valor'].astype(str).replace({r'R\$': '', r'\.': '', ',': '.'}, regex=True).astype(float)
        df = df.sort_values('Data do Negócio')
        
        posicoes = {}
        for _, row in df.iterrows():
            if ignorar_ativo(row['Código de Negociação']): continue
            ticker = str(row['Código de Negociação']).strip().upper()
            ticker = ticker[:-1] if ticker.endswith('F') else ticker
            qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio']
            
            if ticker not in posicoes: posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'primeiro_aporte': data}
            if row['Tipo de Movimentação'] == 'Compra':
                posicoes[ticker]['qtd'] += qtd
                posicoes[ticker]['valor'] += valor
                if pd.notna(data) and data < posicoes[ticker]['primeiro_aporte']: posicoes[ticker]['primeiro_aporte'] = data
            elif row['Tipo de Movimentação'] == 'Venda' and posicoes[ticker]['qtd'] > 0:
                qtd_venda = min(qtd, posicoes[ticker]['qtd'])
                pm_atual = posicoes[ticker]['valor'] / posicoes[ticker]['qtd']
                posicoes[ticker]['qtd'] -= qtd_venda
                posicoes[ticker]['valor'] -= (qtd_venda * pm_atual)

        ativos_limpos = []
        for t, d in posicoes.items():
            if d['qtd'] > 0:
                ativos_limpos.append({
                    "Ativo": t, "Quantidade": float(d['qtd']), 
                    "Preço Médio": float(d['valor'] / d['qtd']), 
                    "Data 1º Aporte": d['primeiro_aporte'].date() if pd.notna(d['primeiro_aporte']) else pd.Timestamp.now().date()
                })
        st.session_state.df_base = pd.DataFrame(ativos_limpos)
        st.rerun()

# ==========================================
# 3. INTERFACE DE EDIÇÃO MANUAL
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Edição Manual da Carteira")
    st.markdown("Ajuste as quantidades, adicione os ativos faltantes ou exclua o que já não possui.")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.error("🗑️ EXCLUIR ATIVO")
        lista_ativos = [""] + sorted(st.session_state.df_base["Ativo"].tolist())
        ticker_del = st.selectbox("Selecione o ativo para remover:", lista_ativos)
        if st.button("Remover Ativo Selecionado"):
            if ticker_del != "":
                st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != ticker_del]
                st.rerun()
                
    with col2:
        st.success("➕ INCLUIR NOVO ATIVO")
        c1, c2, c3 = st.columns(3)
        novo_t = c1.text_input("Ticker (Ex: BBAS3)")
        novo_q = c2.number_input("Qtd", min_value=1)
        novo_p = c3.number_input("PM (R$)", min_value=0.01)
        if st.button("Adicionar à Carteira"):
            if novo_t != "":
                nova_linha = pd.DataFrame([{"Ativo": novo_t.upper(), "Quantidade": float(novo_q), "Preço Médio": float(novo_p), "Data 1º Aporte": pd.Timestamp.now().date()}])
                st.session_state.df_base = pd.concat([st.session_state.df_base, nova_linha], ignore_index=True)
                st.rerun()

    # Tabela Centralizada para Correção Rápida
    df_editado = st.data_editor(
        st.session_state.df_base, use_container_width=True, hide_index=True,
        column_config={
            "Ativo": st.column_config.TextColumn(disabled=True),
            "Quantidade": st.column_config.NumberColumn(min_value=0.0),
            "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0),
            "Data 1º Aporte": st.column_config.DateColumn(format="DD/MM/YYYY")
        }
    )

    if st.button("🚀 Processar Conexão com o Mercado", type="primary"):
        st.session_state.df_base = df_editado # Salva as edições
        df_macro = carregar_macro()
        progresso = st.progress(0)
        total = len(df_editado)
        data_12m = pd.Timestamp.now() - pd.DateOffset(years=1)
        
        dados_mercado = {}
        linhas_simul_iniciais = []

        for i, row in df_editado.iterrows():
            ticker = str(row['Ativo']).strip().upper()
            try:
                acao = yf.Ticker(f"{ticker}.SA")
                hist = acao.history(period="1d")
                preco_atual = float(hist['Close'].iloc[-1]) if not hist.empty else float(row['Preço Médio'])
                
                divs = acao.dividends
                data_compra = pd.to_datetime(row['Data 1º Aporte'])
                divs_total = float(divs[divs.index.tz_localize(None) >= data_compra].sum() * row['Quantidade'])
                divs_12m = float(divs[divs.index.tz_localize(None) >= data_12m].sum())
                
                info = acao.info
                lpa, vpa = float(info.get('trailingEps', 0) or 0.0), float(info.get('bookValue', 0) or 0.0)
            except:
                preco_atual, divs_total, divs_12m, lpa, vpa = float(row['Preço Médio']), 0.0, 0.0, 0.0, 0.0

            cdi, ipca = calcular_macro_acumulado(df_macro, data_compra)
            
            dados_mercado[ticker] = {
                "Qtd": float(row['Quantidade']), "PM": float(row['Preço Médio']),
                "Preço Atual": preco_atual, "Div_Total": divs_total, "CDI": cdi, "IPCA": ipca
            }

            linhas_simul_iniciais.append({
                "Ativo": ticker, "Cotação Atual": preco_atual, "VPA (Contábil)": vpa,
                "LPA Projetado": lpa, "Div. Projetado (R$)": divs_12m
            })
            progresso.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dados_mercado
        st.session_state.df_simul = pd.DataFrame(linhas_simul_iniciais)
        st.success("Dados sincronizados!")

    # ==========================================
    # 4. PAINEL DE RELATÓRIOS (ABAS)
    # ==========================================
    if st.session_state.dados_mercado:
        st.write("---")
        tab1, tab2 = st.tabs(["📈 Rentabilidade e Comparativos", "🔎 Simulador de Valuation e Preço Teto"])
        
        # ABA 1
        with tab1:
            linhas_perf = []
            for t, dm in st.session_state.dados_mercado.items():
                investido = dm['Qtd'] * dm['PM']
                saldo = dm['Qtd'] * dm['Preço Atual']
                linhas_perf.append({
                    "Ativo": t, "Qtd": int(dm['Qtd']), "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
                    "Evolução s/ Div": ((saldo / investido) - 1) * 100 if investido > 0 else np.nan,
                    "Evolução c/ Div": (((saldo + dm['Div_Total']) / investido) - 1) * 100 if investido > 0 else np.nan,
                    "IPCA Acum.": dm['IPCA'], "CDI Acum.": dm['CDI']
                })
            
            st.dataframe(pd.DataFrame(linhas_perf), use_container_width=True, hide_index=True, column_config={
                "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f"),
                "Preço Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Evolução s/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "Evolução c/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "IPCA Acum.": st.column_config.NumberColumn(format="%.2f %%"),
                "CDI Acum.": st.column_config.NumberColumn(format="%.2f %%")
            })

        # ABA 2
        with tab2:
            st.markdown("### Simulador de Aportes")
            yield_desejado = st.number_input("Digite o seu Dividend Yield Desejado (% Bazin):", value=6.0, min_value=0.1, step=0.5) / 100.0
            
            st.markdown("Altere o LPA ou os Dividendos abaixo e veja as margens reagirem instantaneamente:")
            df_simul_editado = st.data_editor(
                st.session_state.df_simul, use_container_width=True, hide_index=True,
                disabled=["Ativo", "Cotação Atual", "VPA (Contábil)"],
                column_config={
                    "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                    "VPA (Contábil)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "LPA Projetado": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Div. Projetado (R$)": st.column_config.NumberColumn(format="R$ %.2f")
                }
            )
            
            linhas_val = []
            for _, row in df_simul_editado.iterrows():
                t = row['Ativo']
                cotacao, vpa = row['Cotação Atual'], row['VPA (Contábil)']
                lpa_proj, div_proj = row['LPA Projetado'], row['Div. Projetado (R$)']
                
                graham = (22.5 * lpa_proj * vpa) ** 0.5 if (lpa_proj > 0 and vpa > 0) else np.nan
                margem_g = ((graham / cotacao) - 1) * 100 if (pd.notna(graham) and cotacao > 0) else np.nan
                bazin = div_proj / yield_desejado if div_proj > 0 else np.nan
                margem_b = ((bazin / cotacao) - 1) * 100 if (pd.notna(bazin) and cotacao > 0) else np.nan
                
                linhas_val.append({"Ativo": t, "Preço Graham": graham, "Margem Graham": margem_g, "Preço Bazin": bazin, "Margem Bazin": margem_b})
                
            st.markdown("### Resultado Matemático (Clique na coluna para ordenar)")
            st.dataframe(pd.DataFrame(linhas_val), use_container_width=True, hide_index=True, column_config={
                "Preço Graham": st.column_config.NumberColumn("Preço Teto Graham", format="R$ %.2f"),
                "Margem Graham": st.column_config.NumberColumn("Margem Graham", format="%.2f %%"),
                "Preço Bazin": st.column_config.NumberColumn("Preço Teto Bazin", format="R$ %.2f"),
                "Margem Bazin": st.column_config.NumberColumn("Margem Bazin", format="%.2f %%")
            })
