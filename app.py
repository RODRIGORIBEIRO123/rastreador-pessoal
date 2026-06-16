import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re

# ==========================================
# 1. CONFIGURAÇÃO E MEMÓRIA (STATE)
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")
st.title("📊 Terminal de Gestão Institucional")

# Inicializando a memória para não perder os dados ao clicar em botões
if 'processado' not in st.session_state:
    st.session_state.processado = False
if 'df_base' not in st.session_state:
    st.session_state.df_base = pd.DataFrame()
if 'dados_mercado' not in st.session_state:
    st.session_state.dados_mercado = {}
if 'df_simul' not in st.session_state:
    st.session_state.df_simul = pd.DataFrame()

@st.cache_data(ttl=86400)
def carregar_macro():
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'], macro['IPCA'] = macro['CDI'] / 100, macro['IPCA'] / 100
        return macro
    except:
        return pd.DataFrame()

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
# 2. MOTOR DE EXTRAÇÃO DA B3
# ==========================================
arquivo = st.file_uploader("1. Faça o upload da planilha da B3 (.xlsx ou .csv)", type=["xlsx", "csv"])

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
            ticker = str(row['Código de Negociação']).strip().upper()
            ticker = ticker[:-1] if ticker.endswith('F') else ticker
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
                    "1º Aporte": d['primeiro_aporte'].date() if pd.notna(d['primeiro_aporte']) else pd.Timestamp.now().date()
                })
        
        st.session_state.df_base = pd.DataFrame(ativos_limpos)
        st.session_state.processado = True

# ==========================================
# 3. INTERFACE DE EDIÇÃO HÍBRIDA
# ==========================================
if st.session_state.processado:
    st.write("---")
    st.subheader("2. Auditoria e Correção da Carteira")
    st.markdown("""
    * **Para Adicionar:** Role até o final da tabela e clique no `+` (ou clique na última linha cinza).
    * **Para Excluir:** Selecione a caixa à esquerda do ativo e aperte a tecla `Delete` no seu teclado.
    * **Para Editar:** Clique no número e digite a Quantidade ou o Preço Médio correto.
    """)
    
    # O Editor de Dados armazena as alterações em 'df_editado' (com num_rows='dynamic' ativado)
    df_editado = st.data_editor(
        st.session_state.df_base, 
        use_container_width=True, 
        num_rows="dynamic",
        key="editor_carteira",
        column_config={
            "Ativo": st.column_config.TextColumn("Ativo", required=True),
            "Quantidade": st.column_config.NumberColumn("Quantidade", min_value=0.0),
            "Preço Médio": st.column_config.NumberColumn("Preço Médio", format="R$ %.2f", min_value=0.0),
            "1º Aporte": st.column_config.DateColumn("1º Aporte", format="DD/MM/YYYY")
        }
    )

    # Botão para processar os dados ajustados
    if st.button("🚀 Executar Integração e Calcular Valuation", type="primary"):
        df_macro = carregar_macro()
        progresso = st.progress(0)
        total = len(df_editado)
        data_12m = pd.Timestamp.now() - pd.DateOffset(years=1)
        
        dados_mercado = {}
        linhas_simul_iniciais = []

        for i, row in df_editado.iterrows():
            ticker = str(row.get('Ativo', '')).strip().upper()
            if not ticker or pd.isna(row.get('Quantidade', 0)) or row.get('Quantidade', 0) <= 0: continue
            
            try:
                acao = yf.Ticker(f"{ticker}.SA")
                hist = acao.history(period="1d")
                preco_atual = float(hist['Close'].iloc[-1]) if not hist.empty else float(row['Preço Médio'])
                
                divs = acao.dividends
                data_compra = pd.to_datetime(row['1º Aporte']) if pd.notna(row['1º Aporte']) else pd.Timestamp.now()
                divs_total = float(divs[divs.index.tz_localize(None) >= data_compra].sum() * row['Quantidade'])
                divs_12m = float(divs[divs.index.tz_localize(None) >= data_12m].sum())
                
                info = acao.info
                lpa = float(info.get('trailingEps', 0) or 0.0)
                vpa = float(info.get('bookValue', 0) or 0.0)
                
            except:
                preco_atual, divs_total, divs_12m, lpa, vpa = float(row['Preço Médio']), 0.0, 0.0, 0.0, 0.0

            cdi, ipca = calcular_macro_acumulado(df_macro, data_compra)
            
            dados_mercado[ticker] = {
                "Qtd": float(row['Quantidade']),
                "PM": float(row['Preço Médio']),
                "Preço Atual": preco_atual,
                "Div_Total": divs_total,
                "CDI": cdi,
                "IPCA": ipca
            }

            # Prepara a base para o simulador de Valuation
            linhas_simul_iniciais.append({
                "Ativo": ticker,
                "Cotação Atual": preco_atual,
                "VPA (Constante)": vpa,
                "LPA Projetado": lpa,
                "Div. Projetado (R$)": divs_12m,
                "Yield Desejado (%)": 6.0  # Padrão Bazin editável
            })
            
            progresso.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dados_mercado
        st.session_state.df_simul = pd.DataFrame(linhas_simul_iniciais)
        st.success("Cotações, Proventos e Fundamentos sincronizados com sucesso!")

    # ==========================================
    # 4. PAINEL DE RELATÓRIOS (ABAS)
    # ==========================================
    if st.session_state.dados_mercado:
        tab1, tab2 = st.tabs(["📈 Evolução de Rentabilidade", "🔎 Simulador Valuation (Graham & Bazin)"])
        
        # --- ABA 1: RENTABILIDADE (Com Ordenação) ---
        with tab1:
            st.markdown("### Performance Histórica (Real-Time)")
            linhas_perf = []
            for t, dm in st.session_state.dados_mercado.items():
                investido = dm['Qtd'] * dm['PM']
                saldo = dm['Qtd'] * dm['Preço Atual']
                var_s_div = ((saldo / investido) - 1) * 100 if investido > 0 else np.nan
                var_c_div = (((saldo + dm['Div_Total']) / investido) - 1) * 100 if investido > 0 else np.nan
                
                linhas_perf.append({
                    "Ativo": t,
                    "Qtd": int(dm['Qtd']),
                    "Preço Médio": dm['PM'],
                    "Preço Atual": dm['Preço Atual'],
                    "Evol. s/ Div": var_s_div,
                    "Evol. c/ Div": var_c_div,
                    "IPCA Acum.": dm['IPCA'],
                    "CDI Acum.": dm['CDI']
                })
            
            cfg_perf = {
                "Preço Médio": st.column_config.NumberColumn("Preço Médio", format="R$ %.2f"),
                "Preço Atual": st.column_config.NumberColumn("Preço Atual", format="R$ %.2f"),
                "Evol. s/ Div": st.column_config.NumberColumn("Evol. s/ Div", format="%.2f %%"),
                "Evol. c/ Div": st.column_config.NumberColumn("Evol. c/ Div", format="%.2f %%"),
                "IPCA Acum.": st.column_config.NumberColumn("IPCA", format="%.2f %%"),
                "CDI Acum.": st.column_config.NumberColumn("CDI", format="%.2f %%")
            }
            # DataFrame puramente Numérico -> Ordenação liberada!
            st.dataframe(pd.DataFrame(linhas_perf), use_container_width=True, hide_index=True, column_config=cfg_perf)

        # --- ABA 2: SIMULADOR DE VALUATION ---
        with tab2:
            st.markdown("### 1. Inserção de Parâmetros Projetivos")
            st.markdown("Altere os campos **LPA Projetado**, **Div. Projetado (R$)** ou o **Yield Desejado (%)** (Bazin). A tabela de resultados abaixo será recalculada automaticamente.")
            
            # Tabela de Input (Editável)
            df_simul_editado = st.data_editor(
                st.session_state.df_simul,
                use_container_width=True,
                hide_index=True,
                disabled=["Ativo", "Cotação Atual", "VPA (Constante)"],
                column_config={
                    "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                    "VPA (Constante)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "LPA Projetado": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Div. Projetado (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Yield Desejado (%)": st.column_config.NumberColumn(format="%.2f %%", min_value=0.1)
                }
            )
            
            # Motor de Cálculo (Reativo às edições acima)
            linhas_valuation = []
            for _, row in df_simul_editado.iterrows():
                t = row['Ativo']
                cotacao = row['Cotação Atual']
                vpa = row['VPA (Constante)']
                lpa_proj = row['LPA Projetado']
                div_proj = row['Div. Projetado (R$)']
                yield_desejado = row['Yield Desejado (%)'] / 100.0  # Converte 6% para 0.06
                
                # Graham (Impede cálculo negativo)
                graham = (22.5 * lpa_proj * vpa) ** 0.5 if (lpa_proj > 0 and vpa > 0) else np.nan
                margem_g = ((graham / cotacao) - 1) * 100 if (pd.notna(graham) and cotacao > 0) else np.nan
                
                # Bazin (Cálculo dinâmico com a taxa do usuário)
                bazin = div_proj / yield_desejado if (div_proj > 0 and yield_desejado > 0) else np.nan
                margem_b = ((bazin / cotacao) - 1) * 100 if (pd.notna(bazin) and cotacao > 0) else np.nan
                
                linhas_valuation.append({
                    "Ativo": t,
                    "Preço Justo (Graham)": graham,
                    "Margem Graham": margem_g,
                    "Preço Teto (Bazin)": bazin,
                    "Margem Bazin": margem_b
                })
                
            cfg_val = {
                "Preço Justo (Graham)": st.column_config.NumberColumn("Preço Justo (Graham)", format="R$ %.2f"),
                "Margem Graham": st.column_config.NumberColumn("Margem Graham", format="%.2f %%"),
                "Preço Teto (Bazin)": st.column_config.NumberColumn("Preço Teto (Bazin)", format="R$ %.2f"),
                "Margem Bazin": st.column_config.NumberColumn("Margem Bazin", format="%.2f %%")
            }
            
            st.markdown("### 2. Resultados do Valuation (Ordenável)")
            st.dataframe(pd.DataFrame(linhas_valuation), use_container_width=True, hide_index=True, column_config=cfg_val)
