import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re
import io
from openpyxl.chart import BarChart, Reference

# ==========================================
# 1. CONFIGURAÇÃO E MEMÓRIA
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")
st.title("📊 Terminal de Gestão Profissional")

# Memória do aplicativo
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

def calcular_meses(data_inicio):
    if pd.isna(data_inicio): return 0
    hoje = pd.Timestamp.now()
    return (hoje.year - data_inicio.year) * 12 + (hoje.month - data_inicio.month)

def ignorar_ativo(ticker):
    if pd.isna(ticker): return True
    t = str(ticker).strip().upper()
    if not t or t == 'NAN': return True
    if t.startswith(('WIN', 'WDO', 'IND', 'DOL')): return True
    t_limpo = t[:-1] if t.endswith('F') else t
    if re.match(r'^[A-Z]{4}[A-Z]\d+', t_limpo) and not t_limpo.endswith(('11', '34', '39')):
        if len(t_limpo) >= 6: return True
    return False

def limpar_numero(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float, np.number)): return float(x)
    x = str(x).replace('R$', '').strip()
    if x == '' or x.upper() == 'NAN': return 0.0
    if '.' in x and ',' in x: x = x.replace('.', '').replace(',', '.') 
    elif ',' in x: x = x.replace(',', '.') 
    try: return float(x)
    except: return 0.0

def ler_arquivo_b3(arquivo_upload):
    if arquivo_upload.name.endswith('.csv'):
        conteudo = arquivo_upload.getvalue()
        try: texto = conteudo.decode('utf-8-sig')
        except: texto = conteudo.decode('latin1')
        separador = ';' if ';' in texto.split('\n')[0] else ','
        df = pd.read_csv(io.StringIO(texto), sep=separador)
    else: df = pd.read_excel(arquivo_upload)
    df.columns = df.columns.astype(str).str.strip()
    return df

# ==========================================
# GERAÇÃO DE EXCEL COM GRÁFICO NATIVO (OPENPYXL)
# ==========================================
def gerar_excel_premium(df_perf, df_val):
    output = io.BytesIO()
    # Preenche NaN com 0 para o Excel não quebrar a matemática do gráfico
    df_p = df_perf.fillna(0)
    df_v = df_val.fillna(0)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_p.to_excel(writer, sheet_name='Rentabilidade', index=False)
        df_v.to_excel(writer, sheet_name='Valuation', index=False)
        
        # Desenha o gráfico nativo na aba de Rentabilidade
        ws = writer.sheets['Rentabilidade']
        chart = BarChart()
        chart.type = "col"
        chart.style = 13
        chart.title = "Retorno Total vs Benchmarks (IPCA e CDI)"
        chart.y_axis.title = "Rentabilidade Acumulada (%)"
        chart.x_axis.title = "Ativos da Carteira"
        
        # Colunas: Ativo (1), Qtd (2), PM (3), Atual (4), Meses (5), Var s/ Div (6), Var c/ Div (7), IPCA (8), CDI (9)
        dados_grafico = Reference(ws, min_col=7, min_row=1, max_col=9, max_row=len(df_p)+1)
        categorias = Reference(ws, min_col=1, min_row=2, max_row=len(df_p)+1)
        
        chart.add_data(dados_grafico, titles_from_data=True)
        chart.set_categories(categorias)
        chart.height = 14
        chart.width = 25
        
        ws.add_chart(chart, "L2") # Insere o gráfico ao lado da tabela
    return output.getvalue()

# ==========================================
# 2. UPLOAD E LEITURA DA B3
# ==========================================
st.sidebar.header("1. Upload da B3")
arquivo = st.sidebar.file_uploader("Arquivo Negociação B3", type=["xlsx", "csv"])

if arquivo and st.session_state.df_base.empty:
    with st.spinner("Processando e purificando dados..."):
        try:
            df = ler_arquivo_b3(arquivo)
            df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
            df['Quantidade'] = df['Quantidade'].apply(limpar_numero)
            df['Preço'] = df['Preço'].apply(limpar_numero)
            df['Valor'] = df['Valor'].apply(limpar_numero)
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
        except Exception as e:
            st.error(f"Erro ao ler o arquivo. Detalhe técnico: {e}")

# ==========================================
# 3. INTERFACE DE EDIÇÃO MANUAL
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Edição Manual da Carteira")
    st.markdown("Ajuste as quantidades e preços médios na tabela. Salva automaticamente.")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.error("🗑️ EXCLUIR ATIVO")
        lista_ativos = [""] + sorted(st.session_state.df_base["Ativo"].tolist())
        ticker_del = st.selectbox("Selecione o ativo para remover:", lista_ativos)
        if st.button("Remover Ativo"):
            if ticker_del != "":
                st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != ticker_del]
                st.rerun()
    with col2:
        st.success("➕ INCLUIR NOVO ATIVO")
        c1, c2, c3 = st.columns(3)
        novo_t = c1.text_input("Ticker (Ex: BBAS3)")
        novo_q = c2.number_input("Qtd", min_value=1)
        novo_p = c3.number_input("PM (R$)", min_value=0.01)
        if st.button("Adicionar Ativo"):
            if novo_t != "":
                nova_linha = pd.DataFrame([{"Ativo": novo_t.upper(), "Quantidade": float(novo_q), "Preço Médio": float(novo_p), "Data 1º Aporte": pd.Timestamp.now().date()}])
                st.session_state.df_base = pd.concat([st.session_state.df_base, nova_linha], ignore_index=True)
                st.rerun()

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
        st.session_state.df_base = df_editado 
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
            meses_investido = calcular_meses(data_compra)
            
            dados_mercado[ticker] = {
                "Qtd": float(row['Quantidade']), "PM": float(row['Preço Médio']),
                "Preço Atual": preco_atual, "Div_Total": divs_total, "CDI": cdi, "IPCA": ipca, "Meses": meses_investido
            }

            linhas_simul_iniciais.append({
                "Ativo": ticker, "Cotação Atual": preco_atual, "VPA (Contábil)": vpa,
                "LPA Projetado": lpa, "Div. Projetado (R$)": divs_12m
            })
            progresso.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dados_mercado
        st.session_state.df_simul = pd.DataFrame(linhas_simul_iniciais)
        st.success("Dados de Mercado Sincronizados!")

    # ==========================================
    # 4. PAINEL DE RELATÓRIOS E GRÁFICOS
    # ==========================================
    if st.session_state.dados_mercado:
        # Prepara a base de performance
        linhas_perf = []
        for t, dm in st.session_state.dados_mercado.items():
            investido = dm['Qtd'] * dm['PM']
            saldo = dm['Qtd'] * dm['Preço Atual']
            var_s_div = ((saldo / investido) - 1) * 100 if investido > 0 else np.nan
            var_c_div = (((saldo + dm['Div_Total']) / investido) - 1) * 100 if investido > 0 else np.nan
            
            linhas_perf.append({
                "Ativo": t, "Qtd": int(dm['Qtd']), "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
                "Tempo (Meses)": int(dm['Meses']),
                "Evolução s/ Div": var_s_div, "Evolução c/ Div": var_c_div,
                "IPCA Acum.": dm['IPCA'], "CDI Acum.": dm['CDI']
            })
        df_perf_final = pd.DataFrame(linhas_perf)

        st.write("---")
        # BOTÃO DE EXPORTAÇÃO PREMIUM (Fica acima das abas)
        st.download_button(
            label="📥 Baixar Análise em Excel (Com Gráfico)",
            data=gerar_excel_premium(df_perf_final, st.session_state.df_simul),
            file_name="Analise_CNPI_Carteira.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        tab1, tab2, tab3 = st.tabs(["📈 Rentabilidade", "🔎 Simulador Valuation", "📊 Gráficos Interativos"])
        
        # --- ABA 1: RENTABILIDADE ---
        with tab1:
            st.dataframe(df_perf_final, use_container_width=True, hide_index=True, column_config={
                "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f"),
                "Preço Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Tempo (Meses)": st.column_config.NumberColumn(format="%d"),
                "Evolução s/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "Evolução c/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "IPCA Acum.": st.column_config.NumberColumn(format="%.2f %%"),
                "CDI Acum.": st.column_config.NumberColumn(format="%.2f %%")
            })

        # --- ABA 2: SIMULADOR VALUATION ---
        with tab2:
            st.markdown("### Simulador de Aportes")
            yield_desejado = st.number_input("Taxa de Risco - Yield Desejado (% Bazin):", value=6.0, min_value=0.1, step=0.5) / 100.0
            
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
                t, cotacao, vpa = row['Ativo'], row['Cotação Atual'], row['VPA (Contábil)']
                lpa_proj, div_proj = row['LPA Projetado'], row['Div. Projetado (R$)']
                
                graham = (22.5 * lpa_proj * vpa) ** 0.5 if (lpa_proj > 0 and vpa > 0) else np.nan
                margem_g = ((graham / cotacao) - 1) * 100 if (pd.notna(graham) and cotacao > 0) else np.nan
                bazin = div_proj / yield_desejado if div_proj > 0 else np.nan
                margem_b = ((bazin / cotacao) - 1) * 100 if (pd.notna(bazin) and cotacao > 0) else np.nan
                
                linhas_val.append({"Ativo": t, "Preço Graham": graham, "Margem Graham": margem_g, "Preço Bazin": bazin, "Margem Bazin": margem_b})
                
            st.dataframe(pd.DataFrame(linhas_val), use_container_width=True, hide_index=True, column_config={
                "Preço Graham": st.column_config.NumberColumn("Preço Teto Graham", format="R$ %.2f"),
                "Margem Graham": st.column_config.NumberColumn("Margem Graham", format="%.2f %%"),
                "Preço Bazin": st.column_config.NumberColumn("Preço Teto Bazin", format="R$ %.2f"),
                "Margem Bazin": st.column_config.NumberColumn("Margem Bazin", format="%.2f %%")
            })

        # --- ABA 3: GRÁFICOS INTERATIVOS ---
        with tab3:
            st.markdown("### Comparativo de Performance")
            st.markdown("Selecione os ativos que deseja cruzar com os índices de referência.")
            
            todos_ativos = df_perf_final['Ativo'].tolist()
            # Pré-seleciona os 5 primeiros como demonstração
            ativos_selecionados = st.multiselect("Selecione os Ativos:", todos_ativos, default=todos_ativos[:5])
            
            if ativos_selecionados:
                # Filtra os dados
                df_grafico = df_perf_final[df_perf_final['Ativo'].isin(ativos_selecionados)]
                # Prepara o formato para o st.bar_chart (Index = Eixo X, Colunas = Barras)
                df_grafico = df_grafico.set_index("Ativo")[["Evolução c/ Div", "CDI Acum.", "IPCA Acum."]]
                
                st.bar_chart(df_grafico, use_container_width=True)
            else:
                st.info("Selecione pelo menos um ativo para gerar o gráfico.")
