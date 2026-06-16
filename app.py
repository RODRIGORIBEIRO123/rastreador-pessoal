import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from bcb import sgs
import re
import io
from openpyxl.chart import BarChart, Reference
import plotly.express as px

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

def calcular_meses(data_inicio):
    if pd.isna(data_inicio): return 0
    hoje = pd.Timestamp.now()
    return int((hoje.year - data_inicio.year) * 12 + (hoje.month - data_inicio.month))

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

def gerar_excel_premium(df_perf, df_val):
    output = io.BytesIO()
    df_p = df_perf.fillna(0)
    df_v = df_val.fillna(0)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_p.to_excel(writer, sheet_name='Rentabilidade', index=False)
        df_v.to_excel(writer, sheet_name='Valuation_Bazin', index=False)
        ws = writer.sheets['Rentabilidade']
        chart = BarChart()
        chart.type = "col"
        chart.style = 13
        chart.title = "Retorno Total vs CDI e IPCA"
        chart.y_axis.title = "Rentabilidade (%)"
        dados_grafico = Reference(ws, min_col=9, min_row=1, max_col=11, max_row=len(df_p)+1)
        categorias = Reference(ws, min_col=1, min_row=2, max_row=len(df_p)+1)
        chart.add_data(dados_grafico, titles_from_data=True)
        chart.set_categories(categorias)
        chart.height = 15
        chart.width = 30
        ws.add_chart(chart, "M2")
    return output.getvalue()

# ==========================================
# 2. UPLOAD E LEITURA DA B3
# ==========================================
st.sidebar.header("1. Upload da Carteira")
arquivo = st.sidebar.file_uploader("Arquivo Base", type=["xlsx", "csv"])

if arquivo and st.session_state.df_base.empty:
    with st.spinner("Purificando base de dados..."):
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
                    data_aporte = d['primeiro_aporte'] if pd.notna(d['primeiro_aporte']) else pd.Timestamp.now()
                    ativos_limpos.append({
                        "Ativo": t, "Quantidade": float(d['qtd']), 
                        "Preço Médio": float(d['valor'] / d['qtd']), 
                        "Data 1º Aporte": data_aporte.date()
                    })
            st.session_state.df_base = pd.DataFrame(ativos_limpos)
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao ler o arquivo. Detalhe técnico: {e}")

# ==========================================
# 3. INTERFACE DE PARAMETRIZAÇÃO
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Parametrização da Carteira")
    st.markdown("Ajuste os preços médios e a **Data do 1º Aporte**. O sistema buscará o CDI e IPCA acumulados exatos a partir da data que você definir.")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        lista_ativos = [""] + sorted(st.session_state.df_base["Ativo"].tolist())
        ticker_del = st.selectbox("Selecione o ativo para remover:", lista_ativos)
        if st.button("🗑️ Remover Ativo"):
            if ticker_del != "":
                st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != ticker_del]
                st.rerun()
    with col2:
        c1, c2, c3 = st.columns(3)
        novo_t = c1.text_input("Ticker (Ex: BBAS3)")
        novo_q = c2.number_input("Qtd", min_value=1)
        novo_p = c3.number_input("PM (R$)", min_value=0.01)
        if st.button("➕ Adicionar Ativo"):
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
            "Data 1º Aporte": st.column_config.DateColumn("Data 1º Aporte (Ajuste Aqui)", format="DD/MM/YYYY")
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
                # O CDI agora obedece estritamente a Data que o usuário editou
                data_compra = pd.to_datetime(row['Data 1º Aporte']) if pd.notna(row['Data 1º Aporte']) else pd.Timestamp.now()
                
                acao = yf.Ticker(f"{ticker}.SA")
                hist = acao.history(period="1d")
                preco_atual = float(hist['Close'].iloc[-1]) if not hist.empty else float(row['Preço Médio'])
                
                divs = acao.dividends
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

            # Prepara a base do simulador com um Yield exigido padrão de 6% (Editável pelo usuário depois)
            linhas_simul_iniciais.append({
                "Ativo": ticker, "Cotação Atual": preco_atual, "VPA (Contábil)": vpa,
                "LPA Projetado": lpa, "Div. Projetado (R$)": divs_12m, "Yield Desejado (%)": 6.0
            })
            progresso.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dados_mercado
        st.session_state.df_simul = pd.DataFrame(linhas_simul_iniciais)
        st.success("Dados de Mercado Sincronizados!")

    # ==========================================
    # 4. PAINEL DE RELATÓRIOS (4 ABAS)
    # ==========================================
    if st.session_state.dados_mercado:
        linhas_perf = []
        for t, dm in st.session_state.dados_mercado.items():
            investido = dm['Qtd'] * dm['PM']
            saldo = dm['Qtd'] * dm['Preço Atual']
            var_s_div = ((saldo / investido) - 1) * 100 if investido > 0 else np.nan
            var_c_div = (((saldo + dm['Div_Total']) / investido) - 1) * 100 if investido > 0 else np.nan
            yoc = (dm['Div_Total'] / investido) * 100 if investido > 0 else np.nan
            
            linhas_perf.append({
                "Ativo": t, "Qtd": int(dm['Qtd']), "Preço Médio": dm['PM'], "Preço Atual": dm['Preço Atual'],
                "Meses": int(dm['Meses']),
                "Total Div. (R$)": dm['Div_Total'], "DY on Cost": yoc,
                "Evolução s/ Div": var_s_div, "Evolução c/ Div": var_c_div,
                "IPCA Acum.": dm['IPCA'], "CDI Acum.": dm['CDI']
            })
        df_perf_final = pd.DataFrame(linhas_perf)

        st.write("---")
        st.download_button(
            label="📥 Baixar Relatório em Excel",
            data=gerar_excel_premium(df_perf_final, st.session_state.df_simul),
            file_name="Analise_CNPI_Carteira.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        tab1, tab2, tab3, tab4 = st.tabs([
            "📈 Rentabilidade e YOC", 
            "💰 Método Bazin (Renda)", 
            "🏢 Método Graham (Valor)", 
            "📊 Gráficos Interativos"
        ])
        
        # --- ABA 1: RENTABILIDADE ---
        with tab1:
            st.dataframe(df_perf_final, use_container_width=True, hide_index=True, column_config={
                "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f"),
                "Preço Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Total Div. (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                "DY on Cost": st.column_config.NumberColumn(format="%.2f %%"),
                "Evolução s/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "Evolução c/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "IPCA Acum.": st.column_config.NumberColumn(format="%.2f %%"),
                "CDI Acum.": st.column_config.NumberColumn(format="%.2f %%")
            })

        # --- ABA 2: BAZIN ---
        with tab2:
            st.markdown("### Método Décio Bazin (Foco em Renda Passiva)")
            st.markdown("Edite o **Dividendo Projetado (R$)** ou o **Yield Desejado (%)** por ativo. O Preço Teto calculará automaticamente.")
            
            df_bazin_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)", "Yield Desejado (%)"]].copy()
            df_bazin_editado = st.data_editor(
                df_bazin_view, use_container_width=True, hide_index=True,
                disabled=["Ativo", "Cotação Atual"],
                column_config={
                    "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Div. Projetado (R$)": st.column_config.NumberColumn("Div. Projetado (R$)", format="R$ %.2f"),
                    "Yield Desejado (%)": st.column_config.NumberColumn("Yield Exigido (%)", format="%.2f %%")
                }
            )
            
            linhas_bazin = []
            for _, row in df_bazin_editado.iterrows():
                t, cotacao = row['Ativo'], row['Cotação Atual']
                div_proj, yield_req = row['Div. Projetado (R$)'], (row['Yield Desejado (%)'] / 100.0)
                
                bazin = div_proj / yield_req if (div_proj > 0 and yield_req > 0) else np.nan
                margem_b = ((bazin / cotacao) - 1) * 100 if (pd.notna(bazin) and cotacao > 0) else np.nan
                linhas_bazin.append({"Ativo": t, "Cotação Atual": cotacao, "Preço Teto (Bazin)": bazin, "Margem de Segurança": margem_b})
                
            st.dataframe(pd.DataFrame(linhas_bazin), use_container_width=True, hide_index=True, column_config={
                "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Preço Teto (Bazin)": st.column_config.NumberColumn(format="R$ %.2f"),
                "Margem de Segurança": st.column_config.NumberColumn(format="%.2f %%")
            })

        # --- ABA 3: GRAHAM ---
        with tab3:
            st.markdown("### Método Benjamin Graham (Valor Intrínseco)")
            st.warning("⚠️ O método de Graham não se aplica a FIIs ou a empresas com prejuízo contábil (VPA ou LPA negativos).")
            
            df_graham_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "VPA (Contábil)", "LPA Projetado"]].copy()
            df_graham_editado = st.data_editor(
                df_graham_view, use_container_width=True, hide_index=True,
                disabled=["Ativo", "Cotação Atual", "VPA (Contábil)"],
                column_config={
                    "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                    "VPA (Contábil)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "LPA Projetado": st.column_config.NumberColumn("LPA Projetado (Editar)", format="R$ %.2f")
                }
            )
            
            linhas_graham = []
            for _, row in df_graham_editado.iterrows():
                t, cotacao = row['Ativo'], row['Cotação Atual']
                vpa, lpa_proj = row['VPA (Contábil)'], row['LPA Projetado']
                
                graham = (22.5 * lpa_proj * vpa) ** 0.5 if (lpa_proj > 0 and vpa > 0) else np.nan
                margem_g = ((graham / cotacao) - 1) * 100 if (pd.notna(graham) and cotacao > 0) else np.nan
                linhas_graham.append({"Ativo": t, "Cotação Atual": cotacao, "Preço Justo (Graham)": graham, "Margem de Segurança": margem_g})
                
            st.dataframe(pd.DataFrame(linhas_graham), use_container_width=True, hide_index=True, column_config={
                "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Preço Justo (Graham)": st.column_config.NumberColumn(format="R$ %.2f"),
                "Margem de Segurança": st.column_config.NumberColumn(format="%.2f %%")
            })

        # --- ABA 4: GRÁFICOS (PLOTLY Engine) ---
        with tab4:
            st.markdown("### Análise Comparativa Visual")
            
            todos_ativos = df_perf_final['Ativo'].tolist()
            ativos_selecionados = st.multiselect("Selecione os ativos para exibir nos gráficos:", todos_ativos, default=todos_ativos[:6])
            
            if ativos_selecionados:
                df_grafico = df_perf_final[df_perf_final['Ativo'].isin(ativos_selecionados)]
                
                # Gráfico 1: Rentabilidade com Labels e Eixo %
                st.markdown("#### Retorno Total vs CDI vs IPCA")
                df_melt = df_grafico.melt(id_vars=["Ativo"], 
                                          value_vars=["Evolução c/ Div", "CDI Acum.", "IPCA Acum."],
                                          var_name="Indicador", value_name="Rentabilidade")
                
                fig1 = px.bar(df_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group", text="Rentabilidade")
                fig1.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                fig1.update_layout(yaxis_ticksuffix=" %", uniformtext_minsize=8, uniformtext_mode='hide', margin=dict(t=30))
                st.plotly_chart(fig1, use_container_width=True)
                
                # Gráfico 2: DY On Cost
                st.markdown("#### Máquina de Geração de Caixa (Yield on Cost)")
                fig2 = px.bar(df_grafico, x="Ativo", y="DY on Cost", text="DY on Cost")
                fig2.update_traces(marker_color='#2ecc71', texttemplate='%{text:.2f}%', textposition='outside')
                fig2.update_layout(yaxis_ticksuffix=" %", margin=dict(t=30))
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("Selecione os ativos no menu acima.")
