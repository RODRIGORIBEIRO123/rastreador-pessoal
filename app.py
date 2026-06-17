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
# 1. CONFIGURAÇÃO E MEMÓRIA
# ==========================================
st.set_page_config(page_title="Terminal de Gestão CNPI", layout="wide")
st.title("📊 Terminal de Gestão Profissional")

MAPEAMENTO_TICKERS = {
    "GALG11": "GARE11", "SOMA3": "ALOS3", "ARZZ3": "ALOS3", 
    "VVAR3": "BHIA3", "VIIA3": "BHIA3", "BRML3": "ALSO3", 
    "BBRK11": "BRCR11", "HCTR11": "TRXD11", "TORD11": "TRXD11"
}

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

# 🛡️ MOTOR DE DADOS CONTÁBEIS (FUNDAMENTUS BRASIL)
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
            pl = float(row['P/L'])
            pvp = float(row['P/VP'])
            
            # Engenharia Reversa para achar VPA e LPA exatos
            vpa = c / pvp if pvp > 0 else 0.0
            lpa = c / pl if pl > 0 else 0.0
            
            fundamentos[t] = {'vpa': vpa, 'lpa': lpa}
        return fundamentos
    except:
        return {}

def calcular_macro_acumulado(df_macro, data_inicio, data_fim=None):
    if df_macro is None or df_macro.empty or pd.isna(data_inicio): return 0.0, 0.0
    try:
        filtro = df_macro.loc[data_inicio:data_fim] if data_fim else df_macro.loc[data_inicio:]
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
        
        soma_tempo = 0
        for _, row in group.iterrows():
            if pd.notna(row['Data Média']):
                peso = row['Quantidade'] * row['Preço Médio']
                ts = pd.Timestamp(row['Data Média']).timestamp()
                soma_tempo += (ts * peso)
                
        dt_media = pd.to_datetime(soma_tempo / valor_total, unit='s').date() if (valor_total > 0 and soma_tempo > 0) else pd.Timestamp.now().date()
        linhas.append({"Ativo": ativo, "Quantidade": qtd, "Preço Médio": float(pm), "Data Média": dt_media})
    return pd.DataFrame(linhas)

def ler_arquivo_universal(arquivo_upload):
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
    df_p, df_v = df_perf.fillna(0), df_val.fillna(0)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_p.to_excel(writer, sheet_name='Rentabilidade', index=False)
        df_v.to_excel(writer, sheet_name='Valuation_Bazin', index=False)
        
        ws = writer.sheets['Rentabilidade']
        chart = BarChart()
        chart.type, chart.style = "col", 13
        chart.title, chart.y_axis.title = "Retorno Total vs CDI e IPCA", "Rentabilidade (%)"
        
        dados_grafico = Reference(ws, min_col=13, min_row=1, max_col=15, max_row=len(df_p)+1)
        categorias = Reference(ws, min_col=1, min_row=2, max_row=len(df_p)+1)
        
        chart.add_data(dados_grafico, titles_from_data=True)
        chart.set_categories(categorias)
        chart.height, chart.width = 15, 30
        ws.add_chart(chart, "P2") 
    return output.getvalue()

# ==========================================
# 2. SISTEMA DE UPLOAD
# ==========================================
st.sidebar.header("1. Upload de Dados")
arquivo = st.sidebar.file_uploader("Arquivo B3 ou Backup da Carteira", type=["xlsx", "csv"])

if arquivo and st.session_state.df_base.empty:
    with st.spinner("Analisando e Consolidando Dados..."):
        try:
            df = ler_arquivo_universal(arquivo)
            
            if 'Data Média' in df.columns and 'Ativo' in df.columns:
                df['Data Média'] = pd.to_datetime(df['Data Média'], errors='coerce').dt.date
                st.session_state.df_base = consolidar_carteira(df)
                st.success("✅ Banco de Dados Restaurado e Consolidado!")
                st.rerun()
                
            else:
                df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], dayfirst=True, errors='coerce')
                df['Quantidade'], df['Preço'], df['Valor'] = df['Quantidade'].apply(limpar_numero), df['Preço'].apply(limpar_numero), df['Valor'].apply(limpar_numero)
                df = df.sort_values('Data do Negócio')
                
                posicoes = {}
                for _, row in df.iterrows():
                    if ignorar_ativo(row['Código de Negociação']): continue
                    ticker = str(row['Código de Negociação']).strip().upper()
                    ticker = ticker[:-1] if ticker.endswith('F') else ticker
                    ticker = MAPEAMENTO_TICKERS.get(ticker, ticker)
                    
                    qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio']
                    
                    if ticker not in posicoes: posicoes[ticker] = {'qtd': 0.0, 'valor': 0.0, 'soma_pesos': 0.0}
                        
                    if row['Tipo de Movimentação'] == 'Compra':
                        if posicoes[ticker]['qtd'] == 0: posicoes[ticker]['soma_pesos'] = 0.0
                        posicoes[ticker]['qtd'] += qtd
                        posicoes[ticker]['valor'] += valor
                        posicoes[ticker]['soma_pesos'] += (pd.Timestamp(data).timestamp() * valor)
                        
                    elif row['Tipo de Movimentação'] == 'Venda' and posicoes[ticker]['qtd'] > 0:
                        qtd_venda = min(qtd, posicoes[ticker]['qtd'])
                        pm_atual = posicoes[ticker]['valor'] / posicoes[ticker]['qtd']
                        posicoes[ticker]['qtd'] -= qtd_venda
                        posicoes[ticker]['valor'] -= (qtd_venda * pm_atual)
                        if posicoes[ticker]['qtd'] <= 0.001: 
                            posicoes[ticker]['qtd'], posicoes[ticker]['valor'], posicoes[ticker]['soma_pesos'] = 0.0, 0.0, 0.0

                ativos_limpos = []
                for t, d in posicoes.items():
                    if d['qtd'] > 0 and d['valor'] > 0:
                        data_proporcional = pd.to_datetime(d['soma_pesos'] / d['valor'], unit='s')
                        ativos_limpos.append({
                            "Ativo": t, "Quantidade": float(d['qtd']), 
                            "Preço Médio": float(d['valor'] / d['qtd']), "Data Média": data_proporcional.date()
                        })
                st.session_state.df_base = consolidar_carteira(pd.DataFrame(ativos_limpos))
                st.rerun()
        except Exception as e:
            st.error(f"Erro ao processar arquivo. Detalhe: {e}")

# ==========================================
# 3. INTERFACE DE PARAMETRIZAÇÃO E BACKUP
# ==========================================
if not st.session_state.df_base.empty:
    st.markdown("### 2. Controle do Banco de Dados")
    
    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_a:
        st.error("🗑️ EXCLUIR ATIVO")
        lista_ativos = [""] + sorted(st.session_state.df_base["Ativo"].tolist())
        ticker_del = st.selectbox("Selecione o ativo:", lista_ativos, key="del_box")
        if st.button("Remover", use_container_width=True):
            if ticker_del != "":
                st.session_state.df_base = st.session_state.df_base[st.session_state.df_base["Ativo"] != ticker_del]
                st.rerun()
                
    with col_b:
        st.success("➕ NOVA COMPRA / ATIVO")
        novo_t = st.text_input("Ticker (Ex: BBAS3)")
        c_q, c_p = st.columns(2)
        novo_q = c_q.number_input("Qtd", min_value=1)
        novo_p = c_p.number_input("PM (R$)", min_value=0.01)
        if st.button("Adicionar à Base", use_container_width=True):
            if novo_t != "":
                nova_linha = pd.DataFrame([{"Ativo": novo_t.upper(), "Quantidade": float(novo_q), "Preço Médio": float(novo_p), "Data Média": pd.Timestamp.now().date()}])
                st.session_state.df_base = consolidar_carteira(pd.concat([st.session_state.df_base, nova_linha], ignore_index=True))
                st.rerun()
                
    with col_c:
        st.info("💾 SALVAR ESTADO ATUAL")
        csv_backup = st.session_state.df_base.to_csv(index=False, sep=';', encoding='utf-8-sig')
        st.download_button(label="📥 Baixar Banco de Dados (.csv)", data=csv_backup, file_name="Banco_de_Dados_Carteira.csv", mime="text/csv", use_container_width=True)

    st.write("---")
    df_editado = st.data_editor(
        st.session_state.df_base, use_container_width=True, hide_index=True,
        column_config={
            "Ativo": st.column_config.TextColumn("Ativo (Editável)", disabled=False),
            "Quantidade": st.column_config.NumberColumn(min_value=0.0),
            "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0),
            "Data Média": st.column_config.DateColumn("Data Média Ponderada", format="DD/MM/YYYY")
        }
    )

    if st.button("🚀 Processar Conexão com o Mercado", type="primary"):
        st.session_state.df_base = consolidar_carteira(df_editado) 
        df_macro = carregar_macro()
        fundamentos_br = obter_fundamentos_brasil() # <--- O Motor Secundário Atuando
        progresso = st.progress(0)
        total = len(st.session_state.df_base)
        data_12m = pd.Timestamp.now() - pd.DateOffset(years=1)
        
        dados_mercado = {}
        linhas_simul_iniciais = []

        for i, row in st.session_state.df_base.iterrows():
            ticker = str(row['Ativo']).strip().upper()
            data_compra = pd.to_datetime(row['Data Média']) if pd.notna(row['Data Média']) else pd.Timestamp.now()
            
            preco_atual = float(row['Preço Médio'])
            divs_total, divs_12m, lpa, vpa = 0.0, 0.0, 0.0, 0.0
            
            try:
                acao = yf.Ticker(f"{ticker}.SA")
                try: 
                    hist = acao.history(period="1d")
                    if not hist.empty: preco_atual = float(hist['Close'].iloc[-1])
                except: pass
                try: 
                    divs = acao.dividends
                    divs_total = float(divs[divs.index.tz_localize(None) >= data_compra].sum() * row['Quantidade'])
                    divs_12m = float(divs[divs.index.tz_localize(None) >= data_12m].sum())
                except: pass
            except: pass

            # ==========================================
            # PREENCHIMENTO INFALÍVEL DE FUNDAMENTOS
            # ==========================================
            if ticker in fundamentos_br:
                vpa = fundamentos_br[ticker]['vpa']
                lpa = fundamentos_br[ticker]['lpa']
            else:
                try: # Se for FII ou Ativo Estrangeiro, tenta o Yahoo como Fallback
                    info = acao.info
                    if isinstance(info, dict):
                        lpa = float(info.get('trailingEps') or info.get('forwardEps') or 0.0)
                        vpa = float(info.get('bookValue') or 0.0)
                        p_vp = info.get('priceToBook')
                        if vpa == 0.0 and p_vp and preco_atual > 0:
                            try: vpa = preco_atual / float(p_vp)
                            except: pass
                except: pass

            cdi, ipca = calcular_macro_acumulado(df_macro, data_compra)
            meses_investido = calcular_meses(data_compra)
            
            dados_mercado[ticker] = {
                "Qtd": float(row['Quantidade']), "PM": float(row['Preço Médio']), "Data": data_compra,
                "Preço Atual": preco_atual, "Div_Total": divs_total, "CDI": cdi, "IPCA": ipca, "Meses": meses_investido
            }

            linhas_simul_iniciais.append({
                "Ativo": ticker, "Cotação Atual": preco_atual, "VPA (Contábil)": vpa,
                "LPA Projetado": lpa, "Div. Projetado (R$)": divs_12m
            })
            progresso.progress((i + 1) / total)
            
        st.session_state.dados_mercado = dados_mercado
        st.session_state.df_simul = pd.DataFrame(linhas_simul_iniciais)
        st.success("Análise Matemática Concluída e Posições Fundidas!")

    # ==========================================
    # 4. PAINEL DE RELATÓRIOS (4 ABAS)
    # ==========================================
    if st.session_state.dados_mercado:
        linhas_perf = []
        for t, dm in st.session_state.dados_mercado.items():
            investido = dm['Qtd'] * dm['PM']
            saldo = dm['Qtd'] * dm['Preço Atual']
            resultado = saldo - investido
            
            var_s_div = ((saldo / investido) - 1) * 100 if investido > 0 else np.nan
            var_c_div = (((saldo + dm['Div_Total']) / investido) - 1) * 100 if investido > 0 else np.nan
            yoc = (dm['Div_Total'] / investido) * 100 if investido > 0 else np.nan
            
            linhas_perf.append({
                "Ativo": t, 
                "Qtd": int(dm['Qtd']), 
                "Preço Médio": dm['PM'], 
                "Preço Atual": dm['Preço Atual'],
                "Total Investido": investido,
                "Saldo Atual": saldo,
                "Resultado (R$)": resultado,
                "Data Média": dm['Data'].strftime('%d/%m/%Y'),
                "Meses (Média)": int(dm['Meses']),
                "Total Div. (R$)": dm['Div_Total'], 
                "DY on Cost": yoc,
                "Evolução s/ Div": var_s_div, 
                "Evolução c/ Div": var_c_div,
                "IPCA Acum.": dm['IPCA'], 
                "CDI Acum.": dm['CDI']
            })
        df_perf_final = pd.DataFrame(linhas_perf)

        st.write("---")
        st.download_button(
            label="📥 Exportar Relatório Executivo (Excel com Gráficos)",
            data=gerar_excel_premium(df_perf_final, st.session_state.df_simul),
            file_name="Relatorio_CNPI_Carteira.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        tab1, tab2, tab3, tab4 = st.tabs([
            "📈 Rentabilidade e YOC", 
            "💰 Método Bazin (Renda)", 
            "🏢 Método Graham (Valor)", 
            "📊 Análise Gráfica (Plotly)"
        ])
        
        with tab1:
            st.dataframe(df_perf_final, use_container_width=True, hide_index=True, column_config={
                "Preço Médio": st.column_config.NumberColumn(format="R$ %.2f"),
                "Preço Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Total Investido": st.column_config.NumberColumn(format="R$ %.2f"),
                "Saldo Atual": st.column_config.NumberColumn(format="R$ %.2f"),
                "Resultado (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                "Total Div. (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                "DY on Cost": st.column_config.NumberColumn(format="%.2f %%"),
                "Evolução s/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "Evolução c/ Div": st.column_config.NumberColumn(format="%.2f %%"),
                "IPCA Acum.": st.column_config.NumberColumn(format="%.2f %%"),
                "CDI Acum.": st.column_config.NumberColumn(format="%.2f %%")
            })

        with tab2:
            st.markdown("### Método Décio Bazin")
            yield_desejado = st.number_input("Taxa de Risco Exigida (%):", value=6.0, min_value=0.1, step=0.5) / 100.0
            
            df_bazin_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "Div. Projetado (R$)"]].copy()
            df_bazin_editado = st.data_editor(df_bazin_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"],
                column_config={"Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"), "Div. Projetado (R$)": st.column_config.NumberColumn("Div. Projetado (R$)", format="R$ %.2f")},
                key="edit_bazin")
            
            st.session_state.df_simul["Div. Projetado (R$)"] = df_bazin_editado["Div. Projetado (R$)"]
            
            linhas_bazin = []
            for _, row in df_bazin_editado.iterrows():
                t = str(row['Ativo'])
                try: cotacao = float(row['Cotação Atual'])
                except: cotacao = 0.0
                try: div_proj = float(row['Div. Projetado (R$)'])
                except: div_proj = 0.0
                
                bazin = (div_proj / yield_desejado) if (div_proj > 0 and yield_desejado > 0) else 0.0
                margem_b = (((bazin / cotacao) - 1) * 100) if (bazin > 0 and cotacao > 0) else 0.0
                linhas_bazin.append({"Ativo": t, "Cotação Atual": cotacao, "Preço Teto (Bazin)": bazin, "Margem de Segurança": margem_b})
                
            st.dataframe(pd.DataFrame(linhas_bazin), use_container_width=True, hide_index=True, column_config={
                "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"), "Preço Teto (Bazin)": st.column_config.NumberColumn(format="R$ %.2f"), "Margem de Segurança": st.column_config.NumberColumn(format="%.2f %%")})

        with tab3:
            st.markdown("### Método Benjamin Graham")
            df_graham_view = st.session_state.df_simul[["Ativo", "Cotação Atual", "VPA (Contábil)", "LPA Projetado"]].copy()
            df_graham_editado = st.data_editor(df_graham_view, use_container_width=True, hide_index=True, disabled=["Ativo", "Cotação Atual"],
                column_config={"Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"), "VPA (Contábil)": st.column_config.NumberColumn("VPA (Editar)", format="R$ %.2f"), "LPA Projetado": st.column_config.NumberColumn("LPA Projetado (Editar)", format="R$ %.2f")},
                key="edit_graham")
            
            st.session_state.df_simul["VPA (Contábil)"] = df_graham_editado["VPA (Contábil)"]
            st.session_state.df_simul["LPA Projetado"] = df_graham_editado["LPA Projetado"]
            
            linhas_graham = []
            for _, row in df_graham_editado.iterrows():
                t = str(row['Ativo'])
                try: cotacao = float(row['Cotação Atual'])
                except: cotacao = 0.0
                try: vpa = float(row['VPA (Contábil)'])
                except: vpa = 0.0
                try: lpa_proj = float(row['LPA Projetado'])
                except: lpa_proj = 0.0
                
                graham = (22.5 * lpa_proj * vpa) ** 0.5 if (lpa_proj > 0 and vpa > 0) else 0.0
                margem_g = (((graham / cotacao) - 1) * 100) if (graham > 0 and cotacao > 0) else 0.0
                linhas_graham.append({"Ativo": t, "Cotação Atual": cotacao, "Preço Justo (Graham)": graham, "Margem de Segurança": margem_g})
                
            st.dataframe(pd.DataFrame(linhas_graham), use_container_width=True, hide_index=True, column_config={
                "Cotação Atual": st.column_config.NumberColumn(format="R$ %.2f"), "Preço Justo (Graham)": st.column_config.NumberColumn(format="R$ %.2f"), "Margem de Segurança": st.column_config.NumberColumn(format="%.2f %%")})

        with tab4:
            st.markdown("### 1. Performance Global (Individualizada por Ativo)")
            todos_ativos = df_perf_final['Ativo'].tolist()
            
            c_sel, c_ind = st.columns([2, 1])
            with c_sel: ativos_selecionados = st.multiselect("Selecione os ativos:", todos_ativos, default=todos_ativos[:6], key="ms_global")
            with c_ind: ind_selecionados = st.multiselect("Indicadores:", ["Evolução c/ Div", "CDI Acum.", "IPCA Acum."], default=["Evolução c/ Div", "CDI Acum.", "IPCA Acum."], key="ind_global")
            
            if ativos_selecionados and ind_selecionados:
                df_grafico = df_perf_final[df_perf_final['Ativo'].isin(ativos_selecionados)].copy()
                df_grafico['Período'] = df_grafico['Data Média'].astype(str) + " até Hoje"
                
                df_melt = df_grafico.melt(id_vars=["Ativo", "Período"], value_vars=ind_selecionados, var_name="Indicador", value_name="Rentabilidade")
                
                titulo_graf1 = f"Performance Desde: {df_grafico.iloc[0]['Data Média']} até Hoje" if len(ativos_selecionados) == 1 else "Performance Baseada nas Datas Médias de Cada Ativo"
                fig1 = px.bar(df_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group", text="Rentabilidade", hover_data=["Período"], title=titulo_graf1)
                
                fig1.update_traces(texttemplate='%{text:.2f}%', textposition='outside', hovertemplate='<b>%{x}</b> (%{data.name})<br>Período: %{customdata[0]}<br>Rentabilidade: %{y:.2f}%<extra></extra>')
                fig1.update_layout(yaxis_ticksuffix=" %", margin=dict(t=40))
                st.plotly_chart(fig1, use_container_width=True)
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
                            fig_custom = px.bar(df_custom_melt, x="Ativo", y="Rentabilidade", color="Indicador", barmode="group", text="Rentabilidade", hover_data=["Período"], title=titulo_graf2)
                            
                            fig_custom.update_traces(texttemplate='%{text:.2f}%', textposition='outside', hovertemplate='<b>%{x}</b> (%{data.name})<br>Período: %{customdata[0]}<br>Rentabilidade: %{y:.2f}%<extra></extra>')
                            fig_custom.update_layout(yaxis_ticksuffix=" %", margin=dict(t=40))
                            st.plotly_chart(fig_custom, use_container_width=True)
                        else: st.error("Sem histórico para o período.")
