import streamlit as st
import pandas as pd
import yfinance as yf
from bcb import sgs
import re

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Terminal de Gestão | CNPI", layout="wide")
st.title("📊 Terminal de Gestão Híbrido")
st.markdown("Ajuste a sua carteira abaixo. **Clique nos cabeçalhos das tabelas finais para ordenar os valores.**")

# --- FUNÇÕES ÚTEIS ---
def eh_opcao_ou_futuro(ticker):
    if pd.isna(ticker): return True
    ticker = str(ticker).strip().upper()
    if ticker.startswith(('WIN', 'WDO', 'IND', 'DOL')): return True
    t = ticker[:-1] if ticker.endswith('F') else ticker
    if re.match(r'^[A-Z]{4}[A-Z]\d+', t) and not t.endswith(('11', '34', '39')):
        if len(t) > 6 or len(t) == 6: return True
    return False

@st.cache_data(ttl=86400)
def obter_dados_macro():
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'], macro['IPCA'] = macro['CDI'] / 100, macro['IPCA'] / 100
        return macro
    except: return None

def calcular_macro_acumulado(df_macro, data_inicio):
    if df_macro is None or df_macro.empty: return 0.0, 0.0
    try:
        filtro = df_macro.loc[data_inicio:]
        return ((1 + filtro['CDI'].dropna()).prod() - 1) * 100, ((1 + filtro['IPCA'].dropna()).prod() - 1) * 100
    except: return 0.0, 0.0

# --- PASSO 1: UPLOAD E LEITURA ---
arquivo = st.file_uploader("1. Upload da planilha 'Negociação' da B3 (.xlsx ou .csv)", type=["xlsx", "csv"])

if arquivo:
    with st.spinner("Processando histórico da B3..."):
        df = pd.read_csv(arquivo, sep=';', encoding='latin1') if arquivo.name.endswith('.csv') else pd.read_excel(arquivo)
        df.columns = df.columns.str.strip()
        df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], format='%d/%m/%Y', errors='coerce')
        df['Preço'] = df['Preço'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        df['Valor'] = df['Valor'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        df = df.sort_values('Data do Negócio')
        
        posicoes = {}
        for _, row in df.iterrows():
            ticker_orig = str(row['Código de Negociação']).strip()
            if eh_opcao_ou_futuro(ticker_orig): continue
            ticker = ticker_orig[:-1] if ticker_orig.endswith('F') else ticker_orig
            qtd, valor, data = row['Quantidade'], row['Valor'], row['Data do Negócio']
            
            if ticker not in posicoes:
                posicoes[ticker] = {'quantidade': 0, 'valor_investido': 0.0, 'primeira_compra': data}
                
            if row['Tipo de Movimentação'] == 'Compra':
                posicoes[ticker]['quantidade'] += qtd
                posicoes[ticker]['valor_investido'] += valor
                if pd.notna(data) and data < posicoes[ticker]['primeira_compra']: 
                    posicoes[ticker]['primeira_compra'] = data
            elif row['Tipo de Movimentação'] == 'Venda' and posicoes[ticker]['quantidade'] > 0:
                qtd_venda = min(qtd, posicoes[ticker]['quantidade'])
                pm_atual = posicoes[ticker]['valor_investido'] / posicoes[ticker]['quantidade']
                posicoes[ticker]['quantidade'] -= qtd_venda
                posicoes[ticker]['valor_investido'] -= (qtd_venda * pm_atual)
                if posicoes[ticker]['quantidade'] <= 0.001: posicoes[ticker]['quantidade'] = 0

        carteira_ativa = {k: v for k, v in posicoes.items() if v['quantidade'] > 0}
        
    dados_edicao = []
    for ticker, dados in sorted(carteira_ativa.items(), key=lambda x: x[1]['valor_investido'], reverse=True):
        pm_estimado = dados['valor_investido'] / dados['quantidade'] if dados['quantidade'] > 0 else 0
        data_pura = dados['primeira_compra'] if pd.notna(dados['primeira_compra']) else pd.Timestamp.now()
        dados_edicao.append({
            "Ativo": ticker, "Quantidade": int(dados['quantidade']),
            "Preço Médio": float(round(pm_estimado, 2)), "Data 1º Aporte": data_pura
        })
        
    df_edicao = pd.DataFrame(dados_edicao)
    if not df_edicao.empty: df_edicao['Data 1º Aporte'] = pd.to_datetime(df_edicao['Data 1º Aporte']).dt.date

    st.write("---")
    st.subheader("2. Edição Livre da Carteira")
    
    df_editado = st.data_editor(
        df_edicao, use_container_width=True, hide_index=False, num_rows="dynamic",
        column_config={
            "Data 1º Aporte": st.column_config.DateColumn("Data 1º Aporte", format="DD/MM/YYYY"),
            "Preço Médio": st.column_config.NumberColumn("Preço Médio", format="R$ %.2f", min_value=0.0)
        }
    )

    # --- PASSO 2: CÁLCULOS E RELATÓRIOS ---
    if st.button("🚀 Gerar Valuation e Retorno Total", type="primary"):
        df_macro = obter_dados_macro()
        progress_bar = st.progress(0)
        dados_perf, dados_val = [], []
        total_ativos = len(df_editado)
        data_12m_atras = pd.Timestamp.now() - pd.DateOffset(years=1)
        
        for i, row in df_editado.iterrows():
            ticker = str(row.get("Ativo", "")).strip().upper()
            if not ticker or ticker == 'NAN' or ticker == 'NONE': continue
            qtd_real = float(row.get("Quantidade", 0))
            if qtd_real <= 0 or pd.isna(qtd_real): continue
            
            pm_real = float(row.get("Preço Médio", 0)) if not pd.isna(row.get("Preço Médio")) else 0.0
            valor_investido_real = qtd_real * pm_real
            
            try: data_compra = pd.to_datetime(row.get("Data 1º Aporte"))
            except: data_compra = pd.Timestamp.now()
            if pd.isna(data_compra): data_compra = pd.Timestamp.now()
            
            try:
                acao = yf.Ticker(f"{ticker}.SA")
                hist = acao.history(period="1d")
                preco_atual = hist['Close'].iloc[-1] if not hist.empty else pm_real
                
                divs = acao.dividends
                total_dividendos = divs[divs.index.tz_localize(None) >= data_compra].sum() * qtd_real
                divs_12m = divs[divs.index.tz_localize(None) >= data_12m_atras].sum()

                info = acao.info
                lpa, vpa = info.get('trailingEps', 0), info.get('bookValue', 0)
                lpa, vpa = lpa if lpa is not None else 0, vpa if vpa is not None else 0
            except:
                preco_atual, total_dividendos, divs_12m, lpa, vpa = pm_real, 0.0, 0.0, 0, 0

            # MATEMÁTICA PURA (FLOAT)
            valor_atual = preco_atual * qtd_real
            var_cota = ((valor_atual / valor_investido_real) - 1) * 100 if valor_investido_real > 0 else 0
            var_total = (((valor_atual + total_dividendos) / valor_investido_real) - 1) * 100 if valor_investido_real > 0 else 0
            cdi_acum, ipca_acum = calcular_macro_acumulado(df_macro, data_compra)

            dados_perf.append({
                "Ativo": ticker, "Qtd": int(qtd_real),
                "PM Real": float(pm_real), "Cotação": float(preco_atual),
                "Investido": float(valor_investido_real), "Saldo": float(valor_atual),
                "Var. Cota": float(var_cota), "Retorno Total": float(var_total),
                "IPCA (Período)": float(ipca_acum), "CDI (Período)": float(cdi_acum)
            })

            graham = (22.5 * lpa * vpa) ** 0.5 if lpa > 0 and vpa > 0 else 0
            margem_g = ((graham / preco_atual) - 1) * 100 if graham > 0 and preco_atual > 0 else 0
            bazin = divs_12m / 0.06 if divs_12m > 0 else 0
            margem_b = ((bazin / preco_atual) - 1) * 100 if bazin > 0 and preco_atual > 0 else 0

            # Identificadores Visuais
            status_g = "🟢" if margem_g > 0 else ("🔴" if graham > 0 else "-")
            status_b = "🟢" if margem_b > 0 else ("🔴" if bazin > 0 else "-")

            dados_val.append({
                "Ativo": ticker, "Cotação": float(preco_atual),
                "LPA": float(lpa), "VPA": float(vpa), "Div. 12m": float(divs_12m),
                "Preço Graham": float(graham), "Status Graham": status_g, "Margem Graham": float(margem_g),
                "Preço Bazin": float(bazin), "Status Bazin": status_b, "Margem Bazin": float(margem_b)
            })
            progress_bar.progress((i + 1) / total_ativos)
        
        st.write("---")
        tab1, tab2 = st.tabs(["📈 Rentabilidade e Retorno Total", "🔎 Valuation (Graham & Bazin)"])
        
        # MÁSCARAS DE COLUNA STREAMLIT (Mantém a ordenação viva)
        col_config_perf = {
            "PM Real": st.column_config.NumberColumn("PM Real", format="R$ %.2f"),
            "Cotação": st.column_config.NumberColumn("Cotação", format="R$ %.2f"),
            "Investido": st.column_config.NumberColumn("Investido", format="R$ %.2f"),
            "Saldo": st.column_config.NumberColumn("Saldo", format="R$ %.2f"),
            "Var. Cota": st.column_config.NumberColumn("Var. Cota", format="%.2f %%"),
            "Retorno Total": st.column_config.NumberColumn("Retorno Total", format="%.2f %%"),
            "IPCA (Período)": st.column_config.NumberColumn("IPCA", format="%.2f %%"),
            "CDI (Período)": st.column_config.NumberColumn("CDI", format="%.2f %%")
        }
        
        col_config_val = {
            "Cotação": st.column_config.NumberColumn("Cotação", format="R$ %.2f"),
            "LPA": st.column_config.NumberColumn("LPA", format="R$ %.2f"),
            "VPA": st.column_config.NumberColumn("VPA", format="R$ %.2f"),
            "Div. 12m": st.column_config.NumberColumn("Div. 12m", format="R$ %.2f"),
            "Preço Graham": st.column_config.NumberColumn("Preço Graham", format="R$ %.2f"),
            "Margem Graham": st.column_config.NumberColumn("Margem Graham", format="%.2f %%"),
            "Preço Bazin": st.column_config.NumberColumn("Preço Bazin", format="R$ %.2f"),
            "Margem Bazin": st.column_config.NumberColumn("Margem Bazin", format="%.2f %%")
        }

        with tab1: 
            st.dataframe(pd.DataFrame(dados_perf), use_container_width=True, hide_index=True, column_config=col_config_perf)
        with tab2: 
            st.dataframe(pd.DataFrame(dados_val), use_container_width=True, hide_index=True, column_config=col_config_val)

else:
    st.info("Aguardando o upload do ficheiro da B3 para iniciar.")
