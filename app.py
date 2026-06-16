import streamlit as st
import pandas as pd
import yfinance as yf
from bcb import sgs
import re

# --- CONFIGURAÇÃO E FUNÇÕES BÁSICAS ---
st.set_page_config(page_title="Terminal de Gestão | Total Return", layout="wide")
st.title("📊 Terminal de Gestão (Correção de Splits Ativada)")

def formatar_brl(valor):
    if valor == 0 or pd.isna(valor): return "N/A"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_pct(valor):
    if pd.isna(valor): return "0,00%"
    return f"{valor:,.2f}%".replace(".", ",")

def eh_opcao_ou_futuro(ticker):
    if ticker.startswith(('WIN', 'WDO', 'IND', 'DOL')): return True
    t = ticker[:-1] if str(ticker).endswith('F') else str(ticker)
    if re.match(r'^[A-Z]{4}[A-Z]\d+', t) and not t.endswith(('11', '34', '39')):
        if len(t) > 6 or len(t) == 6: return True
    return False

@st.cache_data(ttl=86400)
def obter_dados_macro():
    try:
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'], macro['IPCA'] = macro['CDI'] / 100, macro['IPCA'] / 100
        return macro
    except Exception:
        return None

def calcular_macro_acumulado(df_macro, data_inicio):
    if df_macro is None or df_macro.empty: return 0.0, 0.0
    filtro = df_macro.loc[data_inicio:]
    return ((1 + filtro['CDI'].dropna()).prod() - 1) * 100, ((1 + filtro['IPCA'].dropna()).prod() - 1) * 100

# --- LÓGICA DE PROCESSAMENTO ---
arquivo = st.file_uploader("Upload da planilha de Negociação da B3 (.xlsx ou .csv)", type=["xlsx", "csv"])

if arquivo:
    with st.spinner("Lendo ordens e corrigindo eventos corporativos (Desdobramentos/Bonificações)..."):
        df = pd.read_csv(arquivo, sep=';', encoding='latin1') if arquivo.name.endswith('.csv') else pd.read_excel(arquivo)
        df.columns = df.columns.str.strip()
        df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], format='%d/%m/%Y', errors='coerce')
        df['Preço'] = df['Preço'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        df['Valor'] = df['Valor'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        df = df.sort_values('Data do Negócio')
        
        # 1. Agrupamento Base (Exatamente o que você pagou)
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
                if data < posicoes[ticker]['primeira_compra']: posicoes[ticker]['primeira_compra'] = data
            elif row['Tipo de Movimentação'] == 'Venda' and posicoes[ticker]['quantidade'] > 0:
                qtd_venda = min(qtd, posicoes[ticker]['quantidade'])
                pm_atual = posicoes[ticker]['valor_investido'] / posicoes[ticker]['quantidade']
                posicoes[ticker]['quantidade'] -= qtd_venda
                posicoes[ticker]['valor_investido'] -= (qtd_venda * pm_atual)
                if posicoes[ticker]['quantidade'] <= 0.001: 
                    posicoes[ticker]['quantidade'] = 0
                    posicoes[ticker]['valor_investido'] = 0.0

        carteira_ativa = {k: v for k, v in posicoes.items() if v['quantidade'] > 0}
        
    df_macro = obter_dados_macro()
    progress_bar = st.progress(0)
    dados_performance, dados_valuation = [], []
    ativos_lista = sorted(carteira_ativa.items(), key=lambda x: x[1]['valor_investido'], reverse=True)
    total_ativos = len(ativos_lista)
    data_12m_atras = pd.Timestamp.now() - pd.DateOffset(years=1)
    
    # 2. Correção Mágica: Ajuste Quantitativo por Splits
    for i, (ticker, dados) in enumerate(ativos_lista):
        try:
            acao = yf.Ticker(f"{ticker}.SA")
            hist = acao.history(start=dados['primeira_compra'].strftime('%Y-%m-%d'))
            preco_atual = hist['Close'].iloc[-1] if not hist.empty else (dados['valor_investido'] / dados['quantidade'])
            
            # --- O SEGREDO DO AJUSTE DE BBAS3 ---
            splits = acao.splits
            # Pega todos os splits que ocorreram DEPOIS da sua primeira compra
            splits_periodo = splits[splits.index.tz_localize(None) >= dados['primeira_compra']]
            fator_split = splits_periodo.prod() if not splits_periodo.empty else 1.0
            
            # Recalcula a quantidade real que você tem hoje e esmaga o Preço Médio para baixo
            qtd_ajustada = dados['quantidade'] * fator_split
            pm_ajustado = dados['valor_investido'] / qtd_ajustada if qtd_ajustada > 0 else 0
            
            divs = acao.dividends
            divs_periodo = divs[divs.index.tz_localize(None) >= dados['primeira_compra']]
            total_dividendos = divs_periodo.sum() * qtd_ajustada
            divs_12m = divs[divs.index.tz_localize(None) >= data_12m_atras].sum()

            info = acao.info
            lpa, vpa = info.get('trailingEps', 0), info.get('bookValue', 0)
            lpa = lpa if lpa is not None else 0
            vpa = vpa if vpa is not None else 0

        except Exception:
            qtd_ajustada = dados['quantidade']
            pm_ajustado = dados['valor_investido'] / qtd_ajustada if qtd_ajustada > 0 else 0
            preco_atual = pm_ajustado
            total_dividendos, divs_12m, lpa, vpa = 0.0, 0.0, 0, 0

        # --- Matemática com as Cotas Corretas ---
        valor_atual = preco_atual * qtd_ajustada
        var_cota = ((valor_atual / dados['valor_investido']) - 1) * 100 if dados['valor_investido'] > 0 else 0
        var_total = (((valor_atual + total_dividendos) / dados['valor_investido']) - 1) * 100 if dados['valor_investido'] > 0 else 0
        cdi_acum, ipca_acum = calcular_macro_acumulado(df_macro, dados['primeira_compra'])

        dados_performance.append({
            "Ativo": ticker,
            "1º Aporte": dados['primeira_compra'].strftime('%m/%Y'),
            "Qtd Real (C/ Splits)": int(qtd_ajustada),
            "PM Real (R$)": formatar_brl(pm_ajustado),
            "Cotação": formatar_brl(preco_atual),
            "Investido": formatar_brl(dados['valor_investido']),
            "Saldo": formatar_brl(valor_atual),
            "Var. Cota": formatar_pct(var_cota),
            "Retorno Total (c/ Div)": formatar_pct(var_total),
            "CDI Acum.": formatar_pct(cdi_acum),
            "IPCA Acum.": formatar_pct(ipca_acum)
        })

        if lpa > 0 and vpa > 0:
            graham = (22.5 * lpa * vpa) ** 0.5
            margem_g = (graham / preco_atual) - 1
        else:
            graham, margem_g = 0, 0

        if divs_12m > 0:
            bazin = divs_12m / 0.06
            margem_b = (bazin / preco_atual) - 1
        else:
            bazin, margem_b = 0, 0

        dados_valuation.append({
            "Ativo": ticker,
            "Cotação": formatar_brl(preco_atual),
            "LPA": formatar_brl(lpa),
            "VPA": formatar_brl(vpa),
            "Div. 12m": formatar_brl(divs_12m),
            "Preço Justo (Graham)": formatar_brl(graham),
            "Margem Graham": f"{'🟢' if margem_g > 0 else '🔴'} {formatar_pct(margem_g * 100)}" if graham > 0 else "N/A",
            "Preço Teto (Bazin)": formatar_brl(bazin),
            "Margem Bazin": f"{'🟢' if margem_b > 0 else '🔴'} {formatar_pct(margem_b * 100)}" if bazin > 0 else "N/A"
        })
        progress_bar.progress((i + 1) / total_ativos)
    
    tab1, tab2 = st.tabs(["📈 Rentabilidade e Retorno Total", "🔎 Valuation (Graham & Bazin)"])
    with tab1: st.dataframe(pd.DataFrame(dados_performance), use_container_width=True, hide_index=True)
    with tab2: st.dataframe(pd.DataFrame(dados_valuation), use_container_width=True, hide_index=True)

else:
    st.info("Aguardando o upload do arquivo para aplicar a correção de Desdobramentos e Bonificações.")
