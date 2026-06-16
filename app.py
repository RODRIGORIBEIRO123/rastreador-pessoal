import streamlit as st
import pandas as pd
import yfinance as yf
import re
from datetime import datetime

st.set_page_config(page_title="Rastreador de Carteira", layout="wide")
st.title("📊 Análise de Retorno Total (Com Dividendos)")

def is_option(ticker):
    t = ticker[:-1] if str(ticker).endswith('F') else str(ticker)
    match = re.match(r'^[A-Z]{4}[A-Z]\d+', t)
    if match and not t.endswith('11'):
        if len(t) > 6 or (len(t) == 6 and not t.endswith('34') and not t.endswith('11')):
            return True
    return False

arquivo = st.file_uploader("Upload da planilha da B3 (.xlsx)", type=["xlsx", "csv"])

if arquivo:
    st.info("Lendo arquivo e limpando opções / contratos vencidos...")
    if arquivo.name.endswith('.csv'):
        df = pd.read_csv(arquivo, sep=';', encoding='latin1')
    else:
        df = pd.read_excel(arquivo)
        
    df.columns = df.columns.str.strip()
    df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], format='%d/%m/%Y', errors='coerce')
    df['Preço'] = df['Preço'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
    df['Valor'] = df['Valor'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
    
    posicoes = {}
    
    for _, row in df.iterrows():
        ticker_orig = str(row['Código de Negociação']).strip()
        if is_option(ticker_orig) or ticker_orig.startswith(('WIN', 'WDO')):
            continue
            
        ticker = ticker_orig[:-1] if ticker_orig.endswith('F') else ticker_orig
        qtd = row['Quantidade']
        valor = row['Valor']
        data = row['Data do Negócio']
        
        if ticker not in posicoes:
            posicoes[ticker] = {'quantidade': 0, 'valor_investido': 0.0, 'primeira_compra': data}
            
        if row['Tipo de Movimentação'] == 'Compra':
            posicoes[ticker]['quantidade'] += qtd
            posicoes[ticker]['valor_investido'] += valor
            if data < posicoes[ticker]['primeira_compra']:
                posicoes[ticker]['primeira_compra'] = data
                
        elif row['Tipo de Movimentação'] == 'Venda':
            if posicoes[ticker]['quantidade'] > 0:
                qtd_venda = min(qtd, posicoes[ticker]['quantidade'])
                pm_atual = posicoes[ticker]['valor_investido'] / posicoes[ticker]['quantidade']
                posicoes[ticker]['quantidade'] -= qtd_venda
                posicoes[ticker]['valor_investido'] -= (qtd_venda * pm_atual)
                
            if posicoes[ticker]['quantidade'] <= 0.001:
                posicoes[ticker]['quantidade'] = 0
                posicoes[ticker]['valor_investido'] = 0.0

    carteira_ativa = {k: v for k, v in posicoes.items() if v['quantidade'] > 0}
    
    # Prepara a Análise de Rentabilidade
    st.success(f"Foram identificados {len(carteira_ativa)} ativos reais.")
    st.warning("Buscando cotações e dividendos ao vivo (isso pode levar alguns segundos)...")
    
    dados_finais = []
    
    # Processa os 10 maiores ativos para não demorar muito na tela (você pode tirar esse limite depois)
    ativos_ordenados = sorted(carteira_ativa.items(), key=lambda x: x[1]['valor_investido'], reverse=True)
    
    for ticker, dados in ativos_ordenados:
        pm = dados['valor_investido'] / dados['quantidade']
        
        # Conexão com a internet para pegar cotação e dividendos
        try:
            ticker_yf = yf.Ticker(f"{ticker}.SA")
            hist = ticker_yf.history(period="1d")
            preco_atual = hist['Close'].iloc[-1] if not hist.empty else pm
        except:
            preco_atual = pm
            
        valor_atual = preco_atual * dados['quantidade']
        rentabilidade_sem_div = ((preco_atual / pm) - 1) * 100
        
        dados_finais.append({
            "Ativo": ticker,
            "Qtd": int(dados['quantidade']),
            "Preço Médio": f"R$ {pm:.2f}",
            "Preço Atual": f"R$ {preco_atual:.2f}",
            "Investido": f"R$ {dados['valor_investido']:.2f}",
            "Saldo Atual": f"R$ {valor_atual:.2f}",
            "Variação da Cota (%)": f"{rentabilidade_sem_div:.2f}%",
            "1º Aporte": dados['primeira_compra'].strftime('%m/%Y')
        })
        
    st.dataframe(pd.DataFrame(dados_finais), use_container_width=True)

else:
    st.info("Suba o arquivo para iniciar as conexões de rede e calcular o Retorno.")
