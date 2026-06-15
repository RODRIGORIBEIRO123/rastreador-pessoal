import streamlit as st
import pandas as pd
import yfinance as yf
from bcb import sgs

st.set_page_config(page_title="Meu Rastreador de Investimentos", layout="wide")
st.title("📊 Análise de Retorno Total da Carteira")

# Botão para o usuário subir a planilha na hora, sem salvar no servidor
arquivo_csv = st.file_uploader("Faça o upload do seu arquivo de Negociação (.csv)", type=["csv"])

if arquivo_csv is not None:
    # 1. Carregar o arquivo lido na hora
    df = pd.read_csv(arquivo_csv)
    
    # Limpeza e formatação
    df['Preço'] = df['Preço'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
    df['Valor'] = df['Valor'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
    
    ativos = {}
    
    for _, row in df.iterrows():
        ticker = row['Código de Negociação']
        qtd = row['Quantidade'] if row['Tipo de Movimentação'] == 'Compra' else -row['Quantidade']
        valor = row['Valor'] if row['Tipo de Movimentação'] == 'Compra' else -row['Valor']
        
        if ticker not in ativos:
            ativos[ticker] = {'quantidade': 0, 'valor_investido': 0}
            
        ativos[ticker]['quantidade'] += qtd
        if ativos[ticker]['quantidade'] > 0:
            ativos[ticker]['valor_investido'] += valor
            
    # Filtrar apenas ativos atuais
    carteira_ativa = {k: v for k, v in ativos.items() if v['quantidade'] > 0}
    
    st.subheader(f"Foram encontrados {len(carteira_ativa)} ativos atualmente na carteira.")
    
    # Preparar dados para exibição na tela
    dados_tabela = []
    for ticker, dados in carteira_ativa.items():
        pm = dados['valor_investido'] / dados['quantidade']
        dados_tabela.append({"Ativo": ticker, "Quantidade": dados['quantidade'], "Preço Médio (R$)": round(pm, 2)})
        
    st.dataframe(pd.DataFrame(dados_tabela), use_container_width=True)

else:
    st.info("Aguardando o upload do arquivo CSV com o histórico da B3 para iniciar a análise.")
