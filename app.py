import streamlit as st
import pandas as pd

st.set_page_config(page_title="Meu Rastreador de Investimentos", layout="wide")
st.title("📊 Análise da Carteira Atual")

# Atualizado para aceitar Excel e CSV da B3
arquivo = st.file_uploader("Faça o upload do arquivo de Negociação da B3 (.xlsx ou .csv)", type=["csv", "xlsx"])

if arquivo is not None:
    try:
        # Verifica a extensão para ler da forma certa
        if arquivo.name.endswith('.csv'):
            df = pd.read_csv(arquivo, sep=';', encoding='latin1')
        else:
            df = pd.read_excel(arquivo)
        
        # Limpa espaços em branco ocultos nos nomes das colunas (evita o KeyError)
        df.columns = df.columns.str.strip()

        # Limpeza e formatação financeira
        df['Preço'] = df['Preço'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        df['Valor'] = df['Valor'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        
        ativos = {}
        
        for _, row in df.iterrows():
            ticker = str(row['Código de Negociação']).strip()
            qtd = row['Quantidade']
            valor = row['Valor']
            
            if ticker not in ativos:
                ativos[ticker] = {'quantidade': 0, 'valor_investido': 0.0}
                
            if row['Tipo de Movimentação'] == 'Compra':
                ativos[ticker]['quantidade'] += qtd
                ativos[ticker]['valor_investido'] += valor
                
            elif row['Tipo de Movimentação'] == 'Venda':
                # Evita erro se houver venda antes da compra no histórico
                if ativos[ticker]['quantidade'] > 0:
                    # O preço médio atual não muda na venda, apenas retiramos o valor proporcional
                    preco_medio_atual = ativos[ticker]['valor_investido'] / ativos[ticker]['quantidade']
                    ativos[ticker]['quantidade'] -= qtd
                    
                    if ativos[ticker]['quantidade'] > 0:
                        ativos[ticker]['valor_investido'] -= (qtd * preco_medio_atual)
                    else:
                        ativos[ticker]['valor_investido'] = 0.0
                else:
                    ativos[ticker]['quantidade'] -= qtd

        # Filtrar apenas ativos que você ainda tem na carteira
        carteira_ativa = {k: v for k, v in ativos.items() if v['quantidade'] > 0}
        
        st.subheader(f"Foram encontrados {len(carteira_ativa)} ativos atualmente na carteira.")
        
        # Preparar dados para exibição
        dados_tabela = []
        for ticker, dados in carteira_ativa.items():
            pm = dados['valor_investido'] / dados['quantidade']
            dados_tabela.append({
                "Ativo": ticker, 
                "Quantidade Atual": int(dados['quantidade']), 
                "Preço Médio (R$)": round(pm, 2),
                "Capital Alocado (R$)": round(dados['valor_investido'], 2)
            })
            
        # Exibe a tabela ordenada pelo maior peso na carteira
        df_final = pd.DataFrame(dados_tabela).sort_values(by="Capital Alocado (R$)", ascending=False)
        st.dataframe(df_final, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: O sistema não encontrou a coluna esperada. Detalhe técnico: {e}")
        if 'df' in locals():
            st.write("Colunas que o sistema conseguiu ler no seu arquivo:", df.columns.tolist())

else:
    st.info("Aguardando o upload do arquivo Excel (.xlsx) ou CSV com o histórico da B3.")
