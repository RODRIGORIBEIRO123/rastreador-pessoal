import streamlit as st
import pandas as pd
import yfinance as yf
from bcb import sgs
import re
from datetime import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Rastreador de Retorno Total", layout="wide")
st.title("📊 Análise Profissional de Carteira (Total Return)")

# --- FUNÇÕES DE FORMATAÇÃO E CÁLCULO ---
def formatar_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_pct(valor):
    return f"{valor:,.2f}%".replace(".", ",")

def eh_opcao_ou_futuro(ticker):
    # Ignora mini-índice e dólar
    if ticker.startswith(('WIN', 'WDO', 'IND', 'DOL')):
        return True
    t = ticker[:-1] if str(ticker).endswith('F') else str(ticker)
    match = re.match(r'^[A-Z]{4}[A-Z]\d+', t)
    # Identifica opções (ex: PETRC285) ignorando FIIs (11) e BDRs (34, 39)
    if match and not t.endswith('11'):
        if len(t) > 6 or (len(t) == 6 and not t.endswith(('34', '39', '11'))):
            return True
    return False

@st.cache_data(ttl=86400) # Cache de 24h para não sobrecarregar o BC
def obter_dados_macro():
    try:
        # CDI diário (12) e IPCA mensal (433)
        macro = sgs.get({'CDI': 12, 'IPCA': 433}, start='2019-01-01')
        macro['CDI'] = macro['CDI'] / 100
        macro['IPCA'] = macro['IPCA'] / 100
        return macro
    except Exception as e:
        return None

def calcular_macro_acumulado(df_macro, data_inicio):
    if df_macro is None or df_macro.empty:
        return 0.0, 0.0
    filtro = df_macro.loc[data_inicio:]
    cdi_acum = (1 + filtro['CDI'].dropna()).prod() - 1
    ipca_acum = (1 + filtro['IPCA'].dropna()).prod() - 1
    return cdi_acum * 100, ipca_acum * 100

# --- LÓGICA PRINCIPAL ---
arquivo = st.file_uploader("Upload da planilha de Negociação da B3 (.xlsx ou .csv)", type=["xlsx", "csv"])

if arquivo:
    with st.spinner("Processando histórico e higienizando dados da B3..."):
        if arquivo.name.endswith('.csv'):
            df = pd.read_csv(arquivo, sep=';', encoding='latin1')
        else:
            df = pd.read_excel(arquivo)
            
        df.columns = df.columns.str.strip()
        df['Data do Negócio'] = pd.to_datetime(df['Data do Negócio'], format='%d/%m/%Y', errors='coerce')
        
        # Limpeza monetária
        df['Preço'] = df['Preço'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        df['Valor'] = df['Valor'].astype(str).replace({'R\$': '', '\.': '', ',': '.'}, regex=True).astype(float)
        
        # Ordenar do mais antigo para o mais novo
        df = df.sort_values('Data do Negócio')
        
        posicoes = {}
        
        for _, row in df.iterrows():
            ticker_orig = str(row['Código de Negociação']).strip()
            
            if eh_opcao_ou_futuro(ticker_orig):
                continue
                
            # Unifica fracionário com lote padrão
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
        
    st.success(f"Base processada! {len(carteira_ativa)} ativos reais identificados.")
    
    # Carrega base macroeconômica
    df_macro = obter_dados_macro()
    if df_macro is None:
        st.warning("Não foi possível conectar ao Banco Central no momento. IPCA e CDI não serão calculados.")

    # Conexão com mercado para cotações e dividendos
    st.write("### Auditoria de Performance")
    progress_bar = st.progress(0)
    
    dados_finais = []
    ativos_lista = sorted(carteira_ativa.items(), key=lambda x: x[1]['valor_investido'], reverse=True)
    total_ativos = len(ativos_lista)
    
    for i, (ticker, dados) in enumerate(ativos_lista):
        try:
            acao = yf.Ticker(f"{ticker}.SA")
            # Histórico desde a primeira compra
            hist = acao.history(start=dados['primeira_compra'].strftime('%Y-%m-%d'))
            preco_atual = hist['Close'].iloc[-1] if not hist.empty else (dados['valor_investido'] / dados['quantidade'])
            
            # Ajuste de Desdobramentos/Grupamentos (Splits)
            splits = acao.splits
            splits_periodo = splits[splits.index.tz_localize(None) >= dados['primeira_compra']]
            fator_split = splits_periodo.prod() if not splits_periodo.empty else 1.0
            
            qtd_ajustada = dados['quantidade'] * fator_split
            pm_ajustado = dados['valor_investido'] / qtd_ajustada if qtd_ajustada > 0 else 0
            
            # Cálculo de Dividendos
            divs = acao.dividends
            divs_periodo = divs[divs.index.tz_localize(None) >= dados['primeira_compra']]
            total_dividendos = divs_periodo.sum() * qtd_ajustada
            
        except Exception:
            # Fallback caso o ticker não exista no Yahoo Finance (ex: fundos muito novos)
            qtd_ajustada = dados['quantidade']
            pm_ajustado = dados['valor_investido'] / qtd_ajustada if qtd_ajustada > 0 else 0
            preco_atual = pm_ajustado
            total_dividendos = 0.0

        # Matemáticas de Retorno
        valor_atual = preco_atual * qtd_ajustada
        lucro_cota = valor_atual - dados['valor_investido']
        lucro_total = lucro_cota + total_dividendos
        
        var_cota = ((valor_atual / dados['valor_investido']) - 1) * 100 if dados['valor_investido'] > 0 else 0
        var_total = (((valor_atual + total_dividendos) / dados['valor_investido']) - 1) * 100 if dados['valor_investido'] > 0 else 0
        
        cdi_acum, ipca_acum = calcular_macro_acumulado(df_macro, dados['primeira_compra'])

        dados_finais.append({
            "Ativo": ticker,
            "1º Aporte": dados['primeira_compra'].strftime('%m/%Y'),
            "Qtd Ajustada": int(qtd_ajustada),
            "PM Ajustado": formatar_brl(pm_ajustado),
            "Cotação Atual": formatar_brl(preco_atual),
            "Total Investido": formatar_brl(dados['valor_investido']),
            "Saldo Atual": formatar_brl(valor_atual),
            "Var. Cota": formatar_pct(var_cota),
            "Retorno Total (c/ Div)": formatar_pct(var_total),
            "IPCA Acum.": formatar_pct(ipca_acum),
            "CDI Acum.": formatar_pct(cdi_acum)
        })
        
        progress_bar.progress((i + 1) / total_ativos)
    
    df_exibicao = pd.DataFrame(dados_finais)
    st.dataframe(df_exibicao, use_container_width=True, hide_index=True)

else:
    st.info("Aguardando o upload do arquivo da B3 para gerar o relatório de rentabilidade profissional.")
