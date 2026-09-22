import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Dashboard de Fluxo de Caixa", layout="wide")
st.title("📊 Painel de Fluxo de Caixa por Obras")

CAMINHO_DO_EXCEL = r"C:\Users\marce\OneDrive\Documentos\VPP Construtora\Projeto Analises\BaseFI.xlsx"
NOME_DA_ABA = "baseFC"  
NOVA_COLUNA_DETALHE = "Fornecedor"  

@st.cache_data(ttl=60)
def carregar_dados_locais(caminho):
    if not os.path.exists(caminho):
        st.error(f"❌ Arquivo não encontrado no caminho: {caminho}")
        st.stop()
        
    df = pd.read_excel(caminho, sheet_name=NOME_DA_ABA, engine='openpyxl')
    df.columns = df.columns.str.strip()
    
    DICIONARIO_COLUNAS = {
        'Data': 'Data',
        'Obra': 'Obra',
        'Status': 'Status',
        'Tipo de Despesa': 'Tipo de Despesa',
        'Valor Lançamento': 'Valor Lançamento',
        NOVA_COLUNA_DETALHE: NOVA_COLUNA_DETALHE
    }
    
    PROG_COLUNAS = {v: k for k, v in DICIONARIO_COLUNAS.items()}
    df = df.rename(columns=PROG_COLUNAS)
    
    colunas_obrigatorias = ['Data', 'Obra', 'Status', 'Tipo de Despesa', 'Valor Lançamento', NOVA_COLUNA_DETALHE]
    for col in colunas_obrigatorias:
        if col not in df.columns:
            st.error(f"❌ A coluna '{col}' não foi encontrada no seu Excel. Verifique os títulos na primeira linha da planilha.")
            st.stop()
            
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    df = df.dropna(subset=['Data'])  
    df = df.sort_values('Data')
    df['Mês_Ano'] = df['Data'].dt.to_period('M').dt.strftime('%Y-%m')
    
    return df

df = carregar_dados_locais(CAMINHO_DO_EXCEL)

st.sidebar.header("Filtros")

status_selecionados = st.sidebar.multiselect(
    "Status da Obra:",
    options=df['Status'].unique(),
    default=df['Status'].unique()
)

df_status_filtrado = df[df['Status'].isin(status_selecionados)]
obras_ordenadas = sorted(df_status_filtrado['Obra'].unique())

obras_selecionadas = st.sidebar.multiselect(
    "Selecione as Obras:",
    options=obras_ordenadas,
    default=obras_ordenadas
)

despesas_selecionadas = st.sidebar.multiselect(
    "Selecione o Tipo de Despesa:",
    options=df['Tipo de Despesa'].unique(),
    default=df['Tipo de Despesa'].unique()
)

st.sidebar.markdown("---")
st.sidebar.subheader("Período de Análise")

data_minima = df['Data'].min().to_pydatetime()
data_maxima = df['Data'].max().to_pydatetime()

periodo_selecionado = st.sidebar.date_input(
    "Selecione o intervalo de datas:",
    value=(data_minima, data_maxima),
    min_value=data_minima,
    max_value=data_maxima,
    format="DD/MM/YYYY"
)

df_base_filtrada = df[
    df['Status'].isin(status_selecionados) & 
    df['Obra'].isin(obras_selecionadas) & 
    df['Tipo de Despesa'].isin(despesas_selecionadas)
]

if not df_base_filtrada.empty and len(periodo_selecionado) == 2:
    # 💡 CORREÇÃO DEFINITIVA: Extraindo explicitamente o índice [0] e [1] da seleção de datas
    data_inicio = pd.to_datetime(periodo_selecionado[0])
    data_fim = pd.to_datetime(periodo_selecionado[1])
    
    df_base_filtrada['Natureza'] = df_base_filtrada['Valor Lançamento'].apply(lambda x: 'Receitas (Entradas)' if x >= 0 else 'Despesas (Saídas)')
    df_natureza_mensal = df_base_filtrada.groupby(['Mês_Ano', 'Natureza'], as_index=False)['Valor Lançamento'].sum()
    df_natureza_mensal = df_natureza_mensal.sort_values('Mês_Ano')
    
    df_natureza_mensal['Acumulado Bruto'] = df_natureza_mensal.groupby('Natureza')['Valor Lançamento'].cumsum()
    df_natureza_mensal['Acumulado'] = df_natureza_mensal['Acumulado Bruto'].abs()
    
    df_natureza_filtrado_tempo = df_natureza_mensal[
        (df_natureza_mensal['Mês_Ano'] >= data_inicio.strftime('%Y-%m')) & 
        (df_natureza_mensal['Mês_Ano'] <= data_fim.strftime('%Y-%m'))
    ]
    
    df_diario_filtrado_tempo = df_base_filtrada[
        (df_base_filtrada['Data'] >= data_inicio) & 
        (df_base_filtrada['Data'] <= data_fim)
    ]
    
    df_negativos = df_diario_filtrado_tempo[df_diario_filtrado_tempo['Valor Lançamento'] < 0]
    df_barras_dados = df_negativos.groupby('Tipo de Despesa', as_index=False)['Valor Lançamento'].sum()
    df_barras_dados['Valor Lançamento'] = df_barras_dados['Valor Lançamento'].abs()
    df_barras_dados = df_barras_dados.sort_values('Valor Lançamento', ascending=True)
    
    total_mensal_selecionado = df_diario_filtrado_tempo['Valor Lançamento'].sum()
    total_acumulado_historico = df_base_filtrada['Valor Lançamento'].sum()

else:
    df_natureza_filtrado_tempo = pd.DataFrame()
    df_barras_dados = pd.DataFrame()
    df_diario_filtrado_tempo = pd.DataFrame()
    total_mensal_selecionado, total_acumulado_historico = 0, 0

card1, card2 = st.columns(2)
with card1:
    st.metric(
        label="💰 Saldo Líquido do Período Selecionado", 
        value=f"R$ {total_mensal_selecionado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )
with card2:
    st.metric(
        label="📈 Saldo Integral Consolidado (Toda a Planilha)", 
        value=f"R$ {total_acumulado_historico:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

st.markdown("---") 

col1, col2 = st.columns(2)

with col1:
    st.subheader("📈 Ponto de Equilíbrio")
    if not df_natureza_filtrado_tempo.empty:
        fig_linha = px.line(
            df_natureza_filtrado_tempo, 
            x='Mês_Ano', 
            y='Acumulado', 
            color='Natureza',
            color_discrete_map={'Receitas (Entradas)': '#2e7d32', 'Despesas (Saídas)': '#c62828'}, 
            labels={'Acumulado': 'Volume Acumulado Bruto (R$)', 'Mês_Ano': 'Mês/Ano', 'Natureza': 'Fluxo'},
            markers=True
        )
        st.plotly_chart(fig_linha, use_container_width=True)
    else:
        st.info("💡 Selecione um intervalo válido no calendário da barra lateral para exibir o gráfico.")

with col2:
    st.subheader("📊 Divisão de Despesas e Custos do Período")
    if not df_barras_dados.empty:
        fig_barras = px.bar(
            df_barras_dados, 
            x='Valor Lançamento', 
            y='Tipo de Despesa',
            orientation='h',
            text_auto='.2s',
            labels={'Valor Lançamento': 'Total Gasto (R$)', 'Tipo de Despesa': 'Categoria'},
            color='Tipo de Despesa'
        )
        fig_barras.update_layout(showlegend=False, yaxis={'categoryorder':'total ascending'})
        evento_selecao = st.plotly_chart(fig_barras, use_container_width=True, on_select="rerun", key="bar_chart")
    else:
        st.info("💡 Não existem despesas negativas no intervalo de datas selecionado.")

st.markdown("---")

click_data = st.session_state.get("bar_chart", {})
has_selection = click_data.get("selection", {}).get("points", []) if click_data else []

if not df_barras_dados.empty and has_selection:
    categoria_clicada = has_selection[0].get("y")
    
    if categoria_clicada:
        st.subheader(f"🔍 Detalhamento Analítico de Custos: {categoria_clicada}")
        
        df_detalhe_categoria = df_diario_filtrado_tempo[df_diario_filtrado_tempo['Tipo de Despesa'] == categoria_clicada]
        
        tabela_drilldown = df_detalhe_categoria[['Data', 'Obra', NOVA_COLUNA_DETALHE, 'Valor Lançamento']].copy()
        tabela_drilldown['Data'] = tabela_drilldown['Data'].dt.strftime('%d/%m/%Y')
        
        tabela_drilldown_exibicao = tabela_drilldown.style.format({
            'Valor Lançamento': "R$ {:,.2f}"
        })
        
        st.dataframe(tabela_drilldown_exibicao, use_container_width=True, hide_index=True)
else:
    st.info("💡 Dica de Análise: Clique em qualquer barra de custos acima para abrir o detalhamento diário com os Fornecedores.")