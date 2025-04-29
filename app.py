import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

def processar_excel(arquivo):
    df = pd.read_excel(arquivo)

    colunas_excluir = [
        'PATM', 'impressora', 'multifuncional', 'color', 'nºsérie', 'observações',
        'Valores Mono', 'Valores Color', 'Valor total', 'Total Colorido',
        'Franquia', 'Excedente', 'Taxa Fixa Tomo_OKI', 'Consumo', 'Total Geral',
        'Coluna Desnecessária 1', 'Coluna Desnecessária 2'
    ]
    df = df.drop(columns=[col for col in colunas_excluir if col in df.columns])

    if 'Setor' in df.columns:
        df = df[~df['Setor'].str.contains('Trafos:.*4102/4018/4110/4117/4124|Trafos:.*4784/4709/4788', na=False)]

    if 'Contador Atual' not in df.columns or 'Contador Anterior' not in df.columns or 'Setor' not in df.columns:
        st.error("Não foi possível localizar as colunas obrigatórias.")
        return None

    df['Contador Atual'] = pd.to_numeric(df['Contador Atual'], errors='coerce').fillna(0)
    df['Contador Anterior'] = pd.to_numeric(df['Contador Anterior'], errors='coerce').fillna(0)

    df['Consumo Calculado'] = df['Contador Atual'] - df['Contador Anterior']
    consumo_setor = df.groupby('Setor')['Consumo Calculado'].sum().reset_index()

    return consumo_setor

def criar_kpi(titulo, valor, cor):
    st.markdown(f"""
        <div style="background-color:{cor}; padding:15px; border-radius:12px; text-align:center;">
            <h3 style="color:white;">{titulo}</h3>
            <h2 style="color:white;">{valor}</h2>
        </div>
    """, unsafe_allow_html=True)

# Streamlit
st.set_page_config(page_title="Analisador de Consumo", layout="wide")
st.title('📈 Dashboard de Consumo de Impressoras')

st.write('**1. Anexe o arquivo do mês anterior:**')
arquivo_mes_anterior = st.file_uploader('Mês Anterior', type=['xlsx'], key='mes_anterior')

st.write('**2. Anexe o arquivo do mês atual:**')
arquivo_mes_atual = st.file_uploader('Mês Atual', type=['xlsx'], key='mes_atual')

if arquivo_mes_anterior and arquivo_mes_atual:
    consumo_anterior = processar_excel(arquivo_mes_anterior)
    consumo_atual = processar_excel(arquivo_mes_atual)

    if consumo_anterior is not None and consumo_atual is not None:
        consumo_anterior = consumo_anterior.rename(columns={'Consumo Calculado': 'Consumo Anterior'})
        consumo_atual = consumo_atual.rename(columns={'Consumo Calculado': 'Consumo Atual'})

        comparativo = pd.merge(consumo_anterior, consumo_atual, on='Setor', how='outer').fillna(0)

        comparativo['Diferença'] = comparativo['Consumo Atual'] - comparativo['Consumo Anterior']

        comparativo['Porcentagem Variação'] = comparativo.apply(
            lambda row: (row['Diferença'] / row['Consumo Anterior'] * 100) if row['Consumo Anterior'] != 0 else 0,
            axis=1
        )

        comparativo['Tendência'] = comparativo['Diferença'].apply(
            lambda x: '⬆️ Aumentou' if x > 0 else ('⬇️ Diminuiu' if x < 0 else '➖ Sem variação')
        )

        setores_disponiveis = comparativo['Setor'].unique()
        setores_selecionados = st.multiselect('🔍 Selecione os Setores que deseja visualizar:', setores_disponiveis)

        if setores_selecionados:
            comparativo = comparativo[comparativo['Setor'].isin(setores_selecionados)]

        # KPIs
        total_aumentou = comparativo[comparativo['Diferença'] > 0]['Diferença'].sum()
        total_diminuiu = comparativo[comparativo['Diferença'] < 0]['Diferença'].sum()
        setores_avaliados = comparativo.shape[0]

        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        with col1:
            criar_kpi("📈 Total Aumento", f"{total_aumentou:.0f}", "#2ecc71")
        with col2:
            criar_kpi("📉 Total Redução", f"{total_diminuiu:.0f}", "#e74c3c")
        with col3:
            criar_kpi("🏢 Setores Avaliados", setores_avaliados, "#3498db")

        st.markdown("---")

        # Top 5 aumentaram e diminuíram
        aumentaram = comparativo[comparativo['Diferença'] > 0].sort_values(by='Diferença', ascending=False).head(5)
        diminuiram = comparativo[comparativo['Diferença'] < 0].sort_values(by='Diferença', ascending=True).head(5)

        # Gráfico Geral de Setores
        st.subheader('📊 Gráfico Geral de Consumo por Setor')
        fig_geral = px.bar(
            comparativo,
            x='Setor',
            y='Diferença',
            color='Porcentagem Variação',
            color_continuous_scale='Viridis',
            title='Variação Geral de Consumo por Setor'
        )
        fig_geral.update_layout(xaxis_title='Setor', yaxis_title='Diferença de Consumo', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_geral, use_container_width=True)

        # Escolha de tipo de gráfico para os Top 5
        tipo_grafico = st.selectbox(
            '📊 Escolha o tipo de gráfico para visualizar os Top 5:',
            ('Barras', 'Pizza')
        )

        # Gráfico de setores que aumentaram
        st.subheader('📈 Top 5 Setores com Aumento de Consumo')

        if tipo_grafico == 'Barras':
            fig_aumentou = px.bar(
                aumentaram,
                x='Diferença',
                y='Setor',
                orientation='h',
                color='Porcentagem Variação',
                color_continuous_scale='Greens',
                text='Porcentagem Variação',
                title='Top 5 - Aumento (%)'
            )
            fig_aumentou.update_traces(texttemplate='%{text:.2f}%', textposition='inside')
            fig_aumentou.update_layout(xaxis_title='Diferença', yaxis_title='', plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_aumentou, use_container_width=True)

        else:  # tipo Pizza
            fig_pizza_aumentou = px.pie(
                aumentaram,
                names='Setor',
                values='Diferença',
                color_discrete_sequence=px.colors.sequential.Greens,
                title='Top 5 - Aumento (Pizza)'
            )
            fig_pizza_aumentou.update_traces(textinfo='percent+label')
            st.plotly_chart(fig_pizza_aumentou, use_container_width=True)

        # Gráfico de setores que diminuíram
        st.subheader('📉 Top 5 Setores com Redução de Consumo')

        if tipo_grafico == 'Barras':
            fig_diminuiu = px.bar(
                diminuiram,
                x='Diferença',
                y='Setor',
                orientation='h',
                color='Porcentagem Variação',
                color_continuous_scale='Reds',
                text='Porcentagem Variação',
                title='Top 5 - Redução (%)'
            )
            fig_diminuiu.update_traces(texttemplate='%{text:.2f}%', textposition='inside')
            fig_diminuiu.update_layout(xaxis_title='Diferença', yaxis_title='', plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_diminuiu, use_container_width=True)

        else:  # tipo Pizza
            fig_pizza_diminuiu = px.pie(
                diminuiram,
                names='Setor',
                values='Diferença',
                color_discrete_sequence=px.colors.sequential.Reds,
                title='Top 5 - Redução (Pizza)'
            )
            fig_pizza_diminuiu.update_traces(textinfo='percent+label')
            st.plotly_chart(fig_pizza_diminuiu, use_container_width=True)

        # Download Excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Comparativo')
            processed_data = output.getvalue()
            return processed_data

        excel_download = to_excel(comparativo)
        st.download_button(
            label="📥 Baixar Comparativo em Excel",
            data=excel_download,
            file_name="comparativo_consumo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
gi