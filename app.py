import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import openpyxl
from io import BytesIO

@st.cache_data
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

    df = df[(df['Contador Atual'] != 0) & (df['Contador Anterior'] != 0)]

    df['Consumo Calculado'] = df['Contador Atual'] - df['Contador Anterior']
    consumo_setor = df.groupby('Setor')['Consumo Calculado'].sum().reset_index()

    return consumo_setor, df

@st.cache_data
def to_excel(df, grafico_geral, grafico_aumento, grafico_reducao):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Comparativo')

        # Salvar gráficos como imagens no Excel
        workbook = writer.book
        worksheet = workbook.create_sheet('Gráficos')

        # Salvar gráfico geral
        img_geral = grafico_geral.to_image(format='png')
        img_geral_path = BytesIO(img_geral)
        worksheet.insert_image('A1', 'Grafico_Geral.png', {'image_data': img_geral_path})

        # Salvar gráfico de aumento
        img_aumento = grafico_aumento.to_image(format='png')
        img_aumento_path = BytesIO(img_aumento)
        worksheet.insert_image('A30', 'Grafico_Aumento.png', {'image_data': img_aumento_path})

        # Salvar gráfico de redução
        img_reducao = grafico_reducao.to_image(format='png')
        img_reducao_path = BytesIO(img_reducao)
        worksheet.insert_image('A60', 'Grafico_Reducao.png', {'image_data': img_reducao_path})

    return output.getvalue()

def criar_kpi(titulo, valor, cor):
    st.markdown(f"""
        <div style="background-color:{cor}; padding:15px; border-radius:12px; text-align:center; box-shadow: 2px 2px 5px rgba(0,0,0,0.2);">
            <h3 style="color:white; margin-bottom:5px;">{titulo}</h3>
            <h2 style="color:white;">{valor}</h2>
        </div>
    """, unsafe_allow_html=True)

st.set_page_config(page_title="Analisador de Consumo", layout="wide")
st.title('📈 Dashboard de Consumo de Impressoras')

st.write('**1. Anexe o arquivo do mês anterior:**')
arquivo_mes_anterior = st.file_uploader('Mês Anterior', type=['xlsx'], key='mes_anterior')

st.write('**2. Anexe o arquivo do mês atual:**')
arquivo_mes_atual = st.file_uploader('Mês Atual', type=['xlsx'], key='mes_atual')

if arquivo_mes_anterior and arquivo_mes_atual:
    consumo_anterior, dados_mes_anterior = processar_excel(arquivo_mes_anterior)
    consumo_atual, dados_mes_atual = processar_excel(arquivo_mes_atual)

    if consumo_anterior is not None and consumo_atual is not None:

        if consumo_anterior.equals(consumo_atual):
            st.warning("⚠️ Os dados dos dois arquivos parecem ser iguais. Verifique se os arquivos estão corretos.")

        consumo_anterior = consumo_anterior.rename(columns={'Consumo Calculado': 'Consumo Anterior'})
        consumo_atual = consumo_atual.rename(columns={'Consumo Calculado': 'Consumo Atual'})

        comparativo = pd.merge(consumo_anterior, consumo_atual, on='Setor', how='outer').fillna(0)

        comparativo['Diferença'] = comparativo['Consumo Atual'] - comparativo['Consumo Anterior']

        comparativo = comparativo[~((comparativo['Consumo Atual'] == 0) & (comparativo['Consumo Anterior'] == 0))]

        comparativo['Porcentagem Variação'] = comparativo.apply(
            lambda row: (abs(row['Diferença']) / row['Consumo Anterior'] * 100) if row['Consumo Anterior'] != 0 else 0,
            axis=1
        )

        comparativo['Tendência'] = comparativo.apply(
            lambda row: '⬇️ Consumo diminuiu' if row['Diferença'] < 0 else (
                        '⬆️ Consumo aumentou' if row['Diferença'] > 0 else '➖ Sem variação'), axis=1
        )

        st.session_state['comparativo'] = comparativo

        st.markdown("---")
        st.subheader('📋 Lista de Todos os Setores')

        setores_disponiveis = st.session_state['comparativo']['Setor'].unique().tolist()
        setor_clicado = st.multiselect('🔎 Selecione um ou mais setores (inclua "Todos os setores"):', 
                                      options=['Todos os Setores'] + setores_disponiveis,
                                      default=['Todos os Setores'])

        if 'Todos os Setores' not in setor_clicado:
            comparativo = st.session_state['comparativo'][st.session_state['comparativo']['Setor'].isin(setor_clicado)]
        else:
            comparativo = st.session_state['comparativo']

        st.dataframe(comparativo[['Setor', 'Consumo Anterior', 'Consumo Atual', 'Diferença', 'Porcentagem Variação', 'Tendência']])

        total_aumentou = comparativo[comparativo['Diferença'] > 0]['Diferença'].sum()
        total_diminuiu = comparativo[comparativo['Diferença'] < 0]['Diferença'].sum()
        setores_avaliados = comparativo.shape[0]

        col1, col2, col3 = st.columns(3)
        with col1:
            criar_kpi("📈 Total Aumento", f"{total_aumentou:.0f}", "#2ecc71")
        with col2:
            criar_kpi("📉 Total Redução", f"{total_diminuiu:.0f}", "#e74c3c")
        with col3:
            criar_kpi("🏢 Setores Avaliados", setores_avaliados, "#3498db")

        st.markdown("---")

        aumentaram = comparativo[comparativo['Diferença'] > 0].sort_values(by='Diferença', ascending=False).head(5)
        diminuiram = comparativo[comparativo['Diferença'] < 0].sort_values(by='Diferença', ascending=True).head(5)

        st.subheader('📊 Gráfico Geral de Consumo por Setor')
        fig_geral = px.bar(
            comparativo,
            x='Setor',
            y='Diferença',
            color='Tendência',
            text=comparativo['Porcentagem Variação'].map('{:.2f}%'.format),
            color_discrete_map={
                '⬇️ Consumo diminuiu': '#2ecc71',
                '⬆️ Consumo aumentou': '#e74c3c',
                '➖ Sem variação': '#bdc3c7'
            },
            title='Variação Geral de Consumo por Setor'
        )
        fig_geral.update_layout(xaxis_title='Setor', yaxis_title='Diferença de Consumo', plot_bgcolor='rgba(0,0,0,0)', title_font_size=20)
        st.plotly_chart(fig_geral, use_container_width=True)

        tipo_grafico = st.selectbox('📊 Tipo de gráfico para os Top 5:', ('Barras', 'Pizza'))

        st.subheader('📈 Top 5 Setores com Aumento de Consumo')
        if tipo_grafico == 'Barras':
            aumentaram['Texto Exibido'] = aumentaram.apply(
                lambda row: f"{row['Diferença']:.0f} ({row['Porcentagem Variação']:.2f}%)", axis=1
            )
            fig_aumento = px.bar(
                aumentaram,
                x='Setor',
                y='Diferença',
                text='Texto Exibido',
                color='Tendência',
                color_discrete_map={'⬆️ Consumo aumentou': '#e74c3c'}
            )
            st.plotly_chart(fig_aumento, use_container_width=True)
        else:
            fig_aumento_pizza = px.pie(
                aumentaram,
                names='Setor',
                values='Diferença',
                title="Aumento de Consumo - Top 5",
                color_discrete_map={'⬆️ Consumo aumentou': '#e74c3c'}
            )
            st.plotly_chart(fig_aumento_pizza, use_container_width=True)

        st.subheader('📉 Top 5 Setores com Redução de Consumo')
        if tipo_grafico == 'Barras':
            diminuiram['Texto Exibido'] = diminuiram.apply(
                lambda row: f"{row['Diferença']:.0f} ({row['Porcentagem Variação']:.2f}%)", axis=1
            )
            fig_reducao = px.bar(
                diminuiram,
                x='Setor',
                y='Diferença',
                text='Texto Exibido',
                color='Tendência',
                color_discrete_map={'⬇️ Consumo diminuiu': '#2ecc71'}
            )
            st.plotly_chart(fig_reducao, use_container_width=True)
        else:
            fig_reducao_pizza = px.pie(
                diminuiram,
                names='Setor',
                values='Diferença',
                title="Redução de Consumo - Top 5",
                color_discrete_map={'⬇️ Consumo diminuiu': '#2ecc71'}
            )
            st.plotly_chart(fig_reducao_pizza, use_container_width=True)

        st.markdown("---")

        st.download_button(
            label="💾 Baixar Comparativo com Gráficos como Excel",
            data=to_excel(comparativo, fig_geral, fig_aumento, fig_reducao),
            file_name="comparativo_consumo_impressoras_com_graficos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
