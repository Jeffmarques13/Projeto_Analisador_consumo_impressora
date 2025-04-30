import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import openpyxl
from fpdf import FPDF

# Função para processar os dados do Excel
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

# Função para criar o Excel com dados separados
@st.cache_data
def to_excel(comparativo):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        comparativo.to_excel(writer, index=False, sheet_name='Comparativo')
        comparativo[comparativo['Diferença'] > 0].to_excel(writer, index=False, sheet_name='Consumo Aumentado')
        comparativo[comparativo['Diferença'] < 0].to_excel(writer, index=False, sheet_name='Consumo Diminuido')
    return output.getvalue()

# Função para criar KPI
def criar_kpi(titulo, valor, cor):
    st.markdown(f"""
        <div style="background-color:{cor}; padding:15px; border-radius:12px; text-align:center; box-shadow: 2px 2px 5px rgba(0,0,0,0.2);">
            <h3 style="color:white; margin-bottom:5px;">{titulo}</h3>
            <h2 style="color:white;">{valor}</h2>
        </div>
    """, unsafe_allow_html=True)

# Função para gerar PDF
def gerar_pdf(comparativo):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Relatório de Consumo de Impressoras", ln=True, align="C")
    pdf.ln(10)

    for index, row in comparativo.iterrows():
        pdf.cell(200, 10, txt=f"Setor: {row['Setor']} | Consumo Anterior: {row['Consumo Anterior']} | Consumo Atual: {row['Consumo Atual']} | Diferença: {row['Diferença']} | Variação: {row['Porcentagem Variação']}", ln=True)

    # Gera o PDF em memória
    pdf_output = pdf.output(dest='S').encode('latin1')  # Utilizando .encode() para converter para bytes

    return pdf_output  # Retorna os dados binários do PDF para o download

# Configuração da página do Streamlit
st.set_page_config(page_title="Analisador de Consumo", layout="wide")
st.title('📈 Dashboard de Consumo de Impressoras')

# Upload dos arquivos
st.write('**1. Anexe o arquivo do mês anterior:**')
arquivo_mes_anterior = st.file_uploader('Mês Anterior', type=['xlsx'], key='mes_anterior')

st.write('**2. Anexe o arquivo do mês atual:**')
arquivo_mes_atual = st.file_uploader('Mês Atual', type=['xlsx'], key='mes_atual')

# Quando os dois arquivos forem carregados, processar os dados
if arquivo_mes_anterior and arquivo_mes_atual:
    consumo_anterior, dados_mes_anterior = processar_excel(arquivo_mes_anterior)
    consumo_atual, dados_mes_atual = processar_excel(arquivo_mes_atual)

    if consumo_anterior is not None and consumo_atual is not None:

        if consumo_anterior.equals(consumo_atual):
            st.warning("⚠️ Os dados dos dois arquivos parecem ser iguais. Verifique se os arquivos estão corretos.")

        consumo_anterior = consumo_anterior.rename(columns={'Consumo Calculado': 'Consumo Anterior'})
        consumo_atual = consumo_atual.rename(columns={'Consumo Calculado': 'Consumo Atual'})

        comparativo = pd.merge(consumo_anterior, consumo_atual, on='Setor', how='outer').fillna(0)

        comparativo = comparativo[(comparativo['Consumo Anterior'] != 0) & (comparativo['Consumo Atual'] != 0)]

        comparativo['Diferença'] = comparativo['Consumo Atual'] - comparativo['Consumo Anterior']
        comparativo['Porcentagem Variação'] = comparativo.apply(
            lambda row: (abs(row['Diferença']) / row['Consumo Anterior'] * 100) if row['Consumo Anterior'] != 0 else 0,
            axis=1
        )

        comparativo['Porcentagem Variação'] = comparativo['Porcentagem Variação'].map('{:.2f}%'.format)

        comparativo['Tendência'] = comparativo.apply(
            lambda row: '⬇️ Consumo diminuiu' if row['Diferença'] < 0 else (
                        '⬆️ Consumo aumentou' if row['Diferença'] > 0 else '➖ Sem variação'), axis=1
        )

        st.session_state['comparativo'] = comparativo

        st.markdown("---")
        st.subheader('📋 Lista de Todos os Setores')

        setores_disponiveis = comparativo['Setor'].unique().tolist()
        setor_clicado = st.multiselect('🔎 Selecione um ou mais setores (inclua "Todos os setores"):', 
                                      options=['Todos os Setores'] + setores_disponiveis,
                                      default=['Todos os Setores'])

        if 'Todos os Setores' not in setor_clicado:
            comparativo = comparativo[comparativo['Setor'].isin(setor_clicado)]

        def colorir_tendencia(row):
            if row['Tendência'] == '⬇️ Consumo diminuiu':
                return ['background-color: #d4efdf'] * len(row)  # Linha verde claro
            elif row['Tendência'] == '⬆️ Consumo aumentou':
                return ['background-color: #f5b7b1'] * len(row)  # Linha vermelho claro
            else:
                return [''] * len(row)

        st.dataframe(comparativo[['Setor', 'Consumo Anterior', 'Consumo Atual', 'Diferença', 'Porcentagem Variação', 'Tendência']].style.apply(colorir_tendencia, axis=1))

        # Top 5 Consumo Aumentou
        top_5_aumento = comparativo[comparativo['Diferença'] > 0].nlargest(5, 'Diferença')
        st.subheader("🔥 Top 5 Consumo Aumentado")
        st.dataframe(top_5_aumento[['Setor', 'Consumo Anterior', 'Consumo Atual', 'Diferença', 'Porcentagem Variação']])

        # Top 5 Consumo Diminuiu
        top_5_diminuiu = comparativo[comparativo['Diferença'] < 0].nsmallest(5, 'Diferença')
        st.subheader("❄️ Top 5 Consumo Diminuído")
        st.dataframe(top_5_diminuiu[['Setor', 'Consumo Anterior', 'Consumo Atual', 'Diferença', 'Porcentagem Variação']])

        total_aumentou = comparativo[comparativo['Diferença'] > 0]['Diferença'].sum()
        total_diminuiu = comparativo[comparativo['Diferença'] < 0]['Diferença'].sum()
        setores_avaliados = comparativo.shape[0]
        saldo_total = total_aumentou + total_diminuiu

        if saldo_total > 0:
            tendencia_geral = "⬆️ Aumento geral no consumo"
            cor_tendencia = "#e74c3c"
        elif saldo_total < 0:
            tendencia_geral = "⬇️ Redução geral no consumo"
            cor_tendencia = "#2ecc71"
        else:
            tendencia_geral = "➖ Sem variação geral"
            cor_tendencia = "#bdc3c7"

        col1, col2, col3 = st.columns(3)
        with col1:
            criar_kpi("📈 Total Aumento", f"{total_aumentou:.0f}", "#2ecc71")
        with col2:
            criar_kpi("📉 Total Redução", f"{total_diminuiu:.0f}", "#e74c3c")
        with col3:
            criar_kpi("🏢 Setores Avaliados", setores_avaliados, "#3498db")

        st.subheader('📊 Gráfico Geral de Consumo por Setor')
        fig_geral = px.bar(
            comparativo,
            x='Setor',
            y='Diferença',
            color='Tendência',
            text='Porcentagem Variação',
            color_discrete_map={
                '⬇️ Consumo diminuiu': '#d4efdf',
                '⬆️ Consumo aumentou': '#f5b7b1',
                '➖ Sem variação': '#bdc3c7'
            }
        )
        fig_geral.update_layout(showlegend=False)
        st.plotly_chart(fig_geral)

        st.download_button(
            label="💾 Baixar Dados como Excel",
            data=to_excel(comparativo),
            file_name="comparativo_consumo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="💾 Baixar Relatório em PDF",
            data=gerar_pdf(comparativo)
            file_name="relatorio_consumo.pdf",
            mime="application/pdf"
        )
