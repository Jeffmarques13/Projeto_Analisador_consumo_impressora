import streamlit as st
import pandas as pd
import plotly.express as px
import re

def encontrar_coluna(candidatos, df_columns):
    """
    Função que tenta encontrar o nome correto da coluna no DataFrame
    comparando variações.
    """
    for candidato in candidatos:
        for coluna in df_columns:
            if re.sub(r'\W+', '', coluna.lower()) == re.sub(r'\W+', '', candidato.lower()):
                return coluna
    return None

def processar_excel(arquivo):
    # Ler o Excel
    df = pd.read_excel(arquivo)

    # Padronizar colunas para minúsculo e tirar espaços
    df.columns = df.columns.str.strip()

    # Exibir as colunas para conferência
    st.write("📋 Colunas encontradas:", df.columns.tolist())

    # Procurar as colunas de contador atual e contador anterior
    contador_atual = encontrar_coluna(['contador atual', 'contador_atual', 'contadorAtual'], df.columns)
    contador_anterior = encontrar_coluna(['contador anterior', 'contador_anterior', 'contadorAnterior'], df.columns)

    if contador_atual and contador_anterior:
        df['Consumo Calculado'] = df[contador_atual] - df[contador_anterior]
    else:
        st.error("❌ Não foi possível encontrar as colunas de 'Contador Atual' e 'Contador Anterior'.")
        return None

    # Procurar a coluna de setor
    setor_coluna = encontrar_coluna(['setor', 'Setor'], df.columns)

    if setor_coluna:
        consumo_setor = df.groupby(setor_coluna)['Consumo Calculado'].sum().reset_index()
        return consumo_setor
    else:
        st.error("❌ Coluna 'Setor' não encontrada.")
        return None

# Streamlit
st.title('📊 Analisador de Consumo de Impressoras por Setor')

st.write('**1. Anexe o arquivo do mês:**')
arquivo_mes = st.file_uploader('Arquivo do Mês', type=['xlsx'], key='mes')

if arquivo_mes:
    consumo = processar_excel(arquivo_mes)

    if consumo is not None:
        st.subheader('Resultado do Consumo:')
        st.dataframe(consumo)

        # Gráfico de barras
        st.subheader('📈 Gráfico de Consumo por Setor')
        fig = px.bar(
            consumo,
            x='Setor',
            y='Consumo Calculado',
            color='Setor',
            title='Consumo de Impressões por Setor',
            labels={'Consumo Calculado': 'Consumo', 'Setor': 'Setor'},
            height=500
        )
        st.plotly_chart(fig)

        # Botão para baixar o resultado
        excel_download = consumo.to_excel(index=False, engine='openpyxl')
        st.download_button(
            label="📥 Baixar Consumo em Excel",
            data=excel_download,
            file_name="consumo_setor.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
