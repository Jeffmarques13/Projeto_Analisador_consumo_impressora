# utils.py

import pandas as pd

# Função para processar os arquivos Excel
def processar_excel(file):
    try:
        df = pd.read_excel(file)
        # Adapte o código de acordo com o seu processamento
        return df, None
    except Exception as e:
        return None, str(e)

# Função para exportar DataFrame para Excel
def to_excel(df):
    return df.to_excel(index=False)

# Função para criar um KPI em HTML
def criar_kpi(titulo, valor, cor):
    return f"<div style='color:{cor};'><strong>{titulo}</strong>: {valor}</div>"
