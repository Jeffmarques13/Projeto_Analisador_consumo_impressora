# helpers.py

def get_cor_variacao(variacao):
    if variacao < 0:
        return "green", "Diminuiu"
    else:
        return "red", "Aumentou"
