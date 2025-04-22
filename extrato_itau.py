import pandas as pd
from datetime import datetime

def extrair_fluxo_itau(path_excel: str) -> tuple:
    df = pd.read_excel(path_excel, sheet_name="Lançamentos", skiprows=8)
    df.columns = ["Data", "Descricao", "Razao Social", "CPF/CNPJ", "Valor (R$)", "Saldo (R$)"]
    
    df = df[pd.to_numeric(df["Valor (R$)"], errors="coerce").notnull()].copy()
    
    df["Valor (R$)"] = df["Valor (R$)"].astype(float)
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.strftime("%Y-%m-%d")
    
    df = df[df["Data"].notnull()]

    entradas = df[df["Valor (R$)"] > 0]
    saidas = df[df["Valor (R$)"] < 0]

    recebidos = entradas.groupby("Data")["Valor (R$)"].sum().to_dict()
    pagas = saidas.groupby("Data")["Valor (R$)"].sum().to_dict()

    return recebidos, pagas
