import pandas as pd
import os

ARQ_SETOR = "setores.xlsx"
ARQ_EMPRESA = "empresa.xlsx"

def carregar_base():
    if not os.path.exists(ARQ_SETOR):
        raise FileNotFoundError(f"Arquivo {ARQ_SETOR} não encontrado!")
    return pd.read_excel(ARQ_SETOR)

def carregar_empresa():
    if not os.path.exists(ARQ_EMPRESA):
        raise FileNotFoundError(f"Arquivo {ARQ_EMPRESA} não encontrado!")
    return pd.read_excel(ARQ_EMPRESA)

def obter_media_setor(df, setor_nome):
    mask = df['setor'].str.lower() == setor_nome.strip().lower()
    if not mask.any():
        raise KeyError(f"Setor '{setor_nome}' não encontrado em {ARQ_SETOR}")
    return df[mask].iloc[0]
