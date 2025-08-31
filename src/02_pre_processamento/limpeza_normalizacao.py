# poetry run limpeza_normalizacao.py
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from scipy.stats import zscore
import unicodedata

# 1) Selecionar arquivo Excel (.xlsx)
def escolher_arquivo_excel() -> str:
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return caminho

# 2) Normalizar nomes de colunas
def normalizar_nomes_colunas(df: pd.DataFrame) -> pd.DataFrame:
    def _norm(s: str) -> str:
        if not isinstance(s, str):
            s = str(s)
        s = s.strip()
        s = unicodedata.normalize("NFKC", s)
        return s
    df = df.copy()
    df.columns = [_norm(c) for c in df.columns]
    return df

# 3) Validar escala [-3, 3]
def validar_escala(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    escala_valida = {-3, -2, -1, 0, 1, 2, 3}
    df_validado = df.copy()
    for coluna in colunas:
        df_validado[coluna] = df_validado[coluna].apply(
            lambda x: x if x in escala_valida else np.nan
        )
    return df_validado

# 4) Remover duplicatas
def remover_duplicatas(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop_duplicates()

# 5) Remover linhas com NaN
def remover_linhas_incompletas(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    return df.dropna(subset=colunas)

# 6) Converter colunas para numérico
def converter_para_numerico(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    df_convertido = df.copy()
    df_convertido[colunas] = df_convertido[colunas].apply(pd.to_numeric, errors="coerce")
    return df_convertido

# 7) Aplicar z-score (robusto)
def aplicar_zscore(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    df_z = df.copy()
    for c in colunas:
        serie = df_z[c]
        mu = serie.mean()
        sd = serie.std(ddof=0)
        if sd == 0 or np.isnan(sd):
            df_z[c] = 0.0
        else:
            df_z[c] = (serie - mu) / sd
    return df_z

# 8) Salvar como Excel (.xlsx)
def salvar_excel(df: pd.DataFrame) -> None:
    root = tk.Tk()
    root.withdraw()
    caminho_saida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Salvar arquivo como"
    )
    df.to_excel(
        caminho_saida,
        index=False,
        engine="openpyxl"
    )
    print(f"\n✅ Arquivo com z-scores salvo em Excel: {caminho_saida}")

# 9) Execução principal
if __name__ == "__main__":
    print("🔍 Escolha o arquivo Excel com os dados brutos...")
    caminho = escolher_arquivo_excel()
    df = pd.read_excel(caminho)
    print("📄 Arquivo carregado com sucesso.")

    # Normalizar cabeçalhos
    df = normalizar_nomes_colunas(df)

    print("\n📄 Colunas disponíveis:")
    print(df.columns.tolist())

    # Detectar colunas de prazer
    colunas_respostas = [col for col in df.columns if isinstance(col, str) and col.startswith("p_")]
    if not colunas_respostas:
        raise ValueError("Nenhuma coluna de prazer (prefixo 'p_') foi encontrada no arquivo.")

    print(f"\n📌 Colunas de prazer detectadas ({len(colunas_respostas)}):")
    print(colunas_respostas)

    # Converter p/ numérico
    df = converter_para_numerico(df, colunas_respostas)

    # Validar escala
    df = validar_escala(df, colunas_respostas)

    # Limpeza
    linhas_antes = len(df)
    df = remover_duplicatas(df)
    df = remover_linhas_incompletas(df, colunas_respostas)
    linhas_depois = len(df)
    print(f"\n🧹 Linhas removidas (duplicadas/incompletas): {linhas_antes - linhas_depois}")

    # Aplicar z-score
    df_zscore = aplicar_zscore(df, colunas_respostas)

    # Checagem rápida
    medias = df_zscore[colunas_respostas].mean().round(3)
    desvios = df_zscore[colunas_respostas].std(ddof=0).round(3)
    print("\n📈 Checagem pós z-score (média ~0, desvio ~1):")
    print("Means:", medias.to_dict())
    print("Stds :", desvios.to_dict())

    # Salvar como Excel
    salvar_excel(df_zscore)






