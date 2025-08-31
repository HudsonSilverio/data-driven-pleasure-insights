# poetry run frequencias.py

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import unicodedata

# ---- Formatação global: SEM notação científica (mais amigável p/ leigos) ----
np.set_printoptions(suppress=True)                  # numpy não usa e-notation
pd.set_option("display.float_format", "{:,.4f}".format)  # pandas imprime 4 casas
pd.set_option("display.max_rows", 200)
pd.set_option("display.max_columns", 200)
pd.set_option("display.width", 120)

# 00) Selecionar e carregar Excel
def escolher_arquivo_excel() -> str:
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return caminho

def carregar_excel(caminho: str) -> pd.DataFrame:
    return pd.read_excel(caminho)

# Utilitários
def normalizar_nomes_colunas(df: pd.DataFrame) -> pd.DataFrame:
    def _norm(s: str) -> str:
        if not isinstance(s, str):
            s = str(s)
        s = s.strip()
        s = unicodedata.normalize("NFKC", s)
        return s
    out = df.copy()
    out.columns = [_norm(c) for c in out.columns]
    return out

def detectar_colunas_prazer(df: pd.DataFrame) -> list:
    return [c for c in df.columns if isinstance(c, str) and c.startswith("p_")]

def forcar_numerico(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    out = df.copy()
    out[colunas] = out[colunas].apply(pd.to_numeric, errors="coerce")
    return out

# 01) Médianas (Top maiores e Top menores)
def medianas_por_coluna(df: pd.DataFrame, colunas: list) -> pd.Series:
    return df[colunas].median(skipna=True)

# 02) Médias (Top maiores e Top menores)
def medias_por_coluna(df: pd.DataFrame, colunas: list) -> pd.Series:
    return df[colunas].mean(skipna=True)

# 03) Distribuição por escala
def distribuicao_por_escala(df: pd.DataFrame, colunas: list, escala=(-3, -2, -1, 0, 1, 2, 3)) -> pd.DataFrame:
    dist_dict = {}
    for c in colunas:
        serie = df[c]
        contagens = {p: int((serie == p).sum()) for p in escala}  # imprime como int
        dist_dict[c] = contagens
    dist_df = pd.DataFrame(dist_dict).T
    dist_df = dist_df[list(escala)]
    return dist_df

# 04) Desvio padrão
def desvios_por_coluna(df: pd.DataFrame, colunas: list) -> pd.Series:
    return df[colunas].std(skipna=True)

# Impressões amigáveis
def imprimir_ranking_duplo(titulo_alto: str, titulo_baixo: str, serie: pd.Series, n: int = 10):
    # Ordenações separadas
    serie_sorted_desc = serie.sort_values(ascending=False)
    serie_sorted_asc  = serie.sort_values(ascending=True)

    print("\n" + "=" * 100)
    print(titulo_alto)
    print("=" * 100)
    df_top = serie_sorted_desc.head(n).to_frame(name="value")
    print(df_top.to_string())

    print("\n" + "-" * 100)
    print(titulo_baixo)
    print("-" * 100)
    df_bottom = serie_sorted_asc.head(n).to_frame(name="value")
    print(df_bottom.to_string())

def imprimir_tabela(titulo: str, df: pd.DataFrame, max_rows: int = 25):
    print("\n" + "=" * 100)
    print(titulo)
    print("=" * 100)
    if len(df) > max_rows:
        print(df.head(max_rows).to_string())
        print(f"... ({len(df) - max_rows} linhas restantes não exibidas)")
    else:
        print(df.to_string())

# Execução principal
if __name__ == "__main__":
    print("🔍 Escolha o arquivo Excel (.xlsx) com as respostas (escala -3..3).")
    caminho = escolher_arquivo_excel()
    df = carregar_excel(caminho)
    print("📄 Arquivo carregado.")

    # Normalizar nomes e detectar colunas de prazer
    df = normalizar_nomes_colunas(df)
    colunas_prazer = detectar_colunas_prazer(df)
    if not colunas_prazer:
        raise ValueError("Nenhuma coluna de prazer (prefixo 'p_') foi encontrada.")
    print(f"\n📌 Colunas de prazer detectadas ({len(colunas_prazer)}):")
    print(colunas_prazer)

    # Garantir tipo numérico (evita textos e garante cálculos corretos)
    df = forcar_numerico(df, colunas_prazer)

    # 01) MEDIANAS: Top 10 maiores e Top 10 menores
    medians = medianas_por_coluna(df, colunas_prazer)
    imprimir_ranking_duplo(
        "Top 10 – Highest median rating (by pleasure)",
        "Top 10 – Lowest median rating (by pleasure)",
        medians, n=10
    )

    # 02) MÉDIAS: Top 10 maiores e Top 10 menores
    means = medias_por_coluna(df, colunas_prazer)
    imprimir_ranking_duplo(
        "Top 10 – Highest mean rating (by pleasure)",
        "Top 10 – Lowest mean rating (by pleasure)",
        means, n=10
    )

    # 03) DISTRIBUIÇÃO por ponto da escala (-3..3)
    dist = distribuicao_por_escala(df, colunas_prazer, escala=(-3, -2, -1, 0, 1, 2, 3))
    imprimir_tabela(
        "Distribution of responses per scale point (-3..3) by pleasure (counts)",
        dist, max_rows=25
    )

    # 04) DESVIO PADRÃO por coluna (e ranking)
    stds = desvios_por_coluna(df, colunas_prazer)
    imprimir_ranking_duplo(
        "Top 10 – Highest standard deviation (most variable pleasures)",
        "Top 10 – Lowest standard deviation (most consistent pleasures)",
        stds, n=10
    )

    print("\n✅ Finished: results printed without scientific notation.")
