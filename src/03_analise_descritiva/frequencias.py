# poetry run frequencias.py

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import unicodedata

# ---- Formatação legível: SEM notação científica ----
np.set_printoptions(suppress=True)
pd.set_option("display.float_format", "{:,.4f}".format)
pd.set_option("display.width", 120)
pd.set_option("display.max_columns", 200)

# 00) Selecionar e carregar Excel
def escolher_arquivo_excel() -> str:
    """
    Abre uma janela para o usuário selecionar um arquivo .xlsx no computador.
    """
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return caminho

def carregar_excel(caminho: str) -> pd.DataFrame:
    """
    Lê o arquivo Excel no DataFrame do pandas.
    """
    return pd.read_excel(caminho)

# Utilitários
def normalizar_nomes_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza nomes de colunas (remove espaços/normaliza Unicode).
    """
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
    """
    Retorna todas as colunas cujo nome começa com 'p_' (tipos de prazer).
    """
    return [c for c in df.columns if isinstance(c, str) and c.startswith("p_")]

def forcar_numerico(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    """
    Converte colunas de prazer para numérico; valores inválidos viram NaN.
    """
    out = df.copy()
    out[colunas] = out[colunas].apply(pd.to_numeric, errors="coerce")
    return out

# ------------------- Limpezas/validações solicitadas -------------------

def remover_duplicatas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove linhas duplicadas completas.
    """
    return df.drop_duplicates()

def remover_linhas_incompletas(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    """
    Remove linhas que não possuem TODOS os dados preenchidos nas colunas de prazer.
    """
    return df.dropna(subset=colunas)

def validar_escala_por_linha(df: pd.DataFrame, colunas: list) -> pd.DataFrame:
    """
    Mantém apenas linhas onde TODAS as colunas de prazer estão em {-3,-2,-1,0,1,2,3}.
    Linhas com qualquer valor fora desse conjunto são excluídas.
    """
    escala_valida = {-3, -2, -1, 0, 1, 2, 3}
    mask_valida = df[colunas].applymap(lambda x: x in escala_valida).all(axis=1)
    return df[mask_valida].copy()

# ----------------------- FUNÇÕES DE RANKING -----------------------

def top10_maiores_medianas(df: pd.DataFrame, colunas: list, n: int = 10) -> pd.Series:
    """
    Calcula a mediana por coluna e retorna as Top-N colunas com MAIORES medianas.
    """
    medianas = df[colunas].median(skipna=True)
    return medianas.sort_values(ascending=False).head(n)

def top10_maiores_medias(df: pd.DataFrame, colunas: list, n: int = 10) -> pd.Series:
    """
    Calcula a média por coluna e retorna as Top-N colunas com MAIORES médias.
    """
    medias = df[colunas].mean(skipna=True)
    return medias.sort_values(ascending=False).head(n)

# ----------------------- IMPRESSÃO -----------------------

def imprimir_serie(titulo: str, serie: pd.Series):
    """
    Imprime uma Series (ranking) de forma amigável no terminal.
    """
    print("\n" + "=" * 100)
    print(titulo)
    print("=" * 100)
    df_print = serie.to_frame(name="score")
    print(df_print.to_string())

# ----------------------- EXECUÇÃO -----------------------

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

    # Garantir tipo numérico
    df = forcar_numerico(df, colunas_prazer)

    # Limpezas na ordem correta: duplicatas -> incompletas -> validação de escala por linha
    linhas_iniciais = len(df)

    df = remover_duplicatas(df)
    apos_dup = len(df)

    df = remover_linhas_incompletas(df, colunas_prazer)
    apos_incompletas = len(df)

    df = validar_escala_por_linha(df, colunas_prazer)
    apos_validacao = len(df)

    print("\n🧹 Limpeza aplicada:")
    print(f"- Linhas iniciais: {linhas_iniciais}")
    print(f"- Removidas por duplicidade: {linhas_iniciais - apos_dup}")
    print(f"- Removidas por dados faltantes: {apos_dup - apos_incompletas}")
    print(f"- Removidas por fora da escala (-3..3): {apos_incompletas - apos_validacao}")
    print(f"- Linhas finais: {apos_validacao}")

    # (1) Top 10 com MAIORES MEDIANAS
    top_medianas = top10_maiores_medianas(df, colunas_prazer, n=10)
    imprimir_serie("Top 10 – Highest medians (by pleasure)", top_medianas)

    # (2) Top 10 com MAIORES MÉDIAS
    top_medias = top10_maiores_medias(df, colunas_prazer, n=10)
    imprimir_serie("Top 10 – Highest means (by pleasure)", top_medias)

    print("\n✅ Finished: rankings printed without scientific notation.")