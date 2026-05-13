"""
Tratamento da planilha Domínio (relação de empregados).

VERSÃO 3 - sem dependência de groupby().apply()
"""

import io
import pandas as pd
import numpy as np

from .utils import (
    ler_excel_qualquer,
    limpar_cpf,
    limpar_nome,
)
from .empresas import (
    empresa_por_codi_emp,
    empresa_por_nome,
    EMPRESAS,
)


COLUNAS_DESEJADAS = [
    "i_empregados",
    "nome",
    "cpf",
    "quebra",
    "nome_quebra",
    "admissao",
    "datasituacao",
    "situacao",
    "codi_emp",
    "cp_nome_emp",
]


def carregar_dominio(file_like, nome_arquivo: str = "") -> dict:
    """
    Lê e trata a planilha Domínio.
    """
    df = ler_excel_qualquer(file_like, header=0, sheet=0)

    df.columns = [str(c).strip().lower() for c in df.columns]

    cols_existentes = [c for c in COLUNAS_DESEJADAS if c in df.columns]
    if not cols_existentes:
        raise ValueError("Planilha Domínio sem colunas reconhecidas.")
    df = df[cols_existentes].copy()

    # Limpeza de tipos
    if "nome" in df.columns:
        df["nome"] = df["nome"].map(limpar_nome)
    if "cpf" in df.columns:
        df["cpf"] = df["cpf"].map(limpar_cpf)
    if "nome_quebra" in df.columns:
        df["nome_quebra"] = df["nome_quebra"].map(limpar_nome)

    if "situacao" in df.columns:
        df["situacao"] = pd.to_numeric(df["situacao"], errors="coerce")

    if "datasituacao" in df.columns:
        df["datasituacao"] = pd.to_datetime(
            df["datasituacao"], errors="coerce", dayfirst=True
        )

    if "admissao" in df.columns:
        df["admissao"] = pd.to_datetime(df["admissao"], errors="coerce", dayfirst=True)

    # Regra: quem não está demitido (situacao != 8) -> limpar datasituacao
    if "situacao" in df.columns and "datasituacao" in df.columns:
        df.loc[df["situacao"] != 8, "datasituacao"] = pd.NaT

    # Tratamento de duplicados por CPF (SEM groupby.apply)
    if "cpf" in df.columns:
        df = _tratar_duplicados_cpf(df)

    if "nome" in df.columns:
        df = df.sort_values("nome", kind="stable").reset_index(drop=True)

    # Identificar empresa
    empresa_chave = None
    empresa_nome_full = ""
    if "codi_emp" in df.columns and not df.empty:
        amostra = df["codi_emp"].dropna()
        if len(amostra) > 0:
            empresa_chave = empresa_por_codi_emp(amostra.iloc[0])
    if not empresa_chave and "cp_nome_emp" in df.columns and not df.empty:
        amostra_nome = df["cp_nome_emp"].dropna()
        if len(amostra_nome) > 0:
            an = amostra_nome.iloc[0]
            empresa_chave = empresa_por_nome(an)
            if an:
                empresa_nome_full = str(an)
    if not empresa_chave and nome_arquivo:
        empresa_chave = empresa_por_nome(nome_arquivo)

    if empresa_chave and empresa_chave in EMPRESAS and not empresa_nome_full:
        empresa_nome_full = EMPRESAS[empresa_chave]["completo"]

    qtd_total = len(df)
    qtd_demitidos = int((df["situacao"] == 8).sum()) if "situacao" in df.columns else 0
    qtd_ativos = qtd_total - qtd_demitidos

    return {
        "df": df,
        "empresa": empresa_chave or "DESCONHECIDA",
        "empresa_nome_completo": empresa_nome_full,
        "qtd_total": qtd_total,
        "qtd_ativos": qtd_ativos,
        "qtd_demitidos": qtd_demitidos,
        "arquivo": nome_arquivo,
    }


def _tratar_duplicados_cpf(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove duplicados por CPF SEM usar groupby.apply.

    Estratégia:
      - Cria colunas auxiliares de prioridade
      - Ordena tudo de uma vez
      - drop_duplicates(keep='first')

    Prioridade (do mais "guardável" pro menos):
      1. Ativos primeiro (situacao != 8)
      2. Maior i_empregados (cadastro mais novo)
      3. datasituacao mais recente (para os demitidos)
    """
    if "cpf" not in df.columns or df.empty:
        return df

    df = df.copy().reset_index(drop=True)

    # Separa vazios (não dá pra deduplicar)
    mask_vazios = (df["cpf"] == "") | df["cpf"].isna()
    df_vazios = df[mask_vazios].copy()
    df_validos = df[~mask_vazios].copy()

    if df_validos.empty:
        return df_vazios.reset_index(drop=True)

    # __eh_demitido: 0 = ativo, 1 = demitido (ativos vêm primeiro)
    if "situacao" in df_validos.columns:
        df_validos["__eh_demitido"] = np.where(df_validos["situacao"] == 8, 1, 0)
    else:
        df_validos["__eh_demitido"] = 0

    # __i_emp: maior é "mais novo" -> invertemos com negativo
    if "i_empregados" in df_validos.columns:
        df_validos["__i_emp_neg"] = -pd.to_numeric(
            df_validos["i_empregados"], errors="coerce"
        ).fillna(0)
    else:
        df_validos["__i_emp_neg"] = 0

    # __dts_neg: data mais recente vem primeiro (timestamp negativo)
    if "datasituacao" in df_validos.columns:
        dts = pd.to_datetime(df_validos["datasituacao"], errors="coerce")
        # converte para timestamp numérico, NaT vira 0
        df_validos["__dts_neg"] = -dts.astype("int64", errors="ignore").fillna(0) \
            if hasattr(dts, "astype") else 0
        # fallback simples
        df_validos["__dts_neg"] = dts.map(
            lambda x: -x.timestamp() if pd.notna(x) else 0.0
        )
    else:
        df_validos["__dts_neg"] = 0.0

    # Ordena (ascending=True em todas): demitidos por último, i_empregados maior primeiro,
    # data mais recente primeiro (porque negativo)
    df_validos = df_validos.sort_values(
        by=["__eh_demitido", "__i_emp_neg", "__dts_neg"],
        ascending=[True, True, True],
        kind="stable",
    )

    # Mantém o primeiro de cada CPF
    df_validos = df_validos.drop_duplicates(subset="cpf", keep="first")

    # Remove colunas auxiliares
    for c in ("__eh_demitido", "__i_emp_neg", "__dts_neg"):
        if c in df_validos.columns:
            df_validos = df_validos.drop(columns=c)

    df_validos = df_validos.reset_index(drop=True)
    df_vazios = df_vazios.reset_index(drop=True)

    return pd.concat([df_validos, df_vazios], ignore_index=True)


def nome_colaboradores_arquivo(empresa_chave: str) -> str:
    curto = EMPRESAS.get(empresa_chave, {}).get("curto", empresa_chave or "EMPRESA")
    return f"COLABORADORES - {curto}.xlsx"
