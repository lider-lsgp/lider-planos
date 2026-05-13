"""
Utilitários gerais: leitura de planilhas (xls/xlsx/txt), conversão,
tratamento de CPF, valores monetários e nomes.
"""

import io
import re
import unicodedata
from typing import Optional, Tuple

import pandas as pd
import xlrd


# ---------- Leitura de arquivos ----------

def ler_excel_qualquer(file_like, header=0, sheet=0) -> pd.DataFrame:
    """
    Lê .xls ou .xlsx independentemente da extensão real.
    Aceita caminho, BytesIO ou file-like.
    """
    # Garantir bytes
    if hasattr(file_like, "read"):
        data = file_like.read()
        if hasattr(file_like, "seek"):
            file_like.seek(0)
    else:
        with open(file_like, "rb") as f:
            data = f.read()

    # Tenta xlrd (xls antigo + BIFF2-5)
    try:
        wb = xlrd.open_workbook(file_contents=data, logfile=io.StringIO())
        sh = wb.sheet_by_index(sheet) if isinstance(sheet, int) else wb.sheet_by_name(sheet)
        rows = []
        for r in range(sh.nrows):
            rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
        if not rows:
            return pd.DataFrame()
        if header is None:
            return pd.DataFrame(rows)
        head = rows[header]
        body = rows[header + 1:]
        df = pd.DataFrame(body, columns=head)
        return df
    except Exception:
        pass

    # Tenta openpyxl (xlsx moderno)
    try:
        bio = io.BytesIO(data)
        return pd.read_excel(bio, header=header, sheet_name=sheet, engine="openpyxl")
    except Exception:
        pass

    # Última tentativa: pandas auto
    bio = io.BytesIO(data)
    return pd.read_excel(bio, header=header, sheet_name=sheet)


def ler_txt_odonto(file_like, encoding="latin-1") -> Tuple[list, str, str, str]:
    """
    Lê o TXT do demonstrativo Odonto da Unimed.

    Retorna (linhas_tabela, empresa_completa, cnpj, fatura).
    - As 7 primeiras linhas e as 3 últimas são fora da tabela.
    - O header verdadeiro é a linha 7 (índice).
    """
    if hasattr(file_like, "read"):
        data = file_like.read()
        if hasattr(file_like, "seek"):
            file_like.seek(0)
        text = data.decode(encoding, errors="replace") if isinstance(data, (bytes, bytearray)) else data
    else:
        with open(file_like, "r", encoding=encoding, errors="replace") as f:
            text = f.read()

    linhas = text.splitlines()

    # Identifica metadados (CNPJ, empresa, fatura) nas primeiras linhas
    empresa_nome, cnpj, fatura = "", "", ""
    for i in range(min(10, len(linhas))):
        L = linhas[i]
        if "Empresa:" in L:
            # ex: "Empresa:  002.201.230/0001-44 - ATIVA TERCEIRIZACAO DE MAO DE OBRA LTDA"
            after = L.split("Empresa:", 1)[1].strip()
            m = re.match(r"([\d\.\/\-]+)\s*-\s*(.+)", after)
            if m:
                cnpj = re.sub(r"\D", "", m.group(1))
                empresa_nome = m.group(2).strip()
        if "Fatura:" in L:
            m = re.search(r"Fatura:\s*([\w\d\-]+)", L)
            if m:
                fatura = m.group(1).strip()

    return linhas, empresa_nome, cnpj, fatura


# ---------- Tratamentos ----------

def limpar_cpf(cpf) -> str:
    """Remove tudo que não é dígito e pad com zeros à esquerda até 11."""
    if cpf is None:
        return ""
    s = re.sub(r"\D", "", str(cpf))
    if not s:
        return ""
    # Trunca se passar de 11 (erro de digitação) e pad até 11
    s = s[-11:] if len(s) > 11 else s
    return s.zfill(11)


def limpar_nome(nome) -> str:
    """Remove espaços extras (início, fim e duplicados internos)."""
    if nome is None:
        return ""
    s = str(nome)
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_valor_br(v) -> float:
    """
    Converte valor BR ('1.234,56', 'R$ 10,00', '-8,73', '187', 187.5)
    para float. Retorna 0.0 quando inválido.
    """
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.0
    s = str(v).strip()
    if not s:
        return 0.0
    # remove R$, espaços
    s = s.replace("R$", "").replace("\xa0", "").strip()
    # remove separador de milhar BR
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def normaliza_chave(s: str) -> str:
    """Normaliza string para chave (sem acento, maiúscula, sem espaço extra)."""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s


def parse_nome_quebra(quebra) -> Tuple[int, str]:
    """
    'nome_quebra' costuma vir como '12 - ALGUM POSTO'.
    Retorna (codigo_int, nome_string). Se não tiver código, retorna (0, raw).
    """
    if quebra is None:
        return (0, "")
    s = limpar_nome(quebra)
    if not s:
        return (0, "")
    m = re.match(r"^(\d+)\s*-\s*(.+)$", s)
    if m:
        try:
            return (int(m.group(1)), m.group(2).strip())
        except ValueError:
            return (0, s)
    return (0, s)


def codigo_nome_quebra(quebra) -> int:
    """Retorna apenas o código numérico do nome_quebra (para ordenação)."""
    return parse_nome_quebra(quebra)[0]


# ---------- Conversão para xlsx ----------

def df_para_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """Converte DataFrame em bytes de um arquivo XLSX."""
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return bio.getvalue()


def sanitize_filename(name: str) -> str:
    """Limpa nome para usar como arquivo."""
    s = str(name)
    s = re.sub(r"[\\/:*?\"<>|]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s
