"""
Identificação e normalização de empresas do grupo Líder.

Regras (conforme regra de negócio do usuário):
- VSP é identificada pelo NRCONTRATO (6217 = Interior, 5964 = Metropolitano).
  Também por CNPJ ou nome se aparecer no Domínio.
- Demais empresas são identificadas pela LOTAÇÃO (primeiros 4 dígitos):
    * 0001 -> ATIVA
    * 0002 -> L. COMERCIAL  (LIDER LIMPE LIMPEZA COMERCIAL)
    * 0003 -> L. MULTISSERVIÇOS
- Fallback: codi_emp da planilha Domínio.
"""

import re

# Códigos da Domínio (codi_emp) -> chave canônica de empresa
# Mapeamento REAL verificado nos arquivos:
#   codi_emp = 1 -> VSP VIGILANCIA E SEGURANCA PATRIMONIAL
#   codi_emp = 2 -> ATIVA TERCEIRIZACAO
#   codi_emp = 3 -> LIDER MULTISSERVICOS
#   codi_emp = 4 -> LIDER LIMPE LIMPEZA COMERCIAL
CODI_EMP_MAP = {
    1: "VSP",
    2: "ATIVA",
    3: "MULTISSERVICOS",
    4: "COMERCIAL",
}

# Lotação (4 dígitos) -> empresa
LOTACAO_MAP = {
    "0001": "ATIVA",
    "0002": "COMERCIAL",
    "0003": "MULTISSERVICOS",
}

# Contratos Unimed -> empresa + tipo de plano
CONTRATOS_UNIMED = {
    "5957": ("ATIVA_LIDER", "SAUDE_AMBULATORIAL"),
    "6040": ("ATIVA_LIDER", "SAUDE_SANTAS"),
    "6217": ("VSP", "SAUDE_INTERIOR"),
    "5964": ("VSP", "SAUDE_METROPOLITANO"),
}

# Mapa canônico -> nome curto (para nome de arquivo) e nome completo
EMPRESAS = {
    "ATIVA": {
        "curto": "ATIVA",
        "completo": "ATIVA TERCEIRIZACAO DE MAO DE OBRA LTDA",
        "cnpj_prefixo": "02201230",
        "tem_odonto_unimed": True,
        "tem_santas": True,
    },
    "COMERCIAL": {
        "curto": "COMERCIAL",
        "completo": "LIDER LIMPE LIMPEZA COMERCIAL LTDA",
        "cnpj_prefixo": "03659631",
        "tem_odonto_unimed": True,
        "tem_santas": True,
    },
    "MULTISSERVICOS": {
        "curto": "MULTISSERVICOS",
        "completo": "LIDER MULTISSERVICOS LTDA",
        "cnpj_prefixo": "",
        "tem_odonto_unimed": True,
        "tem_santas": False,
    },
    "VSP": {
        "curto": "VSP",
        "completo": "VSP VIGILANCIA E SEGURANCA PATRIMONIAL LTDA",
        "cnpj_prefixo": "",
        "tem_odonto_unimed": False,  # VSP usa SAMP (Sindseg/Sindivigilantes)
        "tem_santas": False,
    },
}


def empresa_por_contrato(nrcontrato: str) -> str | None:
    """Retorna a empresa canônica a partir do número de contrato da Unimed."""
    if nrcontrato is None:
        return None
    s = str(nrcontrato).strip()
    if s in ("6217", "5964"):
        return "VSP"
    if s in ("5957", "6040"):
        # Esses contratos atendem ATIVA, COMERCIAL e MULTISSERVICOS;
        # não definem empresa sozinhos -> precisará da lotação ou codi_emp.
        return None
    return None


def empresa_por_lotacao(lotacao) -> str | None:
    """Identifica empresa a partir da lotação (4 primeiros dígitos)."""
    if lotacao is None:
        return None
    s = str(lotacao).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    # Remove pontos/traços/zeros à esquerda só para casar prefixo
    digitos = re.sub(r"\D", "", s)
    if not digitos:
        return None
    prefix = digitos[:4].zfill(4)
    return LOTACAO_MAP.get(prefix)


def empresa_por_codi_emp(codi_emp) -> str | None:
    """Identifica empresa via codi_emp da planilha Domínio."""
    if codi_emp is None:
        return None
    try:
        c = int(float(str(codi_emp).strip()))
    except (ValueError, TypeError):
        return None
    return CODI_EMP_MAP.get(c)


def empresa_por_nome(nome) -> str | None:
    """Fallback: tenta identificar empresa pelo nome completo (texto)."""
    if not nome:
        return None
    n = str(nome).upper()
    if "VSP" in n or "VIGILANCIA" in n or "VIGILÂNCIA" in n:
        return "VSP"
    if "MULTISSERVI" in n or "MULTI " in n or n.endswith("MULTI"):
        return "MULTISSERVICOS"
    if "COMERCIAL" in n:
        return "COMERCIAL"
    if "ATIVA" in n:
        return "ATIVA"
    return None


def empresa_por_cnpj(cnpj) -> str | None:
    """Identifica empresa pelo CNPJ (string com ou sem máscara)."""
    if not cnpj:
        return None
    digitos = re.sub(r"\D", "", str(cnpj))
    if not digitos:
        return None
    for chave, info in EMPRESAS.items():
        pref = info.get("cnpj_prefixo", "")
        if pref and digitos.startswith(pref):
            return chave
    return None


def nome_curto(empresa_chave: str) -> str:
    """Retorna o nome curto da empresa (para nome de arquivo)."""
    if empresa_chave in EMPRESAS:
        return EMPRESAS[empresa_chave]["curto"]
    return empresa_chave or "DESCONHECIDA"
