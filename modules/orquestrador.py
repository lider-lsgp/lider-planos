"""
Orquestrador: classifica arquivos enviados, agrupa por empresa,
e dispara o processamento de cada módulo.

Aceita:
  - ZIPs (1 ou vários) com subpastas por empresa.
  - Arquivos soltos (xls/xlsx/txt) — detecta pela extensão + conteúdo.

Saída:
  Um ZIP "RESULTADOS.zip" contendo:
    - ATIVA.zip
    - COMERCIAL.zip
    - MULTISSERVICOS.zip
    - VSP.zip
    - RELAÇÃO CCUSTO - UNIMED.xlsx (consolidado)
    - DEPENDENTES SAÚDE.xlsx (consolidado)
    - DEPENDENTES ODONTO.xlsx (consolidado)
"""

import io
import os
import re
import zipfile
from typing import Optional

import pandas as pd

from .empresas import (
    EMPRESAS,
    empresa_por_codi_emp,
    empresa_por_contrato,
    empresa_por_lotacao,
    empresa_por_nome,
    empresa_por_cnpj,
)
from .dominio import carregar_dominio, nome_colaboradores_arquivo
from .saude import (
    carregar_relatorio_saude,
    montar_relatorio_saude,
    gerar_xlsx_saude,
    nome_arquivo_relatorio,
)
from .odonto import (
    carregar_relatorio_odonto,
    montar_relatorio_odonto,
    gerar_xlsx_odonto,
    nome_arquivo_odonto,
)
from .vsp_samp import (
    carregar_boleto_samp,
    consolidar_samp,
    gerar_xlsx_samp_boleto,
    nome_arquivo_samp,
    identificar_sindicato,
)
from .ccusto import gerar_relacao_ccusto, ABA_POR_EMPRESA
from .utils import ler_excel_qualquer


# ---------- Classificação de arquivos ----------

def classificar_arquivo(nome: str, pasta_pai: str = "") -> str:
    """
    Retorna um rótulo:
      'DOMINIO', 'SAUDE', 'ODONTO_TXT', 'SAMP_FUNC', 'SAMP_DEP',
      'DEPENDENTES_SAUDE_MODELO', 'DEPENDENTES_ODONTO_MODELO',
      'CCUSTO_MODELO', 'DESCONHECIDO'.

    Ordem de checagem (do mais específico para o mais genérico).
    """
    base = nome.upper()
    pasta = pasta_pai.upper()
    base_pasta = base + " " + pasta

    # 1) Modelos (antes de qualquer coisa)
    if "MODELO" in base:
        if "DEPENDENTES" in base and ("SAÚDE" in base or "SAUDE" in base):
            return "DEPENDENTES_SAUDE_MODELO"
        if "DEPENDENTES" in base and "ODONTO" in base:
            return "DEPENDENTES_ODONTO_MODELO"
        if "CCUSTO" in base:
            return "CCUSTO_MODELO"
        return "DESCONHECIDO"

    # 2) ODONTO TXT (único .txt do sistema Unimed)
    if base.endswith(".TXT"):
        return "ODONTO_TXT"

    # 3) VSP SAMP
    if base.endswith(".XLSX"):
        if "CONTROLE DE PAGAMENTOS" in base:
            return "SAMP_DEP"
        if "CALCULO RELA" in base or "SINDSEG" in base_pasta or "SINDIVIGILANTES" in base_pasta:
            return "SAMP_FUNC"

    # 4) DOMÍNIO (só quando explícito - tem "DOMINIO" no nome)
    # Importante: precisa vir ANTES de SAUDE porque alguns nomes podem ter "GERAL"
    if base.endswith(".XLS") or base.endswith(".XLSX"):
        if "DOMINIO" in base or "DOMÍNIO" in base:
            return "DOMINIO"
        # Sem palavra "DOMINIO": só considera Domínio se o nome for tipo
        # "ATIVA GERAL.xls", "GERAL VSP.xls" sozinho (não tem "RELAÇÃO" nem "SAUDE")
        if "RELAÇÃO" not in base and "RELACAO" not in base and "EXTRATO" not in base \
           and "SAUDE" not in base and "SAÚDE" not in base and "ODONTO" not in base:
            if "GERAL" in base or base.startswith("ATIVA ") or " ATIVA" in base:
                return "DOMINIO"

        # 5) SAÚDE Unimed (relatórios)
    if base.endswith(".XLS") or base.endswith(".XLSX"):
        if "ODONTO" not in base:
            # Palavras-chave que indicam relatório de saúde da Unimed
            palavras_saude = (
                "SAUDE", "SAÚDE", "EXTRATO", "RELAÇÃO", "RELACAO",
                "RELATORIO", "RELATÓRIO", "RELATIVO",   # typos comuns
                "AMBULATORIAL", "SANTAS", "INTERIOR", "METROPOLITANO",
            )
            if any(p in base for p in palavras_saude):
                return "SAUDE"

    return "DESCONHECIDO"



def detectar_empresa_arquivo(nome: str, pasta_pai: str = "") -> Optional[str]:
    """Tenta detectar empresa pelo nome do arquivo ou pasta-pai."""
    e = empresa_por_nome(nome)
    if e:
        return e
    e = empresa_por_nome(pasta_pai)
    if e:
        return e
    return None


# ---------- Estrutura para acumular tudo ----------

def estrutura_vazia() -> dict:
    return {
        emp: {
            "dominio": None,         # dict do carregar_dominio
            "saudes": [],            # list de relatorios de saude crus (dict)
            "odonto_txt": None,      # dict do carregar_relatorio_odonto
            "samp_boletos": [],      # list de carregar_boleto_samp
            "samp_dep": [],          # list dos 'Controle de Pagamentos'
            "saude_raw_dfs": [],     # para extração de dependentes
        }
        for emp in ("ATIVA", "COMERCIAL", "MULTISSERVICOS", "VSP")
    }


# ---------- Processamento principal ----------

def processar_uploads(uploads: list, modelos: Optional[dict] = None) -> dict:
    """
    uploads: lista de tuplas (nome_arquivo, conteudo_bytes, pasta_pai_opcional)
    modelos: dict opcional com:
        {'dep_saude': df_modelo_saude, 'dep_odonto': df_modelo_odonto}

    Retorna dict com bytes finais e resumo:
      {
        'zip_final': bytes,
        'arquivos': {nome: bytes},
        'log': [str, ...],
        'resumo': {...}
      }
    """
    log = []
    estrutura = estrutura_vazia()
    modelo_dep_saude = (modelos or {}).get("dep_saude")
    modelo_dep_odonto = (modelos or {}).get("dep_odonto")

    # Primeiro passo: expandir ZIPs em arquivos
    arquivos_planos = []  # (nome, bytes, pasta_pai)
    for nome, dados, pasta_pai in uploads:
        if nome.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(dados)) as zf:
                    for zi in zf.infolist():
                        if zi.is_dir():
                            continue
                        nome_int = os.path.basename(zi.filename)
                        pasta_int = os.path.dirname(zi.filename)
                        arquivos_planos.append((nome_int, zf.read(zi), pasta_int))
            except zipfile.BadZipFile:
                log.append(f"⚠️ Arquivo '{nome}' não é um ZIP válido.")
        else:
            arquivos_planos.append((nome, dados, pasta_pai))

    # Classifica e direciona cada arquivo
    for nome, dados, pasta in arquivos_planos:
        cat = classificar_arquivo(nome, pasta)
        empresa = detectar_empresa_arquivo(nome, pasta)

        try:
            if cat == "DOMINIO":
                info = carregar_dominio(io.BytesIO(dados), nome_arquivo=nome)
                # Empresa do Domínio vence (codi_emp). Só cai em hint/pasta se
                # o codi_emp não foi reconhecido.
                emp = info["empresa"]
                if emp not in estrutura:
                    emp = empresa or _empresa_por_pasta_arquivo(pasta, nome) or "ATIVA"
                # Se já existir Domínio nessa empresa, mantém o primeiro
                # (xls e xlsx duplicados normalmente são o mesmo conteúdo)
                if estrutura[emp]["dominio"] is None:
                    estrutura[emp]["dominio"] = info
                    log.append(f"✅ Domínio: {nome} → {emp} ({info['qtd_total']} colaboradores)")
                else:
                    log.append(f"ℹ️ Domínio duplicado ignorado: {nome} → {emp}")

            elif cat == "SAUDE":
                info = carregar_relatorio_saude(io.BytesIO(dados), nome_arquivo=nome)
                contrato = info["contrato"]
                # Empresa do contrato (só VSP por contrato)
                emp_por_contrato = empresa_por_contrato(contrato)
                emp = emp_por_contrato or empresa or _empresa_por_pasta_arquivo(pasta, nome)
                if emp not in estrutura:
                    log.append(f"⚠️ Saúde: '{nome}' sem empresa detectada (contrato {contrato}).")
                    continue
                estrutura[emp]["saudes"].append(info)
                estrutura[emp]["saude_raw_dfs"].append(info["df"])
                log.append(f"✅ Saúde: {nome} → {emp} ({info['tipo_curto']}, contrato {contrato})")

            elif cat == "ODONTO_TXT":
                info = carregar_relatorio_odonto(io.BytesIO(dados), nome_arquivo=nome)
                emp = info["empresa"] or empresa or _empresa_por_pasta_arquivo(pasta, nome)
                if emp not in estrutura:
                    log.append(f"⚠️ Odonto: '{nome}' sem empresa detectada.")
                    continue
                estrutura[emp]["odonto_txt"] = info
                log.append(f"✅ Odonto: {nome} → {emp} (fatura {info.get('fatura','')})")

            elif cat == "SAMP_FUNC":
                info = carregar_boleto_samp(io.BytesIO(dados), nome_arquivo=nome, pasta=pasta)
                estrutura["VSP"]["samp_boletos"].append(info)
                log.append(f"✅ VSP SAMP: {nome} → {info['sindicato']}/{info['tipo']}")

            elif cat == "SAMP_DEP":
                info = carregar_boleto_samp(io.BytesIO(dados), nome_arquivo=nome, pasta=pasta)
                # Marcar como Odonto - Dependentes
                info["tipo"] = "Odonto - Dependentes"
                estrutura["VSP"]["samp_dep"].append(info)
                estrutura["VSP"]["samp_boletos"].append(info)
                log.append(f"✅ VSP SAMP DEP: {nome} → {info['sindicato']}")

            elif cat in ("DEPENDENTES_SAUDE_MODELO", "DEPENDENTES_ODONTO_MODELO", "CCUSTO_MODELO"):
                log.append(f"📋 Modelo recebido: {nome}")

            else:
                log.append(f"❓ Não classificado: {nome} (pasta {pasta or '/'})")
        except Exception as e:
            log.append(f"❌ Erro processando '{nome}': {e}")

    # ---------- Geração de saídas ----------
    arquivos = {}
    dados_ccusto = {}

    for emp, info in estrutura.items():
        if not info["dominio"] and not info["saudes"] and not info["odonto_txt"] and not info["samp_boletos"]:
            continue

        emp_dir = {}  # arquivos da empresa

        # 1. Planilha de colaboradores tratada (Domínio)
        if info["dominio"]:
            dom_df = info["dominio"]["df"]
            import io as _io
            from openpyxl import Workbook
            bio = _io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as w:
                dom_df.to_excel(w, index=False, sheet_name="Colaboradores")
            emp_dir[nome_colaboradores_arquivo(emp)] = bio.getvalue()
        else:
            dom_df = pd.DataFrame(columns=["cpf", "nome_quebra"])

        # 2. Relatórios de Saúde (1 por contrato)
        for s in info["saudes"]:
            payload = montar_relatorio_saude(s, dom_df, emp)
            fname = nome_arquivo_relatorio(emp, payload["tipo_curto"], payload["valor_total"])
            emp_dir[fname] = gerar_xlsx_saude(payload)

            # Dados para CCusto consolidado
            dados_ccusto.setdefault(emp, {})
            if s["contrato"] == "5957":
                dados_ccusto[emp]["SAUDE_AMBULATORIAL"] = payload
            elif s["contrato"] == "6040":
                dados_ccusto[emp]["SAUDE_SANTAS"] = payload
            elif s["contrato"] == "6217":
                dados_ccusto[emp]["SAUDE_INTERIOR"] = payload
            elif s["contrato"] == "5964":
                dados_ccusto[emp]["SAUDE_METROPOLITANO"] = payload

        # 3. Relatório Odonto Unimed (TXT)
        if info["odonto_txt"]:
            payload = montar_relatorio_odonto(info["odonto_txt"], dom_df)
            payload["empresa"] = emp
            fname = nome_arquivo_odonto(emp, payload["valor_total"])
            emp_dir[fname] = gerar_xlsx_odonto(payload)
            dados_ccusto.setdefault(emp, {})
            dados_ccusto[emp]["ODONTO"] = payload

                # 4. VSP SAMP
        if emp == "VSP" and info["samp_boletos"]:
            # Antes de consolidar: tentar resolver sindicato dos SAMP_DEP
            # cruzando os CPFs dos funcionários com os boletos já identificados
            _resolver_sindicato_dep(info["samp_boletos"], info["samp_dep"], log)
            samp_res = consolidar_samp(info["samp_boletos"], dom_df)

            for b in samp_res["boletos"]:
                fname = nome_arquivo_samp(b["sindicato"], b["tipo"], b["valor_total"])
                emp_dir[fname] = gerar_xlsx_samp_boleto(b)
            dados_ccusto.setdefault("VSP", {})
            dados_ccusto["VSP"]["ODONTO_SAMP"] = samp_res

        # Empacotar empresa em sub-ZIP
        sub_bio = io.BytesIO()
        with zipfile.ZipFile(sub_bio, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, data in emp_dir.items():
                zf.writestr(fname, data)
        arquivos[f"{emp}.zip"] = sub_bio.getvalue()

    # ---------- CCUSTO consolidado ----------
    if dados_ccusto:
        ccusto_bytes = gerar_relacao_ccusto(dados_ccusto)
        arquivos["RELAÇÃO CCUSTO - UNIMED.xlsx"] = ccusto_bytes

    # (Parte de Dependentes desativada por solicitação do usuário - v3)

    # ---------- ZIP geral ----------
    final_bio = io.BytesIO()
    with zipfile.ZipFile(final_bio, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in arquivos.items():
            zf.writestr(fname, data)

    return {
        "zip_final": final_bio.getvalue(),
        "arquivos": arquivos,
        "log": log,
        "resumo": _gerar_resumo(estrutura, dados_ccusto),
    }


def _empresa_por_pasta_arquivo(pasta: str, nome: str) -> Optional[str]:
    """Detecta empresa pela pasta-pai (caso seja claro: 'ATIVA/saude.xls')."""
    return empresa_por_nome(pasta) or empresa_por_nome(nome)

def _resolver_sindicato_dep(todos_boletos: list, boletos_dep: list, log: list) -> None:
    """
    Para cada 'Controle de Pagamentos' sem sindicato definido, cruza os CPFs
    dos funcionários titulares com os boletos SAMP_FUNC já identificados.
    """
    mapa = {}
    for b in todos_boletos:
        sind = b.get("sindicato")
        if not sind or sind == "DESCONHECIDO":
            continue
        df = b.get("df")
        if df is None or df.empty:
            continue
        cpf_col = None
        for c in ("CPF", "CPF Funcionário", "CPF Funcionario"):
            if c in df.columns:
                cpf_col = c
                break
        if not cpf_col:
            continue
        for cpf in df[cpf_col].dropna().astype(str):
            if cpf.strip():
                mapa[cpf.strip()] = sind

    if not mapa:
        return

    for b in boletos_dep:
        if b.get("sindicato") and b["sindicato"] != "DESCONHECIDO":
            continue
        df = b.get("df")
        if df is None or df.empty:
            continue
        cpf_col = None
        for c in ("CPF Funcionário", "CPF Funcionario", "CPF"):
            if c in df.columns:
                cpf_col = c
                break
        if not cpf_col:
            continue
        contagem = {}
        for cpf in df[cpf_col].dropna().astype(str):
            cpf = cpf.strip()
            if cpf in mapa:
                contagem[mapa[cpf]] = contagem.get(mapa[cpf], 0) + 1
        if contagem:
            sind_majoritario = max(contagem.items(), key=lambda x: x[1])[0]
            b["sindicato"] = sind_majoritario
            log.append(f"🔍 SAMP DEP resolvido por CPF → {sind_majoritario} ({contagem[sind_majoritario]} matches)")


def _gerar_resumo(estrutura: dict, dados_ccusto: dict) -> dict:
    """Resumo amigável para mostrar no app."""
    out = {}
    for emp, info in estrutura.items():
        emp_resumo = {
            "dominio_qtd": info["dominio"]["qtd_total"] if info["dominio"] else 0,
            "saude_relatorios": len(info["saudes"]),
            "odonto": bool(info["odonto_txt"]),
            "samp_boletos": len(info["samp_boletos"]),
            "valores": {},
        }
        cc = dados_ccusto.get(emp, {})
        for key in ("SAUDE_AMBULATORIAL", "SAUDE_SANTAS", "SAUDE_INTERIOR",
                    "SAUDE_METROPOLITANO", "ODONTO"):
            if key in cc:
                emp_resumo["valores"][key] = cc[key]["valor_total"]
        if "ODONTO_SAMP" in cc:
            total_samp = sum(b["valor_total"] for b in cc["ODONTO_SAMP"]["boletos"])
            emp_resumo["valores"]["VSP_ODONTO_SAMP"] = total_samp
        out[emp] = emp_resumo
    return out
