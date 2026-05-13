# 🧹 Líder Limpe — Painel de Processamento de Planos (v3)

App Streamlit para processar **Domínio**, **Unimed Saúde** (Ambulatorial / Santas / Interior / Metropolitano), **Unimed Odonto** (TXT) e **VSP SAMP** (Sindseg + Sindivigilantes).

> ⚠️ **v3**: módulo de Dependentes foi removido conforme solicitação. Foco apenas em planos.

## 🚀 Como usar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Deploy no Streamlit Cloud

1. Faça push para o GitHub.
2. Em [share.streamlit.io](https://share.streamlit.io/), aponte para `app.py`.
3. Deploy automático.

> ✅ `requirements.txt` fixa `pandas>=2.2,<3.0` para evitar bugs de compatibilidade.

## 📦 Estrutura

```
liderlimpe-app/
├── app.py                  # Interface Streamlit
├── requirements.txt        # pandas>=2.2 fixado
├── README.md
├── .gitignore
├── .streamlit/config.toml
├── assets/                 # (opcional) coloque logo.png aqui
└── modules/
    ├── empresas.py         # codi_emp 1=VSP, 2=ATIVA, 3=MULTI, 4=COMERCIAL
    ├── utils.py            # leitura xls/xlsx/txt, limpa CPF/valor
    ├── dominio.py          # Domínio (sem groupby.apply)
    ├── saude.py            # 4 contratos Unimed Saúde
    ├── odonto.py           # Parser TXT Unimed Odonto
    ├── vsp_samp.py         # SAMP Sindseg + Sindivigilantes
    ├── ccusto.py           # Planilha RELAÇÃO CCUSTO
    └── orquestrador.py     # Cola tudo + ZIPs
```

## ✅ Bugs corrigidos na v3

| Bug | Solução |
|---|---|
| `include_groups=True is no longer allowed` | Reescrevi `_tratar_duplicados_cpf` em `dominio.py` usando `sort_values + drop_duplicates`, sem nenhum `groupby().apply()`. |
| `ValueError: too many values to unpack` | Corrigi o desempacotamento `(emp, label, cor)` em todas as ocorrências do `app.py`. |
| Dependentes dando erro | Removido completamente (módulo + chamadas). |

## 🎯 Regras de negócio

- **Empresa** por contrato (VSP 6217/5964) ou `codi_emp` da Domínio.
- **Domínio**: aceita .xls e .xlsx. Demitidos = `situacao==8`. CPFs duplicados: mantém ativo > demitido; entre ativos, maior `i_empregados`.
- **Saúde**: PROCV `CPFTITULAR × Domínio.cpf` → `nome_quebra` como CCusto; soma `VLFATURADO`.
- **Odonto Unimed (TXT)**: parser `#`, soma "Mensalidade" por CCusto.
- **VSP SAMP**: identifica SINDSEG / SINDIVIGILANTES por nome + pasta; tipo por valor majoritário (15/18,50/10).
- **Saída**: `RESULTADOS_AAAA-MM-DD_HHMM.zip` com `ATIVA.zip`, `COMERCIAL.zip`, `MULTISSERVICOS.zip`, `VSP.zip` + `RELAÇÃO CCUSTO - UNIMED.xlsx`.
