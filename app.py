"""
Líder Limpe - Painel de Processamento de Planos
=================================================
v3 - Sem módulo de Dependentes. Foco: Domínio + Saúde + Odonto + VSP SAMP.
"""

import io
import os
import base64
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from modules.orquestrador import processar_uploads
from modules.empresas import EMPRESAS

# ============================================================
# VERSÃO DO APP (sempre visível para confirmar deploy)
# ============================================================
APP_VERSION = "v3.0.0"
APP_BUILD = "2026-05-13"

# ============================================================
# Configuração da página
# ============================================================
st.set_page_config(
    page_title="Líder Limpe • Painel de Planos",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# Estilo (CSS)
# ============================================================
PRIMARY = "#1F3864"
ACCENT = "#E8731A"
SUCCESS = "#2E7D32"
DANGER = "#C62828"
WARN = "#ED6C02"
LIGHT_BG = "#F7F9FC"

CUSTOM_CSS = f"""
<style>
  .stApp {{ background: linear-gradient(180deg, {LIGHT_BG} 0%, #FFFFFF 100%); }}
  .ld-header {{
      display: flex; align-items: center; gap: 16px;
      padding: 18px 24px;
      background: linear-gradient(90deg, {PRIMARY} 0%, #2D4F8E 100%);
      border-radius: 14px;
      color: white;
      margin-bottom: 18px;
      box-shadow: 0 6px 18px rgba(31,56,100,0.18);
  }}
  .ld-header h1 {{ color: white; margin: 0; font-size: 24px; font-weight: 700; }}
  .ld-header p {{ color: #E9EEF7; margin: 2px 0 0 0; font-size: 14px; }}
  .ld-version-tag {{
      background: rgba(255,255,255,0.18);
      padding: 4px 10px; border-radius: 999px;
      font-size: 11px; font-weight: 600;
      margin-left: auto;
  }}
  .ld-card {{
      background: white; border-radius: 12px; padding: 18px 20px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.06);
      border: 1px solid #E5E9F2;
      margin-bottom: 14px;
  }}
  .ld-card h3 {{ margin: 0 0 8px 0; color: {PRIMARY}; font-size: 17px; }}
  .ld-metric {{
      background: white; border-radius: 10px; padding: 14px 16px;
      border-left: 5px solid {ACCENT};
      box-shadow: 0 1px 6px rgba(0,0,0,0.05);
  }}
  .ld-metric .label {{ color: #6B7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }}
  .ld-metric .value {{ color: {PRIMARY}; font-size: 22px; font-weight: 700; margin-top: 4px; }}
  .stButton > button {{
      background: {PRIMARY}; color: white; border: none;
      font-weight: 600; border-radius: 10px;
      padding: 10px 18px; transition: all 0.18s ease;
  }}
  .stButton > button:hover {{
      background: {ACCENT}; color: white; transform: translateY(-1px);
      box-shadow: 0 4px 12px rgba(232,115,26,0.3);
  }}
  .stTabs [data-baseweb="tab-list"] {{ gap: 6px; }}
  .stTabs [data-baseweb="tab"] {{
      background: white; border-radius: 8px 8px 0 0;
      padding: 10px 16px; font-weight: 600;
  }}
  .stTabs [aria-selected="true"] {{ background: {PRIMARY} !important; color: white !important; }}
  div[data-testid="stFileUploadDropzone"] {{
      border: 2px dashed {PRIMARY};
      background: #FAFBFE;
      border-radius: 12px;
  }}
  .ld-footer {{
      text-align: center; color: #6B7280; font-size: 12px;
      margin-top: 30px; padding: 12px;
  }}
  #MainMenu {{ visibility: hidden; }}
  footer {{ visibility: hidden; }}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ============================================================
# Logo (SVG embutido como fallback se assets/logo.png não existir)
# ============================================================
LOGO_SVG = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 70" width="180" height="52">
  <defs>
    <linearGradient id="g1" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#E8731A"/>
      <stop offset="100%" stop-color="#F5A04E"/>
    </linearGradient>
  </defs>
  <path d="M22 6 C 14 18, 8 28, 14 38 C 18 46, 28 46, 32 38 C 38 28, 30 18, 22 6 Z" fill="url(#g1)"/>
  <text x="48" y="36" font-family="Arial Black, sans-serif" font-size="26" font-weight="900" fill="#1F3864" letter-spacing="-1">LIDER</text>
  <text x="48" y="56" font-family="Arial, sans-serif" font-size="13" font-weight="700" fill="#6B7280" letter-spacing="2">LIMPE</text>
</svg>
"""

def render_header():
    logo_path = Path("assets/logo.png")
    if logo_path.exists() and logo_path.stat().st_size > 200:
        with open(logo_path, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:52px;border-radius:8px;background:white;padding:4px;">'
    else:
        logo_html = LOGO_SVG

    st.markdown(f"""
    <div class="ld-header">
      {logo_html}
      <div>
        <h1>Painel de Processamento de Planos</h1>
        <p>Domínio · Unimed Saúde · Unimed Odonto · VSP SAMP</p>
      </div>
      <div class="ld-version-tag">{APP_VERSION} · {APP_BUILD}</div>
    </div>
    """, unsafe_allow_html=True)


render_header()


# ============================================================
# Estado da sessão
# ============================================================
if "uploads" not in st.session_state:
    st.session_state.uploads = []
if "resultado" not in st.session_state:
    st.session_state.resultado = None


EMP_LIST = [
    ("ATIVA",          "🏢 ATIVA",           "#1F3864"),
    ("COMERCIAL",      "🏬 L. COMERCIAL",    "#0070C0"),
    ("MULTISSERVICOS", "🛠️ L. MULTISSERV.",  "#548235"),
    ("VSP",            "🛡️ VSP",             "#C00000"),
]


# ============================================================
# Sidebar
# ============================================================
with st.sidebar:
    st.markdown(f"### 🧹 Líder Limpe\n**Versão:** `{APP_VERSION}`\n\n**Build:** `{APP_BUILD}`")
    st.markdown("---")
    st.markdown("### 📚 Como usar")
    st.markdown("""
    **1.** Suba os arquivos por empresa (recomendado) ou todos juntos.

    **2.** Clique em **Gerar TUDO** ou em um módulo específico.

    **3.** Baixe o ZIP consolidado com 1 sub-ZIP por empresa + `RELAÇÃO CCUSTO - UNIMED.xlsx`.
    """)

    st.markdown("---")
    st.markdown("### 📌 Identificação de empresa")
    st.markdown("""
    - **VSP**: contratos `6217` / `5964`
    - **ATIVA**: `codi_emp` 2 / lotação `0001`
    - **L. COMERCIAL**: `codi_emp` 4 / lotação `0002`
    - **L. MULTISSERVIÇOS**: `codi_emp` 3 / lotação `0003`
    """)

    st.markdown("---")
    if st.button("🗑️ Limpar tudo", use_container_width=True):
        st.session_state.uploads = []
        st.session_state.resultado = None
        st.rerun()


# ============================================================
# Abas
# ============================================================
tab_upload, tab_gerar, tab_resultado, tab_ajuda = st.tabs([
    "📤 Upload de arquivos",
    "⚙️ Processar",
    "📥 Resultado",
    "❓ Ajuda",
])


# ============================================================
# TAB 1 - Upload
# ============================================================
with tab_upload:
    st.markdown('<div class="ld-card"><h3>📁 Envio por empresa (recomendado)</h3>'
                'Carregue 1 ZIP <b>por empresa</b> contendo as planilhas dela '
                '(Domínio, Saúde Geral/Santas, Odonto TXT, SAMP no caso da VSP). '
                'Você também pode soltar arquivos .xls/.xlsx/.txt diretamente.'
                '</div>', unsafe_allow_html=True)

    cols = st.columns(4)
    for (chave, label, cor), col in zip(EMP_LIST, cols):
        with col:
            st.markdown(f'<div class="ld-card" style="border-top:4px solid {cor};">'
                        f'<h3 style="color:{cor};margin:0;">{label}</h3></div>', unsafe_allow_html=True)
            files = st.file_uploader(
                f"Arquivos da {label}",
                type=["zip", "xls", "xlsx", "txt"],
                accept_multiple_files=True,
                key=f"upl_{chave}",
                label_visibility="collapsed",
            )
            if files:
                for f in files:
                    # Evita duplicados (mesma chave + mesmo nome)
                    ja_existe = any(
                        u["nome"] == f.name and u.get("empresa_hint") == chave
                        for u in st.session_state.uploads
                    )
                    if not ja_existe:
                        st.session_state.uploads.append({
                            "empresa_hint": chave,
                            "nome": f.name,
                            "dados": f.getvalue(),
                            "pasta": chave,
                        })
                st.success(f"✅ {len(files)} arquivo(s) carregado(s)")

    st.markdown('<div class="ld-card"><h3>📥 Envio livre (detecção automática)</h3>'
                'Mande tudo de uma vez ou arquivos avulsos. '
                'O app identifica empresa pelo contrato/codi_emp.'
                '</div>', unsafe_allow_html=True)

    files_livre = st.file_uploader(
        "Arquivos avulsos (ZIP/XLS/XLSX/TXT)",
        type=["zip", "xls", "xlsx", "txt"],
        accept_multiple_files=True,
        key="upl_livre",
    )
    if files_livre:
        for f in files_livre:
            ja_existe = any(
                u["nome"] == f.name and u.get("empresa_hint") is None
                for u in st.session_state.uploads
            )
            if not ja_existe:
                st.session_state.uploads.append({
                    "empresa_hint": None,
                    "nome": f.name,
                    "dados": f.getvalue(),
                    "pasta": "",
                })
        st.success(f"✅ {len(files_livre)} arquivo(s) carregado(s)")

    # Resumo
    st.markdown("### 📋 Arquivos prontos para processar")
    if not st.session_state.uploads:
        st.info("Nenhum arquivo enviado ainda.")
    else:
        df_resumo = pd.DataFrame([{
            "Empresa (hint)": u.get("empresa_hint") or "(auto)",
            "Arquivo": u["nome"],
            "Tamanho": f"{len(u['dados'])/1024:.1f} KB",
        } for u in st.session_state.uploads])
        st.dataframe(df_resumo, use_container_width=True, hide_index=True)

        c1, c2 = st.columns([1, 4])
        with c1:
            if st.button("🗑️ Remover todos", use_container_width=True):
                st.session_state.uploads = []
                st.rerun()


# ============================================================
# TAB 2 - Processar
# ============================================================
with tab_gerar:
    st.markdown('<div class="ld-card"><h3>⚙️ Selecione o que processar</h3>'
                'Você pode rodar tudo de uma vez ou um módulo específico.'
                '</div>', unsafe_allow_html=True)

    if not st.session_state.uploads:
        st.warning("⚠️ Envie arquivos na aba **Upload** primeiro.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            btn_tudo = st.button("🚀 Gerar TUDO", use_container_width=True, type="primary")
        with c2:
            btn_saude = st.button("💊 Só Saúde", use_container_width=True)
        with c3:
            btn_odonto = st.button("🦷 Só Odonto", use_container_width=True)
        with c4:
            btn_samp = st.button("🛡️ Só VSP SAMP", use_container_width=True)

        def filtrar_uploads(modulo: str):
            todos = [(u["nome"], u["dados"], u["pasta"]) for u in st.session_state.uploads]
            if modulo == "TUDO":
                return todos
            # Sempre precisa de Domínio e dos ZIPs
            base = [t for t in todos if "DOMINIO" in t[0].upper() or "DOMÍNIO" in t[0].upper()
                    or t[0].lower().endswith(".zip")]
            if modulo == "SAUDE":
                base += [t for t in todos if ("SAUDE" in t[0].upper() or "SAÚDE" in t[0].upper()
                                              or "EXTRATO" in t[0].upper()) and "ODONTO" not in t[0].upper()]
            elif modulo == "ODONTO":
                base += [t for t in todos if t[0].lower().endswith(".txt")]
            elif modulo == "SAMP":
                base += [t for t in todos if ("SINDSEG" in (t[0]+t[2]).upper()
                                              or "SINDIVIGILANTES" in (t[0]+t[2]).upper()
                                              or "CONTROLE DE PAGAMENTOS" in t[0].upper())]
            visto = set()
            out = []
            for t in base:
                if t[0] not in visto:
                    visto.add(t[0])
                    out.append(t)
            return out

        def rodar(modulo: str):
            uploads_filt = filtrar_uploads(modulo)
            if not uploads_filt:
                st.error("Nenhum arquivo relevante para este módulo.")
                return
            with st.spinner(f"🔄 Processando {modulo}..."):
                try:
                    resultado = processar_uploads(uploads_filt, modelos=None)
                    st.session_state.resultado = resultado
                    st.session_state.resultado["modulo"] = modulo
                    st.success(f"✅ Processamento concluído!")
                except Exception as e:
                    import traceback
                    st.error(f"❌ Erro fatal: {e}")
                    st.code(traceback.format_exc())

        if btn_tudo: rodar("TUDO")
        elif btn_saude: rodar("SAUDE")
        elif btn_odonto: rodar("ODONTO")
        elif btn_samp: rodar("SAMP")

    if st.session_state.resultado:
        with st.expander("📜 Log de processamento", expanded=True):
            for linha in st.session_state.resultado["log"]:
                st.markdown(f"- {linha}")


# ============================================================
# TAB 3 - Resultado
# ============================================================
with tab_resultado:
    if not st.session_state.resultado:
        st.info("Nenhum processamento concluído ainda. Vá para **Processar** primeiro.")
    else:
        res = st.session_state.resultado

        # Métricas
        st.markdown("### 📊 Resumo por empresa")
        cols = st.columns(4)
        for (emp, label, cor), col in zip(EMP_LIST, cols):
            r = res["resumo"].get(emp, {})
            qtd = r.get("dominio_qtd", 0)
            valores = r.get("valores", {})
            total = sum(valores.values()) if valores else 0
            valor_fmt = f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            with col:
                st.markdown(
                    f'<div class="ld-metric" style="border-left-color:{cor};">'
                    f'<div class="label">{label}</div>'
                    f'<div class="value">{valor_fmt}</div>'
                    f'<div style="color:#6B7280;font-size:12px;margin-top:4px;">'
                    f'{qtd} colaboradores</div></div>',
                    unsafe_allow_html=True,
                )

        # Detalhes por empresa
        st.markdown("### 📂 Totais por contrato")
        for emp, label, cor in EMP_LIST:
            r = res["resumo"].get(emp, {})
            if not r or (r.get("dominio_qtd", 0) == 0 and not r.get("valores")):
                continue
            with st.expander(f"{label}", expanded=False):
                valores = r.get("valores", {})
                if valores:
                    df_v = pd.DataFrame([
                        {"Contrato": k.replace("_", " ").title(), "Valor total": v}
                        for k, v in valores.items()
                    ])
                    st.dataframe(df_v, hide_index=True, use_container_width=True,
                                 column_config={
                                     "Valor total": st.column_config.NumberColumn(format="R$ %.2f")
                                 })
                else:
                    st.info("Sem valores processados.")

        # Downloads
        st.markdown("### 📥 Downloads")
        c1, c2 = st.columns([1, 1])
        with c1:
            st.download_button(
                "⬇️ Baixar ZIP geral",
                data=res["zip_final"],
                file_name=f"RESULTADOS_{datetime.now().strftime('%Y-%m-%d_%H%M')}.zip",
                mime="application/zip",
                use_container_width=True,
                type="primary",
            )
        with c2:
            st.metric("Arquivos no ZIP", len(res["arquivos"]))

        st.markdown("#### Downloads individuais")
        for fname, data in res["arquivos"].items():
            ext = fname.split(".")[-1].lower()
            icon = "📦" if ext == "zip" else "📄"
            st.download_button(
                f"{icon} {fname}",
                data=data,
                file_name=fname,
                mime="application/zip" if ext == "zip" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{fname}",
            )


# ============================================================
# TAB 4 - Ajuda
# ============================================================
with tab_ajuda:
    st.markdown(f"""
    ### 📖 Como o app funciona ({APP_VERSION})

    #### Domínio
    - Aceita `.xls` e `.xlsx`.
    - Demitidos = `situacao == 8`. Para os demais, limpa `datasituacao`.
    - Remove CPFs duplicados sem `groupby.apply` (compatível com pandas 2.2+).
    - Mantém ativos sobre demitidos; mais novo (maior `i_empregados`) sobre antigo.

    #### Saúde Unimed
    - Contratos: `5957` (Ambulatorial), `6040` (Santas), `6217` (Interior VSP), `5964` (Metropolitano VSP).
    - Cruza `CPFTITULAR × Domínio` para puxar **CCusto** (`nome_quebra`).
    - Soma `VLFATURADO` por CCusto.

    #### Odonto Unimed (TXT)
    - Parse do `.txt` delimitado por `#`.
    - Identifica empresa pelo CNPJ.
    - Soma "Mensalidade" por CCusto.

    #### VSP SAMP
    - Identifica SINDSEG vs SINDIVIGILANTES por nome do arquivo + pasta no ZIP.
    - Tipo de boleto pelo valor majoritário (15,00 / 18,50 / 10,00).

    #### Saída
    - `RESULTADOS_AAAA-MM-DD_HHMM.zip` com:
      - `ATIVA.zip`, `COMERCIAL.zip`, `MULTISSERVICOS.zip`, `VSP.zip`
      - `RELAÇÃO CCUSTO - UNIMED.xlsx`
    """)


st.markdown(
    f'<div class="ld-footer">Líder Limpe — Painel de Processamento de Planos • '
    f'{APP_VERSION} build {APP_BUILD}</div>',
    unsafe_allow_html=True,
)
