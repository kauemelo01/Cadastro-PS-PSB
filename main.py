"""
Pastoral Social — Consulta de Cadastro
Mobile-first Streamlit app that reads from a GitHub-hosted xlsx file.

Setup:
  1. Upload cadastro.xlsx to your GitHub repo.
  2. Set GITHUB_RAW_URL below (or use st.secrets["GITHUB_RAW_URL"]).
  3. Deploy to Streamlit Community Cloud pointing at this repo.
"""

import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime, timedelta

# ──────────────────────────────────────────────────────────────
#  CONFIG — Replace with your own GitHub raw URL
#  Format: https://raw.githubusercontent.com/{USER}/{REPO}/{BRANCH}/cadastro.xlsx
# ──────────────────────────────────────────────────────────────
try:
    GITHUB_RAW_URL = st.secrets["GITHUB_RAW_URL"]
except Exception:
    GITHUB_RAW_URL = (
        "https://github.com/kauemelo01/Cadastro-PS-PSB/raw/refs/heads/main/cadastro.xlsx"
    )

# ──────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pastoral Social",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────────────────────
#  MOBILE-FIRST CSS
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
  /* ── Layout ── */
  .main .block-container {
    padding: 0.75rem 0.75rem 2rem;
    max-width: 500px;
    margin: auto;
  }

  /* ── Header ── */
  .ps-header {
    background: linear-gradient(135deg, #1b5e20 0%, #388e3c 100%);
    border-radius: 14px;
    padding: 1rem 1.25rem 0.75rem;
    color: #fff;
    margin-bottom: 1rem;
  }
  .ps-header h1 { margin: 0; font-size: 1.35rem; color: #fff; }
  .ps-header p  { margin: 0.15rem 0 0; font-size: 0.82rem; opacity: 0.85; }

  /* ── Search box ── */
  .stTextInput > div > div > input {
    font-size: 16px !important;
    border-radius: 12px !important;
    border: 2px solid #c8e6c9 !important;
    padding: 0.65rem 1rem !important;
    background: #f9fbe7 !important;
  }
  .stTextInput > div > div > input:focus {
    border-color: #388e3c !important;
    box-shadow: 0 0 0 3px rgba(56,142,60,.15) !important;
  }

  /* ── Result card ── */
  .card {
    background: #fff;
    border-radius: 14px;
    padding: 1rem 1.1rem;
    margin-bottom: 0.85rem;
    box-shadow: 0 2px 10px rgba(0,0,0,.08);
    border-left: 5px solid #388e3c;
  }
  .card-name {
    font-size: 1.05rem;
    font-weight: 700;
    color: #1b5e20;
    margin-bottom: 5px;
  }
  .card-num {
    font-size: 0.78rem;
    color: #757575;
    margin-bottom: 8px;
  }

  /* ── Badges ── */
  .badge {
    display: inline-block;
    padding: 3px 11px;
    border-radius: 20px;
    font-size: 0.77rem;
    font-weight: 700;
    margin: 2px 3px 2px 0;
  }
  .b-tipo    { background: #e3f2fd; color: #1565c0; }
  .b-ativo   { background: #e8f5e9; color: #2e7d32; }
  .b-inativo { background: #ffebee; color: #b71c1c; }
  .b-cid     { background: #fce4ec; color: #880e4f; }

  /* ── Section titles inside card ── */
  .sec-title {
    font-size: 0.78rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: #616161;
    margin: 10px 0 5px;
  }

  /* ── Alert box ── */
  .alert-box {
    background: #fff8e1;
    border: 1px solid #ffe082;
    border-radius: 8px;
    padding: 7px 10px;
    font-size: 0.88rem;
    margin-top: 8px;
    color: #5d4037;
  }

  /* ── Monthly delivery grid ── */
  .month-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 5px;
    margin-top: 4px;
  }
  .m-ok {
    background: #e8f5e9;
    border-radius: 7px;
    padding: 4px 3px;
    text-align: center;
    font-size: 0.75rem;
    color: #2e7d32;
    font-weight: 600;
  }
  .m-no {
    background: #f5f5f5;
    border-radius: 7px;
    padding: 4px 3px;
    text-align: center;
    font-size: 0.75rem;
    color: #bdbdbd;
  }

  /* ── Tier-2 hint ── */
  .tier2-hint {
    background: #fff3e0;
    border: 1px solid #ffcc80;
    border-radius: 9px;
    padding: 8px 11px;
    margin-top: 10px;
    font-size: 0.83rem;
    color: #6d4c41;
  }

  /* ── Empty state ── */
  .empty-state {
    text-align: center;
    padding: 3rem 1rem;
    color: #9e9e9e;
  }
  .empty-state .icon { font-size: 2.8rem; }
  .empty-state p { margin: 0.5rem 0 0; font-size: 0.95rem; }

  /* ── Result count ── */
  .result-count {
    font-size: 0.82rem;
    color: #757575;
    margin-bottom: 0.6rem;
    padding-left: 2px;
  }

  /* ── Hide Streamlit chrome ── */
  #MainMenu { visibility: hidden; }
  footer    { visibility: hidden; }
  header    { visibility: hidden; }
</style>
""",
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────────────────────
_MONTHS_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
               "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def serial_to_month(serial: str) -> str:
    """Convert an Excel date serial number string to 'Mmm/YY'."""
    try:
        n = int(float(serial))
        dt = datetime(1899, 12, 30) + timedelta(days=n)
        return f"{_MONTHS_PT[dt.month - 1]}/{str(dt.year)[-2:]}"
    except Exception:
        return str(serial)


def is_numeric(s: str) -> bool:
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False


# ──────────────────────────────────────────────────────────────
#  DATA LOADING
# ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_data(url: str):
    """
    Returns (df, meta, error_msg).
    meta = {col_name: tag}  where tag ∈ {"SEARCH", "Tier 1", "Tier 2"}
    """
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
    except Exception as exc:
        return None, None, str(exc)

    try:
        raw = pd.read_excel(BytesIO(resp.content), header=None, dtype=str)
    except Exception as exc:
        return None, None, str(exc)

    if raw.shape[0] < 3:
        return None, None, "Arquivo sem dados suficientes."

    tags_row    = raw.iloc[0].tolist()   # Row 0: SEARCH / Tier 1 / Tier 2
    names_row   = raw.iloc[1].tolist()   # Row 1: human-readable column names

    # Build deduplicated column names
    seen: dict[str, int] = {}
    final_names: list[str] = []
    final_tags:  list[str] = []

    for tag, name in zip(tags_row, names_row):
        raw_name = str(name).strip() if pd.notna(name) else ""
        tag_str  = str(tag).strip()  if pd.notna(tag)  else ""

        # Convert numeric column names (date serials) to month strings
        display_name = serial_to_month(raw_name) if is_numeric(raw_name) else raw_name

        # Deduplicate
        if display_name in seen:
            seen[display_name] += 1
            display_name = f"{display_name}_{seen[display_name]}"
        else:
            seen[display_name] = 0

        final_names.append(display_name)
        final_tags.append(tag_str)

    df = raw.iloc[2:].copy()
    df.columns = final_names
    df = df.reset_index(drop=True)

    meta = dict(zip(final_names, final_tags))
    return df, meta, None


# ──────────────────────────────────────────────────────────────
#  LOAD DATA
# ──────────────────────────────────────────────────────────────
with st.spinner("Carregando cadastro…"):
    df, meta, load_error = load_data(GITHUB_RAW_URL)


def cols_by_tag(tag: str) -> list[str]:
    if not meta:
        return []
    return [c for c, t in meta.items() if t == tag]


SEARCH_COLS = cols_by_tag("SEARCH")
TIER1_COLS  = cols_by_tag("Tier 1")
TIER2_COLS  = cols_by_tag("Tier 2")

# Monthly columns are Tier 1 columns whose names look like "Mmm/YY"
MONTH_COLS  = [c for c in TIER1_COLS if "/" in c and len(c) == 6]
INFO_COLS   = [c for c in TIER1_COLS if c not in MONTH_COLS]


# ──────────────────────────────────────────────────────────────
#  SEARCH FUNCTION
# ──────────────────────────────────────────────────────────────
def search_df(query: str) -> pd.DataFrame:
    q = query.strip().lower()
    if not q or df is None:
        return pd.DataFrame()
    mask = pd.Series([False] * len(df), index=df.index)
    for col in SEARCH_COLS:
        if col in df.columns:
            mask |= df[col].fillna("").str.lower().str.contains(q, regex=False)
    return df[mask].copy()


# ──────────────────────────────────────────────────────────────
#  RENDER RECORD
# ──────────────────────────────────────────────────────────────
def cell(row: pd.Series, col: str) -> str:
    v = str(row.get(col, "")).strip()
    return "" if v in ("nan", "None", "") else v


def render_record(row: pd.Series) -> None:
    nome   = cell(row, "NOME")
    numero = cell(row, "NUMERO")
    tipo   = cell(row, "TIPO")
    status = cell(row, "STATUS")
    cid    = cell(row, "CID 2026")
    alerta = cell(row, "ALERTA PARA A MESA")

    status_cls = "b-ativo" if status.lower() == "ativo" else "b-inativo"

    # ── Header
    html = f"""
    <div class="card">
      <div class="card-name">{nome or "—"}</div>
      <div class="card-num">Nº {numero}</div>
      <div>
        <span class="badge b-tipo">{tipo}</span>
        <span class="badge {status_cls}">{status}</span>
        {f'<span class="badge b-cid">CID {cid}</span>' if cid else ""}
      </div>
    """

    # ── Alert
    if alerta and alerta not in ("0",):
        html += f'<div class="alert-box">⚠️ {alerta}</div>'

    # ── Monthly deliveries
    month_data = [(mc, cell(row, mc)) for mc in MONTH_COLS]
    if any(v for _, v in month_data):
        html += '<div class="sec-title">📦 Histórico de Entregas</div>'
        html += '<div class="month-grid">'
        for mc, val in month_data:
            if val.lower() == "ok":
                html += f'<div class="m-ok">✓ {mc}</div>'
            else:
                html += f'<div class="m-no">{mc}</div>'
        html += "</div>"

    # ── Tier-2 hint
    html += f"""
      <div class="tier2-hint">
        🔒 <strong>Dados adicionais disponíveis</strong> para usuários autorizados
        &nbsp;·&nbsp; <em>{len(TIER2_COLS)} campos restritos</em>
        <br><small>(endereço, contato, informações pessoais, observações internas…)</small>
      </div>
    </div>
    """

    st.markdown(html, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
#  UI — HEADER
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
<div class="ps-header">
  <h1>📋 Pastoral Social</h1>
  <p>Consulta de Cadastro · Jardim Colombo</p>
</div>
""",
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────────────────────
#  UI — ERROR STATE
# ──────────────────────────────────────────────────────────────
if load_error or df is None:
    st.error("❌ Não foi possível carregar o arquivo.")
    st.markdown(
        """
**Possíveis causas:**
- A URL do arquivo no GitHub ainda não foi configurada.
- O repositório ou arquivo não existe / é privado.
- Problema de rede.

**Solução:** edite `app.py` e defina `GITHUB_RAW_URL` com a URL raw do seu arquivo,
ou crie o arquivo `secrets.toml` com a chave `GITHUB_RAW_URL`.
"""
    )
    if load_error:
        with st.expander("Detalhes do erro"):
            st.code(load_error)
    st.stop()

# ──────────────────────────────────────────────────────────────
#  UI — SEARCH
# ──────────────────────────────────────────────────────────────
query = st.text_input(
    "Buscar",
    placeholder="🔍  Nome, CPF, número do registro ou reserva…",
    label_visibility="collapsed",
)

# ──────────────────────────────────────────────────────────────
#  UI — RESULTS
# ──────────────────────────────────────────────────────────────
if query:
    results = search_df(query)

    if results.empty:
        st.markdown(
            """
<div class="empty-state">
  <div class="icon">🔎</div>
  <p>Nenhum registro encontrado.<br>
  Tente outro nome, CPF ou número.</p>
</div>
""",
            unsafe_allow_html=True,
        )
    else:
        count = len(results)
        st.markdown(
            f'<div class="result-count">{count} registro{"s" if count > 1 else ""} encontrado{"s" if count > 1 else ""}</div>',
            unsafe_allow_html=True,
        )
        for _, row in results.iterrows():
            render_record(row)
else:
    st.markdown(
        f"""
<div class="empty-state">
  <div class="icon">🔍</div>
  <p>Digite um nome, CPF, número<br>ou nome de reserva para buscar.</p>
  <p style="margin-top:1.5rem;font-size:0.8rem;color:#bdbdbd">
    {len(df)} registros · {len(SEARCH_COLS)} campos de busca
  </p>
</div>
""",
        unsafe_allow_html=True,
    )
