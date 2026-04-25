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

  /* ── Search box — light mode ── */
  .stTextInput > div > div > input {
    font-size: 16px !important;
    border-radius: 12px !important;
    border: 2px solid #c8e6c9 !important;
    padding: 0.65rem 1rem !important;
    background: #f9fbe7 !important;
    color: #212121 !important;
  }
  .stTextInput > div > div > input:focus {
    border-color: #388e3c !important;
    box-shadow: 0 0 0 3px rgba(56,142,60,.15) !important;
  }
  /* ── Search box — dark mode ── */
  @media (prefers-color-scheme: dark) {
    .stTextInput > div > div > input {
      background: #1a2e1a !important;
      color: #e8f5e9 !important;
      border-color: #4caf50 !important;
    }
    .stTextInput > div > div > input::placeholder {
      color: #81c784 !important;
      opacity: 0.7;
    }
    .card { background: #1e2b1e !important; }
    .info-value { color: #e0e0e0 !important; }
    .card-name  { color: #a5d6a7 !important; }
    .card-num   { color: #9e9e9e !important; }
    .sec-title  { color: #9e9e9e !important; }
    .info-grid  { border-top-color: #2e3e2e !important; }
    .info-row   { border-bottom-color: #2a3a2a !important; }
    .info-alert-row { background: #3e2e10 !important; }
    .tier2-hint { background: #2e2010 !important; border-color: #795548 !important; color: #bcaaa4 !important; }
    .m-no { background: #2a352a !important; }
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

  /* ── Info grid (labeled rows) ── */
  .info-grid {
    margin-top: 10px;
    border-top: 1px solid #f0f0f0;
    padding-top: 8px;
  }
  .info-row {
    display: flex;
    align-items: flex-start;
    padding: 5px 0;
    border-bottom: 1px solid #f9f9f9;
    gap: 8px;
  }
  .info-label {
    font-size: 0.75rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.03em;
    color: #9e9e9e;
    min-width: 80px;
    flex-shrink: 0;
    padding-top: 3px;
  }
  .info-value {
    font-size: 0.9rem;
    color: #212121;
    flex: 1;
  }
  .info-empty { color: #bdbdbd; font-style: italic; }
  .info-alert-row { background: #fff8e1; border-radius: 6px; padding: 5px 7px; margin: 2px 0; }
  .info-alert-val { color: #6d4c41; font-weight: 600; }

  /* ── Monthly delivery grid ── */
  .month-grid {
    display: grid;
    grid-template-columns: repeat(6, 1fr);
    gap: 4px;
    margin-top: 4px;
  }
  .m-ok {
    background: #e8f5e9;
    border-radius: 6px;
    padding: 3px 2px;
    text-align: center;
    font-size: 0.68rem;
    color: #2e7d32;
    font-weight: 700;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: clip;
  }
  .m-no {
    background: #eeeeee;
    border-radius: 6px;
    min-height: 22px;
    padding: 3px 2px;
  }

  /* ── Search highlight ── */
  mark.hl {
    background: #fff176;
    color: #212121;
    border-radius: 3px;
    padding: 0 2px;
    font-weight: 700;
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

        # Convert date column names (Excel serial numbers OR datetime strings) to "Mmm/YY"
        if is_numeric(raw_name):
            display_name = serial_to_month(raw_name)
        else:
            # pandas reads Excel dates as "YYYY-MM-DD HH:MM:SS" strings when dtype=str
            try:
                dt = pd.to_datetime(raw_name, errors="raise")
                display_name = f"{_MONTHS_PT[dt.month - 1]}/{str(dt.year)[-2:]}"
            except Exception:
                display_name = raw_name

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
_CPF_COLS = {"CPF", "CPF RESERVA 1", "CPF RESERVA 2"}


def _strip_cpf(s: str) -> str:
    """Remove dots and dashes so raw and formatted CPFs compare equal."""
    return s.replace(".", "").replace("-", "")


def search_df(query: str, cols: list[str] | None = None) -> pd.DataFrame:
    q = query.strip().lower()
    if not q or df is None:
        return pd.DataFrame()
    if cols is None:
        cols = SEARCH_COLS
    q_cpf = _strip_cpf(q)          # normalised query for CPF columns
    mask = pd.Series([False] * len(df), index=df.index)
    for col in cols:
        if col not in df.columns:
            continue
        if col in _CPF_COLS:
            # compare stripped stored value against stripped query
            mask |= (
                df[col].fillna("")
                .apply(_strip_cpf)
                .str.lower()
                .str.contains(q_cpf, regex=False)
            )
        else:
            mask |= df[col].fillna("").str.lower().str.contains(q, regex=False)
    return df[mask].copy()


# ──────────────────────────────────────────────────────────────
#  RENDER RECORD
# ──────────────────────────────────────────────────────────────
def cell(row: pd.Series, col: str) -> str:
    v = str(row.get(col, "")).strip()
    return "" if v in ("nan", "None", "") else v


import re as _re

def highlight(text: str, query: str) -> str:
    """Wrap every occurrence of query in text with a highlight span."""
    if not query or not text:
        return text
    escaped = _re.escape(query)
    return _re.sub(
        f"({escaped})",
        r'<mark class="hl">\1</mark>',
        text,
        flags=_re.IGNORECASE,
    )


def render_record(row: pd.Series, query: str = "") -> None:
    q = query.strip().lower()

    nome   = cell(row, "NOME")
    numero = cell(row, "NUMERO")
    tipo   = cell(row, "TIPO")
    status = cell(row, "STATUS")
    cid    = cell(row, "CID 2026")
    alerta = cell(row, "ALERTA PARA A MESA")
    reserva1     = cell(row, "RESERVA 1")
    cpf_reserva1 = cell(row, "CPF RESERVA 1")
    reserva2     = cell(row, "RESERVA 2")
    cpf_reserva2 = cell(row, "CPF RESERVA 2")

    status_cls = "b-ativo" if status.lower() == "ativo" else "b-inativo"

    # ── Header: name + number
    html = f"""
    <div class="card">
      <div class="card-name">{highlight(nome, q) or "—"}</div>
      <div class="card-num">Nº {highlight(numero, q)}</div>
    """

    # ── Info rows: always show all Tier 1 fields explicitly
    html += '<div class="info-grid">'

    html += f"""
      <div class="info-row">
        <span class="info-label">Tipo</span>
        <span class="info-value"><span class="badge b-tipo">{tipo or "—"}</span></span>
      </div>
      <div class="info-row">
        <span class="info-label">Status</span>
        <span class="info-value"><span class="badge {status_cls}">{status or "—"}</span></span>
      </div>
      <div class="info-row">
        <span class="info-label">CID 2026</span>
        <span class="info-value">{highlight(cid, q) if cid else '<span class="info-empty">—</span>'}</span>
      </div>
    """

    # ── RESERVA rows
    def reserva_row(label: str, val: str) -> str:
        v_hl = highlight(val, q) if val else '<span class="info-empty">—</span>'
        row_cls = ' info-alert-row' if val and q and q in val.lower() else ''
        return f'''
      <div class="info-row{row_cls}">
        <span class="info-label">{label}</span>
        <span class="info-value">{v_hl}</span>
      </div>'''

    html += reserva_row("Reserva 1", reserva1)
    html += reserva_row("CPF Res. 1", cpf_reserva1)
    html += reserva_row("Reserva 2", reserva2)
    html += reserva_row("CPF Res. 2", cpf_reserva2)

    # ── Alert row — always visible
    if alerta and alerta not in ("0",):
        html += f"""
      <div class="info-row info-alert-row">
        <span class="info-label">⚠️ Alerta</span>
        <span class="info-value info-alert-val">{highlight(alerta, q)}</span>
      </div>
        """
    else:
        html += """
      <div class="info-row">
        <span class="info-label">Alerta</span>
        <span class="info-value info-empty">Nenhum</span>
      </div>
        """

    # ── Any other non-monthly Tier 1 columns not yet handled above
    already_shown = {"TIPO", "CID 2026", "STATUS", "ALERTA PARA A MESA",
                     "RESERVA 1", "CPF RESERVA 1", "RESERVA 2", "CPF RESERVA 2"}
    for col in INFO_COLS:
        if col not in already_shown:
            val = cell(row, col)
            html += f"""
      <div class="info-row">
        <span class="info-label">{col}</span>
        <span class="info-value">{highlight(val, q) if val else '<span class="info-empty">—</span>'}</span>
      </div>
            """

    html += "</div>"  # close info-grid

    # ── Monthly deliveries — compact grid: OK = green cell, blank = empty dot
    html += '<div class="sec-title">📦 Histórico de Entregas</div>'
    html += '<div class="month-grid">'
    for mc in MONTH_COLS:
        val = cell(row, mc)
        if val.lower() == "ok":
            html += f'<div class="m-ok" title="{mc}">✓ {mc}</div>'
        else:
            html += f'<div class="m-no" title="{mc}"></div>'
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
query_num = st.text_input(
    "Buscar por Número",
    placeholder="🔢  Número do registro…",
    label_visibility="collapsed",
    key="q_num",
)
query_gen = st.text_input(
    "Buscar geral",
    placeholder="🔍  Nome, CPF, reserva…",
    label_visibility="collapsed",
    key="q_gen",
)


def show_results(results: pd.DataFrame, query: str) -> None:
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
            render_record(row, query=query)


# ──────────────────────────────────────────────────────────────
#  UI — RESULTS
# ──────────────────────────────────────────────────────────────
if query_num:
    show_results(search_df(query_num, cols=["NUMERO"]), query_num)
elif query_gen:
    show_results(search_df(query_gen), query_gen)
else:
    st.markdown(
        f"""
<div class="empty-state">
  <div class="icon">🔍</div>
  <p>Digite um número de registro<br>ou busque por nome, CPF ou reserva.</p>
  <p style="margin-top:1.5rem;font-size:0.8rem;color:#bdbdbd">
    {len(df)} registros · {len(SEARCH_COLS)} campos de busca
  </p>
</div>
""",
        unsafe_allow_html=True,
    )
