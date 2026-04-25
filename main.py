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
def _fetch_gdrive(url: str) -> bytes:
    """
    Download a file from a Google Drive 'uc?export=download' URL.
    Handles the virus-scan confirmation redirect that Drive adds for larger files.
    """
    session = requests.Session()
    resp = session.get(url, timeout=15)
    resp.raise_for_status()

    # Drive returns an HTML confirmation page when the file is large.
    # Detect it by content-type and extract the confirmed download URL.
    if "text/html" in resp.headers.get("Content-Type", ""):
        # Look for the confirmation token in the response
        import re as _re2
        match = _re2.search(r'confirm=([0-9A-Za-z_\-]+)', resp.text)
        token = match.group(1) if match else "t"
        confirmed_url = url + f"&confirm={token}"
        resp = session.get(confirmed_url, timeout=30)
        resp.raise_for_status()

    return resp.content


@st.cache_data(ttl=300, show_spinner=False)
def load_data(url: str):
    """
    Returns (df, meta, error_msg).
    meta = {col_name: tag}  where tag ∈ {"SEARCH", "Tier 1", "Tier 2"}
    """
    try:
        content = _fetch_gdrive(url)
    except Exception as exc:
        return None, None, str(exc)

    try:
        raw = pd.read_excel(BytesIO(content), header=None, dtype=str)
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
import unicodedata as _ud

_CPF_COLS = {"CPF", "CPF RESERVA 1", "CPF RESERVA 2"}

# Map each base vowel/ç/ñ (and its accented variants) to a regex char-class
# so that searching "Antonio" also matches "Antônio" and vice-versa.
_ACCENT_CLASS: dict[str, str] = {}
for _base, _variants in [
    ("a", "aàáâãäå"), ("e", "eèéêë"), ("i", "iìíîï"),
    ("o", "oòóôõö"), ("u", "uùúûü"), ("c", "cç"),  ("n", "nñ"),
]:
    _cls = f"[{_variants}]"
    for _ch in _variants:
        _ACCENT_CLASS[_ch] = _cls


def _normalize(s: str) -> str:
    """Lowercase + strip all combining diacritical marks."""
    return "".join(
        c for c in _ud.normalize("NFD", s) if _ud.category(c) != "Mn"
    ).lower()


def _strip_cpf(s: str) -> str:
    """Remove dots and dashes so raw and formatted CPFs compare equal."""
    return s.replace(".", "").replace("-", "")


def search_numero_exact(query: str) -> pd.DataFrame:
    """Return only the row whose NUMERO exactly matches the query."""
    q = query.strip()
    if not q or df is None or "NUMERO" not in df.columns:
        return pd.DataFrame()
    return df[df["NUMERO"].fillna("").str.strip() == q].copy()


def search_df(query: str, cols: list[str] | None = None) -> pd.DataFrame:
    q_norm = _normalize(query.strip())
    if not q_norm or df is None:
        return pd.DataFrame()
    if cols is None:
        cols = SEARCH_COLS
    q_cpf = _strip_cpf(q_norm)
    mask = pd.Series([False] * len(df), index=df.index)
    for col in cols:
        if col not in df.columns:
            continue
        if col in _CPF_COLS:
            mask |= (
                df[col].fillna("").apply(lambda v: _strip_cpf(_normalize(v)))
                .str.contains(q_cpf, regex=False)
            )
        else:
            mask |= (
                df[col].fillna("").apply(_normalize)
                .str.contains(q_norm, regex=False)
            )
    return df[mask].copy()


# ──────────────────────────────────────────────────────────────
#  RENDER RECORD
# ──────────────────────────────────────────────────────────────
def cell(row: pd.Series, col: str) -> str:
    v = str(row.get(col, "")).strip()
    return "" if v in ("nan", "None", "") else v


import re as _re


def _query_to_pattern(query: str) -> str:
    """Build a regex pattern from query that matches accented variants."""
    parts = []
    for ch in query.lower():
        if ch in _ACCENT_CLASS:
            parts.append(_ACCENT_CLASS[ch])
        else:
            parts.append(_re.escape(ch))
    return "".join(parts)


def highlight(text: str, query: str) -> str:
    """Highlight query in text, accent-insensitive."""
    if not query or not text:
        return text
    try:
        pattern = _query_to_pattern(query)
        return _re.sub(f"({pattern})", r'<mark class="hl">\1</mark>', text, flags=_re.IGNORECASE)
    except _re.error:
        return text


def highlight_cpf(text: str, query: str) -> str:
    """Highlight a CPF field, matching regardless of dots/dashes formatting."""
    if not query or not text:
        return text
    result = highlight(text, query)
    if result != text:
        return result
    if _strip_cpf(_normalize(query)) in _strip_cpf(_normalize(text)):
        return f'<mark class="hl">{text}</mark>'
    return text


def sort_results(results: pd.DataFrame, query: str) -> pd.DataFrame:
    """Push rows where NOME or CPF match the query to the top."""
    if results.empty or not query:
        return results
    q_norm = _normalize(query.strip())
    q_cpf  = _strip_cpf(q_norm)

    def priority(row: pd.Series) -> int:
        if q_norm in _normalize(str(row.get("NOME", ""))):
            return 0
        if q_cpf in _strip_cpf(_normalize(str(row.get("CPF", "")))):
            return 0
        return 1

    results = results.copy()
    results["_pri"] = results.apply(priority, axis=1)
    return results.sort_values("_pri", kind="stable").drop(columns=["_pri"])


def render_record(row: pd.Series, query: str = "", numero_query: str = "") -> None:
    q     = query.strip().lower()
    q_num = numero_query.strip()   # used only for NUMERO header highlight

    nome   = cell(row, "NOME")
    numero = cell(row, "NUMERO")
    cpf    = cell(row, "CPF")
    tipo   = cell(row, "TIPO")
    status = cell(row, "STATUS")
    cid    = cell(row, "CID 2026")
    alerta = cell(row, "ALERTA PARA A MESA")
    reserva1     = cell(row, "RESERVA 1")
    cpf_reserva1 = cell(row, "CPF RESERVA 1")
    reserva2     = cell(row, "RESERVA 2")
    cpf_reserva2 = cell(row, "CPF RESERVA 2")

    status_cls = "b-ativo" if status.lower() == "ativo" else "b-inativo"

    # ── Header: name + number + CPF
    # NUMERO highlights via numero_query (NUMERO-bar search) or general query
    numero_hl = highlight(numero, q_num) if q_num else highlight(numero, q)
    html = f"""
    <div class="card">
      <div class="card-name">{highlight(nome, q) or "—"}</div>
      <div class="card-num">Nº {numero_hl} &nbsp;·&nbsp; CPF {highlight_cpf(cpf, q) if cpf else "—"}</div>
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

    def reserva_row(label: str, val: str, is_cpf: bool = False) -> str:
        if val:
            v_hl = highlight_cpf(val, q) if is_cpf else highlight(val, q)
            matched = (
                (_strip_cpf(q) in _strip_cpf(val.lower())) if is_cpf
                else (q in val.lower())
            )
        else:
            v_hl = '<span class="info-empty">—</span>'
            matched = False
        row_cls = ' info-alert-row' if matched else ''
        return f'''
      <div class="info-row{row_cls}">
        <span class="info-label">{label}</span>
        <span class="info-value">{v_hl}</span>
      </div>'''

    html += reserva_row("Reserva 1",   reserva1)
    html += reserva_row("CPF Res. 1",  cpf_reserva1, is_cpf=True)
    html += reserva_row("Reserva 2",   reserva2)
    html += reserva_row("CPF Res. 2",  cpf_reserva2, is_cpf=True)

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


def show_results(
    num_results: pd.DataFrame,
    gen_results: pd.DataFrame,
    num_query: str,
    gen_query: str,
) -> None:
    # OR logic: union of both result sets; gen_results take highlight priority
    gen_idx = set(gen_results.index) if not gen_results.empty else set()
    num_idx = set(num_results.index) if not num_results.empty else set()
    all_idx = gen_idx | num_idx

    if not all_idx:
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
        return

    # Preserve original df order; sort within that by NOME/CPF priority
    combined = df.loc[sorted(all_idx)].copy()
    # Use gen_query for sorting if available, else num_query
    combined = sort_results(combined, gen_query or num_query)

    count = len(combined)
    st.markdown(
        f'<div class="result-count">{count} registro{"s" if count > 1 else ""} encontrado{"s" if count > 1 else ""}</div>',
        unsafe_allow_html=True,
    )
    for idx, row in combined.iterrows():
        if idx in gen_idx:
            # General match: highlight all matching fields
            render_record(row, query=gen_query, numero_query="")
        else:
            # NUMERO-only match: highlight only the NUMERO header
            render_record(row, query="", numero_query=num_query)


# ──────────────────────────────────────────────────────────────
#  UI — RESULTS
# ──────────────────────────────────────────────────────────────
num_results = search_numero_exact(query_num) if query_num else pd.DataFrame()
gen_results = search_df(query_gen)           if query_gen else pd.DataFrame()

if query_num or query_gen:
    show_results(num_results, gen_results, query_num, query_gen)
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
