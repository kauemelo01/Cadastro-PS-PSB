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
import json
import threading
import queue
import time
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
    .info-alerta-row { background: #3e1010 !important; }
    .m-no { background: #2a352a !important; }
    .kid { background: #2e2545 !important; color: #d1c4e9 !important; }
    .kid-age { color: #b39ddb !important; }
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
  .info-alerta-row { background: #ffebee; border-radius: 6px; padding: 5px 7px; margin: 2px 0; }
  .info-alerta-val { color: #b71c1c; font-weight: 600; font-size: 0.97rem; }

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

  /* ── Crianças ── */
  .kids { display: flex; flex-wrap: wrap; gap: 5px; margin-top: 4px; }
  .kid {
    background: #ede7f6;
    color: #4527a0;
    border-radius: 11px;
    padding: 3px 9px;
    font-size: 0.8rem;
    font-weight: 600;
    white-space: nowrap;
  }
  .kid-age {
    color: #7e57c2;
    font-weight: 500;
    font-size: 0.74rem;
  }

  /* ── Admin footer (discreet) ── */
  .admin-footer { margin-top: 2.5rem; }
  .admin-badge {
    font-size: 0.78rem;
    color: #2e7d32;
    font-weight: 700;
    padding-top: 8px;
  }
  /* Render the admin link as plain faint text, not a button */
  .st-key-admin_login_link button {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    color: #c5c5c5 !important;
    font-size: 0.72rem !important;
    font-weight: 400 !important;
    padding: 2px 0 !important;
    min-height: 0 !important;
  }
  .st-key-admin_login_link button:hover {
    color: #388e3c !important;
    text-decoration: underline;
  }
  .st-key-admin_login_link button p {
    font-size: 0.72rem !important;
    font-weight: 400 !important;
  }
  .st-key-admin_logout button {
    font-size: 0.75rem !important;
    padding: 2px 8px !important;
    min-height: 0 !important;
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
def header_display_name(value: object) -> str:
    """
    Spreadsheet header → the name the app uses for that column.

    The monthly columns are real dates in the file, which the app shows as
    "Mar/26". Any lookup by header name has to apply the same conversion, or
    it will never match the month columns. Handles both the datetime objects
    openpyxl returns and the strings pandas produces.
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return f"{_MONTHS_PT[value.month - 1]}/{str(value.year)[-2:]}"
    raw = str(value).strip()
    if not raw:
        return ""
    if is_numeric(raw):
        return serial_to_month(raw)
    try:
        dt = pd.to_datetime(raw, errors="raise")
        return f"{_MONTHS_PT[dt.month - 1]}/{str(dt.year)[-2:]}"
    except Exception:
        return raw


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
    Returns (df, meta, col_index_map, error_msg).
    meta          = {col_name: tag}   tag ∈ {"SEARCH", "Tier 1", "Tier 2"}
    col_index_map = {col_name: 0-based xlsx column index}
    """
    try:
        content = _fetch_gdrive(url)
    except Exception as exc:
        return None, None, {}, str(exc)

    try:
        raw = pd.read_excel(BytesIO(content), sheet_name=0, header=None, dtype=str)
    except Exception as exc:
        return None, None, {}, str(exc)

    if raw.shape[0] < 3:
        return None, None, {}, "Arquivo sem dados suficientes."

    tags_row    = raw.iloc[0].tolist()   # Row 0: SEARCH / Tier 1 / Tier 2
    names_row   = raw.iloc[1].tolist()   # Row 1: human-readable column names

    # Build deduplicated column names
    seen: dict[str, int] = {}
    final_names: list[str] = []
    final_tags:  list[str] = []

    for tag, name in zip(tags_row, names_row):
        raw_name = str(name).strip() if pd.notna(name) else ""
        tag_str  = str(tag).strip()  if pd.notna(tag)  else ""

        # Same conversion the write path uses, so names match on both sides.
        display_name = header_display_name(raw_name)

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
    # col_index_map: display_name → 0-based xlsx column index
    col_index_map = {name: idx for idx, name in enumerate(final_names)}
    return df, meta, col_index_map, None


# ──────────────────────────────────────────────────────────────
#  LOAD DATA
# ──────────────────────────────────────────────────────────────
with st.spinner("Carregando cadastro…"):
    df, meta, col_index_map, load_error = load_data(GITHUB_RAW_URL)


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
#  SESSION STATE — in-memory edits overlay
# ──────────────────────────────────────────────────────────────
if "edits" not in st.session_state:
    st.session_state.edits = {}  # {numero: {col: value}}

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False


# Admin password — read from secrets only. If unset, admin access is disabled.
try:
    ADMIN_PASSWORD = st.secrets["ADMIN_PASSWORD"]
except Exception:
    ADMIN_PASSWORD = None


def apply_edits(row: pd.Series) -> pd.Series:
    """
    Return a copy of row with any in-session edits applied.

    Keyed by NUMERO, never by row position: the frame is reloaded whenever the
    cache expires, and a row inserted or removed in the spreadsheet shifts
    every position below it — the overlay would then paint one person's edits
    onto whoever moved into that slot.
    """
    numero = str(row.get("NUMERO", "")).strip()
    overlay = st.session_state.edits.get(numero, {}) if numero else {}
    if overlay:
        row = row.copy()
        for col, val in overlay.items():
            row[col] = val
    return row


def current_month_col() -> str | None:
    now = datetime.now()
    col = f"{_MONTHS_PT[now.month - 1]}/{str(now.year)[-2:]}"
    return col if col in MONTH_COLS else None


# ──────────────────────────────────────────────────────────────
#  GOOGLE DRIVE WRITE-BACK
# ──────────────────────────────────────────────────────────────
def _gdrive_write_enabled() -> bool:
    """True when both Drive secrets are configured."""
    return (
        "GDRIVE_FILE_ID" in st.secrets
        and "GDRIVE_SA_CREDENTIALS" in st.secrets
    )


def _build_drive_service(sa_info: dict):
    """
    Build a Drive service from an already-extracted credentials dict.
    Takes no Streamlit state, so it is safe to call from a background thread.
    """
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    creds = Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _drive_secrets() -> tuple[str, dict] | None:
    """
    Read the Drive secrets. MUST be called from the main thread — background
    threads have no Streamlit context and cannot touch st.secrets.
    Returns (file_id, sa_info) or None when not configured.
    """
    if not _gdrive_write_enabled():
        return None
    try:
        return (
            st.secrets["GDRIVE_FILE_ID"],
            json.loads(st.secrets["GDRIVE_SA_CREDENTIALS"]),
        )
    except Exception:
        return None


def _get_drive_service():
    """Main-thread convenience wrapper."""
    sec = _drive_secrets()
    if sec is None:
        raise RuntimeError("Credenciais do Google Drive não configuradas.")
    return _build_drive_service(sec[1])


@st.cache_resource(show_spinner=False)
def get_drive_lock():
    """
    Process-wide lock held for the whole download → modify → upload cycle.

    Every writer overwrites the ENTIRE file, so two cycles running at once
    means the slower upload silently reverts the faster one: a delivery, an
    edit or a whole record just disappears. Serialising them is what makes
    concurrent use by several admins safe.
    """
    return threading.RLock()


def _find_numero_column(ws) -> int | None:
    """Locate the NUMERO column from the live header row (row 2)."""
    for cell in ws[2]:
        if cell.value is not None and \
           str(cell.value).strip().upper() in ("NUMERO", "NÚMERO"):
            return cell.column
    return None


def _find_row_by_numero(ws, numero: str, numero_col: int) -> list[int]:
    """All data rows whose NUMERO matches. Read from the file, never cached."""
    hits = []
    for cells in ws.iter_rows(min_row=3, min_col=numero_col, max_col=numero_col):
        value = cells[0].value
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        if value is not None and str(value).strip() == str(numero).strip():
            hits.append(cells[0].row)
    return hits


def _download_workbook(service, file_id):
    """Download the xlsx from Drive and open it with openpyxl."""
    from googleapiclient.http import MediaIoBaseDownload
    import openpyxl
    req = service.files().get_media(fileId=file_id)
    buf = BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return openpyxl.load_workbook(buf)


def _upload_workbook(service, file_id, wb) -> None:
    """Serialise the workbook and overwrite the Drive file."""
    from googleapiclient.http import MediaIoBaseUpload
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    media = MediaIoBaseUpload(
        buf,
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
        resumable=False,
    )
    service.files().update(fileId=file_id, media_body=media).execute()


def save_cells_to_drive(numero: str, changes: dict[str, str]) -> str | None:
    """
    Update one or more cells of a single registration, in one roundtrip.
    Returns None on success, or an error string on failure.

    The row is located by NUMERO and the columns by their header names, both
    read from the freshly downloaded file. Addressing them by cached position
    would write into whoever happens to sit at that row now: the cached frame
    can be up to 5 minutes old, and a row inserted or removed in the
    spreadsheet meanwhile shifts every record below it.
    """
    if not changes:
        return None
    numero = str(numero).strip()
    if not numero:
        return "O cadastro não possui NÚMERO; alteração não salva."

    clicked_at = datetime.now()     # before the lock and the network
    try:
        file_id = st.secrets["GDRIVE_FILE_ID"]
        service = _get_drive_service()

        with get_drive_lock():
            wb = _download_workbook(service, file_id)
            try:
                ws = wb.worksheets[0]   # main data always on the first sheet

                numero_col = _find_numero_column(ws)
                if numero_col is None:
                    return "Coluna NÚMERO não encontrada. Nada foi alterado."

                hits = _find_row_by_numero(ws, numero, numero_col)
                if not hits:
                    return (f"Cadastro Nº {numero} não encontrado no arquivo "
                            "atual. Nada foi alterado.")
                if len(hits) > 1:
                    return (f"Há mais de um cadastro com o Nº {numero}. "
                            "Nada foi alterado.")
                target_row = hits[0]

                # Resolve columns from the live header row, not from cache.
                headers = {}
                for cell in ws[2]:
                    name = header_display_name(cell.value)
                    if name:
                        headers.setdefault(name, cell.column)

                missing = [c for c in changes if c not in headers]
                if missing:
                    return ("Coluna(s) não encontrada(s) na planilha: "
                            + ", ".join(missing) + ". Nada foi alterado.")

                for col_name, value in changes.items():
                    ws.cell(row=target_row, column=headers[col_name], value=value)

                _stamp_timestamps(ws, target_row, clicked_at)
                _upload_workbook(service, file_id, wb)
            finally:
                wb.close()

        load_data.clear()
        return None

    except Exception as exc:
        return str(exc)


def save_cell_to_drive(numero: str, col_name: str, value: str) -> str | None:
    """Single-cell convenience wrapper around save_cells_to_drive."""
    return save_cells_to_drive(numero, {col_name: value})


# ──────────────────────────────────────────────────────────────
#  CRIANÇAS — JSON in column Z (Tier 2)
# ──────────────────────────────────────────────────────────────
# ── Columns the app writes to, identified by HEADER NAME (row 2) ──
#
# Each one lists the current name first and the previous name as a fallback.
# Renaming the spreadsheet and deploying this code never happen at the same
# instant, so accepting both keeps the app working throughout the switch.
# Once the spreadsheet is renamed, the second entry can be dropped.
#
# A fixed column position is deliberately NOT used: it would silently shift
# onto a neighbouring Tier 2 column (OBSERVAÇÕES INTERNAS sits right beside
# the children column) if a column were ever inserted.
CHILDREN_COL_NAMES = ["CRIANÇAS", "Última Alteração"]
TS_APP_COL_NAMES   = ["Salvo no App", "Alterado Por:"]
TS_SHEET_COL_NAMES = ["Gravado na Planilha", "Motivo da Alteração"]

CHILDREN_COL_NAME = CHILDREN_COL_NAMES[0]   # the name shown in messages
CHILDREN_XLSX_COL = 26          # fallback only, when the header is missing
EXCEL_CELL_MAX_CHARS = 32767
TS_FORMAT = "%d/%m/%Y %H:%M:%S"


def _find_col_by_header(ws, names: list[str]) -> int | None:
    """
    Column whose header (row 2) matches any of `names`, in order of preference.
    Tolerant to accents and case.
    """
    headers = [(cell.column, _normalize(header_display_name(cell.value)))
               for cell in ws[2] if cell.value is not None]
    for name in names:
        wanted = _normalize(name)
        for col, got in headers:
            if got == wanted:
                return col
    return None


def _stamp_timestamps(ws, row: int, clicked_at: datetime) -> None:
    """
    Record both moments for a change: when the user saved in the app, and when
    the spreadsheet was actually written. They differ by the queue wait plus
    the download/upload, so a gap between them shows where time was spent.
    Missing columns are skipped — a timestamp must never block the real save.
    """
    for names, moment in ((TS_APP_COL_NAMES, clicked_at),
                          (TS_SHEET_COL_NAMES, datetime.now())):
        col = _find_col_by_header(ws, names)
        if col is not None:
            ws.cell(row=row, column=col, value=moment.strftime(TS_FORMAT))


def _find_children_column(ws) -> tuple[int | None, str | None]:
    """
    Locate the children column by header name (row 2), case/accent tolerant.
    Returns (column, None) or (None, error message).
    """
    headers = [(cell.column, _normalize(header_display_name(cell.value)))
               for cell in ws[2] if cell.value is not None]

    col = None
    for name in CHILDREN_COL_NAMES:
        wanted = _normalize(name)
        hits = [c for c, got in headers if got == wanted]
        if len(hits) > 1:
            return None, (f"Há mais de uma coluna '{name}'. "
                          "Nenhum dado foi alterado.")
        if hits:
            col = hits[0]
            break

    if col is None:
        return None, (f"Coluna '{CHILDREN_COL_NAME}' não encontrada na "
                      "planilha. Nenhum dado foi alterado.")

    if str(ws.cell(row=1, column=col).value).strip() != "Tier 2":
        return None, (f"A coluna '{CHILDREN_COL_NAME}' não está identificada "
                      "como Tier 2.")
    return col, None


def append_crianca_json(existing: object, record: dict) -> str:
    """Append a child without discarding any previously stored JSON data."""
    if existing is None or (isinstance(existing, str) and not existing.strip()):
        payload = {"criancas": []}
    else:
        try:
            payload = json.loads(existing)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                "A coluna Z já contém um valor que não é um JSON válido. "
                "O conteúdo anterior foi preservado; corrija-o antes de salvar."
            ) from exc

    if isinstance(payload, list):
        # Preserve legacy JSON arrays in their original format.
        payload.append(record)
    elif isinstance(payload, dict):
        if "criancas" not in payload:
            payload["criancas"] = []
        if not isinstance(payload["criancas"], list):
            raise ValueError(
                "O campo 'criancas' do JSON existente deve ser uma lista. "
                "O conteúdo anterior foi preservado."
            )
        payload["criancas"].append(record)
    else:
        # Keep valid scalar JSON values rather than replacing their contents.
        payload = {"dados_anteriores": payload, "criancas": [record]}

    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"),
                         allow_nan=False)
    # openpyxl truncates longer strings, which would corrupt the saved JSON.
    if len(encoded) > EXCEL_CELL_MAX_CHARS:
        raise ValueError(
            "O JSON excederia o limite de 32.767 caracteres da célula do Excel. "
            "Nenhum dado foi alterado."
        )
    return encoded


SEXO_EMOJI = {"Feminino": "👧", "Masculino": "👦"}
SEXO_OPCOES = ["", "Feminino", "Masculino", "Não informado"]


def parse_criancas(raw: object) -> list[dict]:
    """Children from a raw cell value. Never raises — bad JSON reads as empty."""
    if raw is None:
        return []
    text = str(raw).strip()
    if not text or text in ("nan", "None"):
        return []
    try:
        payload = json.loads(text)
    except (TypeError, ValueError):
        return []          # never let a malformed cell break the card
    if isinstance(payload, list):
        items = payload
    elif isinstance(payload, dict):
        items = payload.get("criancas", [])
    else:
        return []
    return [c for c in items if isinstance(c, dict)] if isinstance(items, list) else []


def read_criancas(row: pd.Series) -> list[dict]:
    """Children stored for this registration, under whichever name the file uses."""
    for name in CHILDREN_COL_NAMES:
        if name in row.index:
            items = parse_criancas(row.get(name, ""))
            if items:
                return items
    return []


def criancas_fingerprint(items: list[dict]) -> str:
    """Stable signature of a children list, for optimistic concurrency checks."""
    return json.dumps(
        [{k: str(c.get(k, "")) for k in ("NOME", "SEXO", "IDADE", "OBS")}
         for c in items],
        ensure_ascii=False, sort_keys=True,
    )


def replace_criancas_json(existing: object, children: list[dict]) -> str:
    """Replace the children list while keeping anything else stored in the cell."""
    if existing is None or (isinstance(existing, str) and not existing.strip()):
        payload = {"criancas": []}
    else:
        try:
            payload = json.loads(existing)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                "A coluna já contém um valor que não é um JSON válido. "
                "O conteúdo anterior foi preservado; corrija-o antes de salvar."
            ) from exc

    if isinstance(payload, list):
        payload = list(children)          # keep legacy array format
    elif isinstance(payload, dict):
        payload["criancas"] = list(children)
    else:
        payload = {"dados_anteriores": payload, "criancas": list(children)}

    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"),
                         allow_nan=False)
    if len(encoded) > EXCEL_CELL_MAX_CHARS:
        raise ValueError(
            "O JSON excederia o limite de 32.767 caracteres da célula do Excel. "
            "Nenhum dado foi alterado."
        )
    return encoded


def save_criancas_to_drive(numero: str, children: list[dict],
                           expected: list[dict] | None = None) -> str | None:
    """
    Overwrite the whole children list for one registration.

    `expected` is the list as it was when the editor opened. Because this
    replaces the list wholesale, saving a stale copy would delete a child that
    somebody else added meanwhile — the lock cannot help, since the stale data
    is already baked into the request. So the cell is compared against
    `expected` inside the lock and the save is refused if it moved.
    """
    def transform(existing):
        if expected is not None:
            atuais = parse_criancas(existing)
            if criancas_fingerprint(atuais) != criancas_fingerprint(expected):
                raise ValueError(
                    "A lista de crianças foi alterada por outra pessoa "
                    "enquanto você editava. Nada foi alterado — feche e "
                    "reabra para ver a versão atual."
                )
        return replace_criancas_json(existing, children)

    return _write_criancas(numero, transform)


def _write_criancas(numero: str, transform) -> str | None:
    """
    Shared read-modify-write for the children cell. `transform` receives the
    current cell value and returns the new JSON string.

    The row is found by NUMERO and the column by header name, both read from
    the freshly downloaded file, and the whole cycle runs under the shared
    lock so a concurrent upload cannot revert it.
    """
    if not _gdrive_write_enabled():
        return "Credenciais do Google Drive não configuradas."
    numero = str(numero).strip()
    if not numero:
        return "O cadastro não possui NÚMERO."

    clicked_at = datetime.now()     # before the lock and the network
    try:
        file_id = st.secrets["GDRIVE_FILE_ID"]
        service = _get_drive_service()

        with get_drive_lock():
            wb = _download_workbook(service, file_id)
            try:
                ws = wb.worksheets[0]
                children_col, col_err = _find_children_column(ws)
                if col_err:
                    return col_err

                numero_col = _find_numero_column(ws)
                if numero_col is None:
                    return "Não foi possível identificar uma única coluna NÚMERO."

                matches = _find_row_by_numero(ws, numero, numero_col)
                if not matches:
                    return f"Cadastro Nº {numero} não encontrado no arquivo atual."
                if len(matches) != 1:
                    return f"Há mais de um cadastro com o Nº {numero}. Nenhum dado foi alterado."

                target = ws.cell(row=matches[0], column=children_col)
                target.value = transform(target.value)

                _stamp_timestamps(ws, matches[0], clicked_at)
                _upload_workbook(service, file_id, wb)
            finally:
                wb.close()

        load_data.clear()
        return None
    except Exception as exc:
        return str(exc)


def save_crianca_to_drive(numero: str, record: dict) -> str | None:
    """Append one child, preserving everything already stored in the cell."""
    numero = str(numero).strip()
    child = dict(record)
    # The number comes from the selected registration, never from the input.
    child["NÚMERO"] = numero
    return _write_criancas(numero,
                           lambda existing: append_crianca_json(existing, child))


# ──────────────────────────────────────────────────────────────
#  CESTAS EXTRAS (Sheet2)
# ──────────────────────────────────────────────────────────────
EXTRAS_SHEET   = "Sheet2"
EXTRAS_HEADERS = ["NÚMERO", "NOME", "CPF", "TELEFONE", "MOTIVO"]
EXTRAS_DONE_COL    = 6          # column F
EXTRAS_DONE_HEADER = "ENTREGUE"


def _read_extras_from_drive(url: str) -> list[dict]:
    """Read Sheet2 from Drive and return it as a list of plain dicts."""
    content = _fetch_gdrive(url)
    extras = pd.read_excel(
        BytesIO(content), sheet_name=EXTRAS_SHEET, dtype=str
    ).fillna("")

    # Normalise the delivered column (F). It may be missing entirely, or
    # present without a header, in which case pandas names it "Unnamed: 5".
    if EXTRAS_DONE_HEADER not in extras.columns:
        if len(extras.columns) >= EXTRAS_DONE_COL:
            extras = extras.rename(
                columns={extras.columns[EXTRAS_DONE_COL - 1]: EXTRAS_DONE_HEADER}
            )
        else:
            extras[EXTRAS_DONE_HEADER] = ""

    # Drop rows with no NÚMERO and no NOME (blank spacer rows)
    if "NÚMERO" in extras.columns and "NOME" in extras.columns:
        keep = (extras["NÚMERO"].str.strip() != "") | (extras["NOME"].str.strip() != "")
        extras = extras[keep]

    cols = EXTRAS_HEADERS + [EXTRAS_DONE_HEADER]
    for c in cols:
        if c not in extras.columns:
            extras[c] = ""
    return [
        {c: str(r.get(c, "")).strip() for c in cols}
        for _, r in extras.iterrows()
    ]


# ──────────────────────────────────────────────────────────────
#  SHARED STORE
#
#  st.cache_resource is shared by every session in the server
#  process, so a write by one user is visible to every other
#  instance on its next render — with no Drive read involved.
#  Drive stays the system of record, written asynchronously by a
#  single background worker so the UI never blocks on the network.
# ──────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_extras_store() -> dict:
    return {
        "lock":    threading.RLock(),
        "rows":    [],      # list[dict] — the shared truth shown in the UI
        "loaded":  False,   # hydrated from Drive at least once?
        "version": 0,       # bumped on every change
        "queue":   queue.Queue(),
        "pending": 0,       # writes still in flight
        "failed":  [],      # jobs that could not be written (retryable)
        "stale":   0,       # records not found in the sheet (NOT retryable)
        "errors":  [],      # human-readable error strings
        "worker":  False,
        "synced":  None,    # timestamp of last successful Drive write
        "file_id": None,
        "sa_info": None,
    }


# ── Workbook mutations (plain Python — no Streamlit, thread-safe) ──
def _ws_extras(wb):
    """Get or create Sheet2, ensuring the header row is complete."""
    if EXTRAS_SHEET in wb.sheetnames:
        ws = wb[EXTRAS_SHEET]
    else:
        ws = wb.create_sheet(EXTRAS_SHEET)
    for col, header in enumerate(EXTRAS_HEADERS, start=1):
        if not ws.cell(row=1, column=col).value:
            ws.cell(row=1, column=col, value=header)
    if not ws.cell(row=1, column=EXTRAS_DONE_COL).value:
        ws.cell(row=1, column=EXTRAS_DONE_COL, value=EXTRAS_DONE_HEADER)
    return ws


def _rows_with_numero(ws, numero: str) -> list[int]:
    """Rows in the extras sheet whose NÚMERO matches."""
    hits = []
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val is not None and str(val).strip() == str(numero).strip():
            hits.append(r)
    return hits


def _apply_append(ws, record: dict) -> str | None:
    """
    Write a new row at the first genuinely empty line.
    Refuses if the number already exists, so the app never creates a duplicate
    that a later delivery could flag on the wrong row.
    """
    numero = str(record.get("NÚMERO", "")).strip()
    if numero and _rows_with_numero(ws, numero):
        return (f"Nº {numero} já existe na aba {EXTRAS_SHEET}; "
                "a cesta não foi gravada para não duplicar.")

    target = ws.max_row + 1
    for r in range(2, ws.max_row + 2):
        if all(ws.cell(row=r, column=c).value in (None, "")
               for c in range(1, len(EXTRAS_HEADERS) + 1)):
            target = r
            break
    for col, header in enumerate(EXTRAS_HEADERS, start=1):
        ws.cell(row=target, column=col, value=record.get(header, ""))
    if str(record.get(EXTRAS_DONE_HEADER, "")).strip():
        ws.cell(row=target, column=EXTRAS_DONE_COL,
                value=record[EXTRAS_DONE_HEADER])
    return None


def _apply_deliver(ws, numero: str) -> str | None:
    """Flag column F for the row whose NÚMERO matches. Returns an error or None."""
    hits = _rows_with_numero(ws, numero)
    if not hits:
        return f"Nº {numero} não encontrado na aba {EXTRAS_SHEET}."
    if len(hits) > 1:
        # Flagging just the first would leave the others stuck as pending
        # forever, with no way to tell which one was actually delivered.
        return (f"Há {len(hits)} linhas com o Nº {numero} na aba "
                f"{EXTRAS_SHEET}. Nada foi marcado — corrija a planilha.")
    ws.cell(row=hits[0], column=EXTRAS_DONE_COL, value="OK")
    return None


def _rollback_job(store: dict, job: dict) -> bool:
    """
    Undo a job's optimistic in-memory change after the sheet refused it.
    Caller must hold store["lock"]. Returns True if something was undone.

    The point is to keep the screen honest: whatever the spreadsheet rejected
    must stop looking saved, so the operator can see it and act again.
    """
    if job["op"] == "deliver":
        numero = str(job["numero"]).strip()
        for r in store["rows"]:
            if r.get("NÚMERO", "").strip() == numero:
                r[EXTRAS_DONE_HEADER] = job.get("prev", "")
                return True
        return False

    if job["op"] == "append":
        record = job["record"]
        for i, r in enumerate(store["rows"]):
            # Identity first: the queued record IS the dict that was appended.
            # A re-sync may have replaced the list, in which case there is
            # nothing of ours left to undo.
            if r is record:
                store["rows"].pop(i)
                return True
        return False

    return False


def _flush_jobs(store: dict, jobs: list[dict]) -> None:
    """
    Apply a batch of jobs to the Drive file in ONE download/upload cycle.
    Raises on network failure so the caller can mark the batch as failed.
    """
    service = _build_drive_service(store["sa_info"])
    file_id = store["file_id"]

    soft_errors: list[tuple[dict, str]] = []
    # Same lock as the synchronous writers: an edit saved from another screen
    # mid-cycle would otherwise be reverted by this upload, or revert it.
    with get_drive_lock():
        wb = _download_workbook(service, file_id)
        try:
            ws = _ws_extras(wb)
            for job in jobs:
                if job["op"] == "append":
                    err = _apply_append(ws, job["record"])
                    if err:
                        soft_errors.append((job, err))
                elif job["op"] == "deliver":
                    err = _apply_deliver(ws, job["numero"])
                    if err:
                        soft_errors.append((job, err))
            _upload_workbook(service, file_id, wb)
        finally:
            wb.close()

    # These are not network failures — the upload succeeded, but a record
    # could not be located in the sheet. Retrying would fail identically, so
    # they are counted separately and always surfaced rather than queued.
    #
    # The optimistic in-memory change MUST be undone here. Without the
    # rollback the screen would keep showing a cesta as delivered while the
    # spreadsheet has nothing — and the Confirmar button would be gone, so
    # nobody could try again or even notice which record failed.
    if soft_errors:
        with store["lock"]:
            for job, err in soft_errors:
                if not _rollback_job(store, job):
                    store["errors"].append(err)
                elif job["op"] == "append":
                    store["errors"].append(err + " A cesta saiu da lista.")
                else:
                    store["errors"].append(err + " A marcação foi desfeita na tela.")
            store["stale"] += len(soft_errors)
            store["version"] += 1


def _extras_worker(store: dict) -> None:
    """
    Single writer. Serialising every Drive write through one thread is what
    prevents two concurrent download-modify-upload cycles from overwriting
    each other — last-write-wins would silently drop a record.
    """
    q: queue.Queue = store["queue"]
    while True:
        try:
            job = q.get()
            batch = [job]
            # Coalesce everything already queued into the same roundtrip.
            try:
                while True:
                    batch.append(q.get_nowait())
            except queue.Empty:
                pass

            try:
                _flush_jobs(store, batch)
                with store["lock"]:
                    store["pending"] = max(0, store["pending"] - len(batch))
                    store["synced"] = time.time()
            except Exception as exc:
                with store["lock"]:
                    store["pending"] = max(0, store["pending"] - len(batch))
                    store["failed"].extend(batch)
                    store["errors"].append(
                        f"{datetime.now():%H:%M} — falha ao gravar "
                        f"{len(batch)} alteração(ões): {exc}"
                    )
            finally:
                for _ in batch:
                    q.task_done()
        except Exception:
            time.sleep(1)   # never let the worker die


def _ensure_worker(store: dict) -> None:
    with store["lock"]:
        if store["worker"]:
            return
        store["worker"] = True
    t = threading.Thread(target=_extras_worker, args=(store,), daemon=True)
    t.start()


def _enqueue(store: dict, job: dict) -> None:
    """Queue a Drive write. No-op (in-memory only) when Drive is unconfigured."""
    if not store["file_id"]:
        return
    _ensure_worker(store)
    with store["lock"]:
        store["pending"] += 1
    store["queue"].put(job)


# ── Public API used by the UI ──
def extras_init(force: bool = False) -> str | None:
    """
    Hydrate the shared store from Drive. Runs once per server process
    (the slow path), or again when the user explicitly asks to re-sync.
    Returns an error/refusal message, or None on success.
    """
    store = get_extras_store()

    sec = _drive_secrets()
    with store["lock"]:
        if sec:
            store["file_id"], store["sa_info"] = sec
        already  = store["loaded"]
        inflight = store["pending"]
    if already and not force:
        return None

    # Re-reading Drive replaces the shared rows wholesale. Doing that while a
    # write is still queued would drop the new record from every screen until
    # the next re-sync, even though it does reach Drive moments later.
    if force and inflight:
        return (
            f"Há {inflight} alteração(ões) ainda sendo gravada(s). "
            "Aguarde a sincronização terminar antes de recarregar."
        )

    try:
        rows = _read_extras_from_drive(GITHUB_RAW_URL)
    except Exception as exc:
        rows = None
        with store["lock"]:
            store["errors"].append(f"{datetime.now():%H:%M} — leitura do Drive: {exc}")

    with store["lock"]:
        if rows is not None:
            store["rows"] = rows
            store["version"] += 1
        store["loaded"] = True
    return None


def extras_snapshot() -> list[dict]:
    """A copy of the shared rows. Pure memory read — no network."""
    store = get_extras_store()
    with store["lock"]:
        return [dict(r) for r in store["rows"]]


def extras_status() -> dict:
    store = get_extras_store()
    with store["lock"]:
        return {
            "pending": store["pending"],
            "failed":  len(store["failed"]),
            "stale":   store["stale"],
            "errors":  list(store["errors"][-5:]),
            "synced":  store["synced"],
            "version": store["version"],
            "drive":   bool(store["file_id"]),
        }


def extras_next_numero() -> int:
    """Next sequential number, reserved against the shared in-memory rows."""
    store = get_extras_store()
    with store["lock"]:
        best = 0
        for r in store["rows"]:
            try:
                best = max(best, int(str(r.get("NÚMERO", "")).strip()))
            except (ValueError, TypeError):
                continue
        return best + 1


def extras_add(record: dict) -> str | None:
    """
    Add a Cesta Extra. Visible to every instance at once; Drive write is queued.
    Returns an error string if the number is already taken.
    """
    store = get_extras_store()
    row = {c: str(record.get(c, "")).strip()
           for c in EXTRAS_HEADERS + [EXTRAS_DONE_HEADER]}

    with store["lock"]:
        # Two admins can open the dialog at the same time and be offered the
        # same number, so the collision is checked at save time, not at open.
        if any(r.get("NÚMERO", "").strip() == row["NÚMERO"]
               for r in store["rows"]):
            best = 0
            for r in store["rows"]:
                try:
                    best = max(best, int(str(r.get("NÚMERO", "")).strip()))
                except (ValueError, TypeError):
                    continue
            return (
                f"O Nº {row['NÚMERO']} já foi usado por outra pessoa. "
                f"Use o Nº {best + 1} e salve novamente "
                "(os demais campos foram mantidos)."
            )
        store["rows"].append(row)
        store["version"] += 1

    _enqueue(store, {"op": "append", "record": row})
    return None


def extras_mark_delivered(numero: str) -> None:
    """Flag a record as delivered in memory, then queue the Drive write."""
    store = get_extras_store()
    prev = ""
    with store["lock"]:
        for r in store["rows"]:
            if r.get("NÚMERO", "").strip() == str(numero).strip():
                # Kept so the flag can be undone if the sheet refuses it.
                prev = r.get(EXTRAS_DONE_HEADER, "")
                r[EXTRAS_DONE_HEADER] = "OK"
                break
        store["version"] += 1

    _enqueue(store, {"op": "deliver", "numero": str(numero).strip(),
                     "prev": prev})


def extras_retry_failed() -> int:
    """Re-queue every write that previously failed."""
    store = get_extras_store()
    with store["lock"]:
        jobs = store["failed"]
        store["failed"] = []
        store["errors"] = []
        store["stale"]  = 0
    for job in jobs:
        _enqueue(store, job)
    return len(jobs)


# ──────────────────────────────────────────────────────────────
#  DIALOGS
# ──────────────────────────────────────────────────────────────
@st.dialog("Cadastrar Criança")
def cadastrar_crianca_dialog(row_idx: int, nome: str, numero: str) -> None:
    st.markdown(f"Cadastro de **{nome}**")
    form_version = st.session_state.get("crianca_form_version", 0)
    prefix = f"crianca_{row_idx}_{form_version}"
    st.text_input("NÚMERO", value=numero, disabled=True, key=f"{prefix}_numero")
    child_name = st.text_input("NOME", key=f"{prefix}_nome")
    child_sex = st.selectbox(
        "SEXO", ["", "Feminino", "Masculino", "Não informado"],
        key=f"{prefix}_sexo",
    )
    child_age = st.number_input(
        "IDADE", min_value=0, value=None, step=1,
        help="Idade em anos completos.", key=f"{prefix}_idade",
    )
    child_obs = st.text_area(
        "OBS. (opcional)", height=90, key=f"{prefix}_obs",
    )

    st.write("")
    col1, col2 = st.columns(2)
    save_clicked = col1.button("💾 Salvar", use_container_width=True, type="primary",
                               key=f"{prefix}_salvar")
    cancel_clicked = col2.button("❌ Cancelar", use_container_width=True,
                                 key=f"{prefix}_cancelar")
    if cancel_clicked:
        st.rerun()

    if save_clicked:
        errors = []
        if not numero:
            errors.append("O cadastro não possui NÚMERO.")
        if not child_name.strip():
            errors.append("NOME é obrigatório.")
        if not child_sex:
            errors.append("SEXO é obrigatório.")
        if child_age is None:
            errors.append("IDADE é obrigatória.")
        if errors:
            for error in errors:
                st.error(error)
            st.stop()

        record = {
            "NÚMERO": numero,
            "NOME": child_name.strip(),
            "SEXO": child_sex,
            "IDADE": int(child_age),
            "OBS": child_obs.strip(),
        }
        with st.spinner("Salvando no Google Drive…"):
            err = save_crianca_to_drive(numero, record)
        if err:
            st.error(f"Erro ao salvar: {err}")
            st.stop()

        st.session_state.crianca_saved = (
            f"Criança {record['NOME']} cadastrada no cadastro Nº {numero}."
        )
        st.rerun()


@st.dialog("Editar Crianças")
def editar_criancas_dialog(row_idx: int, nome: str, numero: str,
                           children: list[dict]) -> None:
    st.markdown(f"Crianças de **{nome}** &nbsp;·&nbsp; Nº {numero}",
                unsafe_allow_html=True)

    form_version = st.session_state.get("criancas_form_version", 0)
    prefix = f"edcri_{row_idx}_{form_version}"

    kept: list[dict] = []
    removidas = 0

    for i, c in enumerate(children):
        c_nome  = str(c.get("NOME", "")).strip()
        c_sexo  = str(c.get("SEXO", "")).strip()
        emoji   = SEXO_EMOJI.get(c_sexo, "🧒")
        try:
            c_idade = int(float(c.get("IDADE")))
        except (TypeError, ValueError):
            c_idade = None

        rotulo = f"{emoji} {c_nome or 'Sem nome'}"
        if c_idade is not None:
            rotulo += f" — {c_idade} ano{'s' if c_idade != 1 else ''}"

        with st.expander(rotulo, expanded=len(children) == 1):
            remover = st.checkbox("🗑️ Remover esta criança",
                                  key=f"{prefix}_rm{i}")
            n = st.text_input("NOME", value=c_nome,
                              key=f"{prefix}_n{i}", disabled=remover)
            s = st.selectbox(
                "SEXO", SEXO_OPCOES,
                index=SEXO_OPCOES.index(c_sexo) if c_sexo in SEXO_OPCOES else 0,
                key=f"{prefix}_s{i}", disabled=remover,
            )
            a = st.number_input("IDADE", min_value=0, step=1,
                                value=c_idade, key=f"{prefix}_a{i}",
                                disabled=remover)
            o = st.text_area("OBS. (opcional)",
                             value=str(c.get("OBS", "") or ""), height=70,
                             key=f"{prefix}_o{i}", disabled=remover)

        if remover:
            removidas += 1
            continue
        # Preserve any other keys the record may carry
        novo = dict(c)
        novo.update({"NÚMERO": numero, "NOME": n.strip(), "SEXO": s,
                     "IDADE": a, "OBS": o.strip()})
        kept.append(novo)

    st.divider()
    with st.expander("➕ Adicionar outra criança", expanded=not children):
        nv_nome = st.text_input("NOME", key=f"{prefix}_nvn")
        nv_sexo = st.selectbox("SEXO", SEXO_OPCOES, key=f"{prefix}_nvs")
        nv_idade = st.number_input("IDADE", min_value=0, value=None, step=1,
                                   key=f"{prefix}_nva")
        nv_obs = st.text_area("OBS. (opcional)", height=70, key=f"{prefix}_nvo")

    nova = None
    if nv_nome.strip() or nv_sexo or nv_idade is not None:
        nova = {"NÚMERO": numero, "NOME": nv_nome.strip(), "SEXO": nv_sexo,
                "IDADE": nv_idade, "OBS": nv_obs.strip()}

    st.write("")
    col1, col2 = st.columns(2)
    salvar   = col1.button("💾 Salvar", use_container_width=True,
                           type="primary", key=f"{prefix}_ok")
    cancelar = col2.button("❌ Cancelar", use_container_width=True,
                           key=f"{prefix}_no")

    if cancelar:
        st.rerun()

    if salvar:
        errors: list[str] = []
        for idx, c in enumerate(kept, start=1):
            if not c["NOME"]:
                errors.append(f"Criança {idx}: NOME é obrigatório.")
            if not c["SEXO"]:
                errors.append(f"Criança {idx}: SEXO é obrigatório.")
            if c["IDADE"] is None:
                errors.append(f"Criança {idx}: IDADE é obrigatória.")
        if nova is not None:
            if not nova["NOME"]:
                errors.append("Nova criança: NOME é obrigatório.")
            if not nova["SEXO"]:
                errors.append("Nova criança: SEXO é obrigatório.")
            if nova["IDADE"] is None:
                errors.append("Nova criança: IDADE é obrigatória.")
        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        final = [dict(c, IDADE=int(c["IDADE"])) for c in kept]
        if nova is not None:
            final.append(dict(nova, IDADE=int(nova["IDADE"])))

        with st.spinner("Salvando no Google Drive…"):
            # `children` is the list this dialog opened with — the save is
            # refused if somebody changed it in the meantime.
            err = save_criancas_to_drive(numero, final, expected=children)
        if err:
            st.error(f"Erro ao salvar: {err}")
            st.stop()

        partes = [f"{len(final)} criança(s) salva(s)"]
        if removidas:
            partes.append(f"{removidas} removida(s)")
        st.session_state.crianca_saved = (
            f"Cadastro Nº {numero}: " + ", ".join(partes) + "."
        )
        st.rerun()


@st.dialog("Confirmar entrega")
def confirm_month_dialog(row_idx: int, month_col: str, nome: str, numero: str) -> None:
    st.markdown(
        f"Confirmar entrega de **{month_col}** para:<br>"
        f"<b>{nome}</b> &nbsp;·&nbsp; Nº {numero}",
        unsafe_allow_html=True,
    )
    st.write("")
    col1, col2 = st.columns(2)
    if col1.button("✅ Confirmar", use_container_width=True, type="primary"):
        # Persist first: the overlay must not claim a save that did not happen.
        if _gdrive_write_enabled():
            with st.spinner("Salvando no Google Drive…"):
                err = save_cell_to_drive(numero, month_col, "OK")
            if err:
                st.error(f"Erro ao salvar: {err}")
                st.stop()
        st.session_state.edits.setdefault(numero, {})[month_col] = "OK"
        st.rerun()
    if col2.button("❌ Cancelar", use_container_width=True):
        st.rerun()


def _rerun_app() -> None:
    """
    Rerun the whole page. Inside a fragment st.rerun() defaults to fragment
    scope, which would leave the rest of the page stale.
    """
    try:
        st.rerun(scope="app")
    except Exception:
        st.rerun()


@st.dialog("Confirmar entrega")
def confirm_extra_delivery_dialog(numero: str, nome: str) -> None:
    st.markdown(
        f"Confirmar entrega da cesta extra:<br>"
        f"<b>{nome}</b> &nbsp;·&nbsp; Nº {numero}",
        unsafe_allow_html=True,
    )
    st.write("")
    col1, col2 = st.columns(2)
    if col1.button("✅ Confirmar", use_container_width=True, type="primary"):
        # Shared memory first → every instance sees it on its next tick.
        # The Drive write is queued and happens in the background.
        extras_mark_delivered(numero)
        st.toast(f"✅ Cesta extra Nº {numero} entregue.")
        st.session_state.extra_saved = f"Cesta extra Nº {numero} marcada como entregue."
        _rerun_app()
    if col2.button("❌ Cancelar", use_container_width=True):
        _rerun_app()


@st.dialog("Dar Cesta Extra")
def cesta_extra_dialog(suggested_num: int) -> None:
    st.markdown("Registrar uma nova cesta extra.")

    num = st.text_input(
        "NÚMERO *", value=str(suggested_num), max_chars=10, key="ce_num"
    )
    nome = st.text_input("NOME *", key="ce_nome")
    cpf  = st.text_input("CPF", key="ce_cpf")
    tel  = st.text_input("TELEFONE", key="ce_tel")
    motivo = st.text_area("MOTIVO *", height=90, key="ce_motivo")

    st.caption("* campos obrigatórios")

    # ── Validation
    errors: list[str] = []
    if not num.strip():
        errors.append("NÚMERO é obrigatório.")
    elif not num.strip().isdigit():
        errors.append("NÚMERO deve conter apenas dígitos.")
    if not nome.strip():
        errors.append("NOME é obrigatório.")
    if not motivo.strip():
        errors.append("MOTIVO é obrigatório.")

    st.write("")
    col1, col2 = st.columns(2)
    save_clicked   = col1.button("💾 Salvar",  use_container_width=True, type="primary")
    cancel_clicked = col2.button("❌ Cancelar", use_container_width=True)

    if cancel_clicked:
        st.rerun()

    if save_clicked:
        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        record = {
            "NÚMERO":   num.strip(),
            "NOME":     nome.strip(),
            "CPF":      cpf.strip(),
            "TELEFONE": tel.strip(),
            "MOTIVO":   motivo.strip(),
        }

        # Shared memory first (instant, visible to every instance);
        # the Drive write is queued and runs in the background.
        err = extras_add(record)
        if err:
            st.error(err)
            st.stop()

        st.toast(f"✅ Cesta extra Nº {record['NÚMERO']} registrada.")
        st.session_state.extra_saved = f"Cesta extra Nº {record['NÚMERO']} registrada."
        _rerun_app()


@st.dialog("Acesso Administrativo")
def admin_login_dialog() -> None:
    if not ADMIN_PASSWORD:
        st.error(
            "Acesso administrativo indisponível: "
            "`ADMIN_PASSWORD` não está configurada nos Secrets."
        )
        if st.button("Fechar", use_container_width=True):
            st.rerun()
        return

    st.markdown("Informe a senha para habilitar a edição de dados.")
    pwd = st.text_input("Senha", type="password", key="admin_pwd_input")
    st.write("")
    col1, col2 = st.columns(2)
    if col1.button("🔓 Entrar", use_container_width=True, type="primary"):
        if pwd and pwd == ADMIN_PASSWORD:
            st.session_state.is_admin = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    if col2.button("❌ Cancelar", use_container_width=True):
        st.rerun()


@st.dialog("Editar Dados")
def edit_dados_dialog(row_idx: int, nome: str, numero: str,
                      current: dict[str, str]) -> None:
    """Edit all Tier 1 editable fields for one record."""
    st.markdown(f"Editando dados de **{nome}**")

    # ── Tipo: dropdown of values already present in the data
    tipo_opts = sorted({v for v in df["TIPO"].fillna("").astype(str).str.strip() if v}) \
        if "TIPO" in df.columns else []
    cur_tipo = current.get("TIPO", "")
    if cur_tipo and cur_tipo not in tipo_opts:
        tipo_opts.insert(0, cur_tipo)
    if "" not in tipo_opts:
        tipo_opts.insert(0, "")
    new_tipo = st.selectbox(
        "Tipo", tipo_opts,
        index=tipo_opts.index(cur_tipo) if cur_tipo in tipo_opts else 0,
    )

    # ── Status: fixed binary
    status_opts = ["", "Ativo", "Inativo"]
    cur_status = current.get("STATUS", "")
    if cur_status and cur_status not in status_opts:
        status_opts.append(cur_status)
    new_status = st.selectbox(
        "Status", status_opts,
        index=status_opts.index(cur_status) if cur_status in status_opts else 0,
    )

    new_cid = st.text_input("CID 2026", value=current.get("CID 2026", ""), max_chars=40)

    st.markdown("**Reservas**")
    new_res1  = st.text_input("Reserva 1",  value=current.get("RESERVA 1", ""))
    new_cpf1  = st.text_input("CPF Res. 1", value=current.get("CPF RESERVA 1", ""))
    new_res2  = st.text_input("Reserva 2",  value=current.get("RESERVA 2", ""))
    new_cpf2  = st.text_input("CPF Res. 2", value=current.get("CPF RESERVA 2", ""))

    new_alerta = st.text_area(
        "⚠️ Alerta para a Mesa",
        value=current.get("ALERTA PARA A MESA", ""),
        height=80,
    )

    proposed = {
        "TIPO":               new_tipo.strip(),
        "STATUS":             new_status.strip(),
        "CID 2026":           new_cid.strip(),
        "RESERVA 1":          new_res1.strip(),
        "CPF RESERVA 1":      new_cpf1.strip(),
        "RESERVA 2":          new_res2.strip(),
        "CPF RESERVA 2":      new_cpf2.strip(),
        "ALERTA PARA A MESA": new_alerta.strip(),
    }
    # Only persist fields that actually changed
    changed = {k: v for k, v in proposed.items() if v != current.get(k, "")}

    st.write("")
    if changed:
        st.caption(f"{len(changed)} campo(s) alterado(s): " + ", ".join(changed.keys()))
    else:
        st.caption("Nenhuma alteração.")

    col1, col2 = st.columns(2)
    if col1.button("💾 Salvar", use_container_width=True, type="primary",
                   disabled=not changed):
        # Persist first, then mirror into the overlay. Updating the overlay
        # before a failed save would show the new values as if they had been
        # stored, hiding the loss until someone reloaded the page.
        if _gdrive_write_enabled():
            with st.spinner("Salvando no Google Drive…"):
                err = save_cells_to_drive(numero, changed)
            if err:
                st.error(f"Erro ao salvar: {err}")
                st.stop()
        st.session_state.edits.setdefault(numero, {}).update(changed)
        st.rerun()
    if col2.button("❌ Cancelar", use_container_width=True):
        st.rerun()


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
    row = apply_edits(row)
    row_idx = row.name

    q     = query.strip().lower()
    q_num = numero_query.strip()

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
      <div class="info-row info-alerta-row">
        <span class="info-label">⚠️ Alerta</span>
        <span class="info-value info-alerta-val">{highlight(alerta, q)}</span>
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

    # ── Crianças (only when there are any)
    criancas = read_criancas(row)
    if criancas:
        html += '<div class="sec-title">🧒 Crianças</div>'
        html += '<div class="kids">'
        for c in criancas:
            k_nome = str(c.get("NOME", "")).strip() or "—"
            emoji  = SEXO_EMOJI.get(str(c.get("SEXO", "")).strip(), "🧒")
            try:
                k_idade = int(float(c.get("IDADE")))
                idade_txt = f'<span class="kid-age">{k_idade}a</span>'
            except (TypeError, ValueError):
                idade_txt = ""
            # Deliberately NOT highlighted: the children column is not a SEARCH
            # field, so lighting up a child's name would suggest the record
            # matched because of it.
            html += (f'<span class="kid">{emoji} '
                     f'{k_nome} {idade_txt}</span>')
        html += "</div>"

    html += "</div>"   # close card

    st.markdown(html, unsafe_allow_html=True)

    # ── Action buttons
    cur_month = current_month_col()
    month_ok  = cur_month and cell(row, cur_month).lower() == "ok"

    btn_cols = st.columns(3)
    with btn_cols[0]:
        is_admin = st.session_state.get("is_admin", False)
        if st.button(
            "✏️ Editar Dados" if is_admin else "🔒 Editar Dados",
            key=f"cid_{row_idx}",
            use_container_width=True,
            disabled=not is_admin,
            help=None if is_admin else "Requer acesso administrativo",
        ):
            edit_dados_dialog(row_idx, nome, numero, {
                "TIPO":               tipo,
                "STATUS":             status,
                "CID 2026":           cid,
                "RESERVA 1":          reserva1,
                "CPF RESERVA 1":      cpf_reserva1,
                "RESERVA 2":          reserva2,
                "CPF RESERVA 2":      cpf_reserva2,
                "ALERTA PARA A MESA": alerta,
            })
    with btn_cols[1]:
        if cur_month:
            if month_ok:
                st.button(
                    f"✅ {cur_month} OK",
                    key=f"flag_{row_idx}",
                    use_container_width=True,
                    disabled=True,
                )
            else:
                if st.button(
                    f"📦 Marcar {cur_month}",
                    key=f"flag_{row_idx}",
                    use_container_width=True,
                ):
                    confirm_month_dialog(row_idx, cur_month, nome, numero)
    with btn_cols[2]:
        if criancas:
            if st.button(
                f"👧 Editar Crianças ({len(criancas)})",
                key=f"cadastrar_crianca_{row_idx}",
                use_container_width=True,
            ):
                # Fresh widget keys on opening; retained during dialog reruns.
                st.session_state.criancas_form_version = (
                    st.session_state.get("criancas_form_version", 0) + 1
                )
                editar_criancas_dialog(row_idx, nome, numero, criancas)
        else:
            if st.button(
                "👶 Cadastrar Criança",
                key=f"cadastrar_crianca_{row_idx}",
                use_container_width=True,
            ):
                st.session_state.crianca_form_version = (
                    st.session_state.get("crianca_form_version", 0) + 1
                )
                cadastrar_crianca_dialog(row_idx, nome, numero)


# ──────────────────────────────────────────────────────────────
#  SCREEN — LISTA DE CESTAS EXTRAS
# ──────────────────────────────────────────────────────────────
def _render_extras_list() -> None:
    """
    Body of the Cestas Extras screen.
    Reads the shared in-memory store, so it never waits on the network.
    """
    extras = extras_snapshot()
    status = extras_status()

    # Rendered inside the fragment so it survives a fragment-scoped rerun
    if st.session_state.get("extra_saved"):
        st.success(f"✅ {st.session_state.pop('extra_saved')}")

    # ── Sync status bar
    bits = []
    if not status["drive"]:
        bits.append("⚠️ Drive não configurado — alterações só em memória")
    elif status["pending"]:
        bits.append(f"⏳ gravando {status['pending']} no Drive…")
    elif status["synced"]:
        bits.append(
            "✓ sincronizado com o Drive às "
            + datetime.fromtimestamp(status["synced"]).strftime("%H:%M:%S")
        )
    else:
        bits.append("✓ em dia")
    bits.append(f"tela atualizada às {datetime.now():%H:%M:%S}")

    scol1, scol2 = st.columns([3, 1])
    scol1.caption(" · ".join(bits))
    if scol2.button("🔄 Drive", key="extras_resync", use_container_width=True,
                    disabled=bool(status["pending"]),
                    help="Recarregar do Google Drive"):
        refused = extras_init(force=True)
        if refused:
            st.warning(refused)
        else:
            _rerun_app()

    if status["failed"]:
        st.error(
            f"⚠️ {status['failed']} alteração(ões) não foram gravadas no Drive. "
            "Os dados continuam visíveis aqui."
        )
        if st.button("↻ Tentar gravar novamente", key="extras_retry",
                     use_container_width=True):
            n = extras_retry_failed()
            st.toast(f"Reenviando {n} alteração(ões)…")
            _rerun_app()

    # Upload worked but a record was not found in the sheet — retrying would
    # not help, so this is reported on its own rather than queued.
    if status["stale"]:
        st.warning(
            f"⚠️ {status['stale']} alteração(ões) não puderam ser aplicadas "
            "na planilha (registro não encontrado ou duplicado na aba). "
            "A tela foi desfeita para refletir a planilha — veja o motivo "
            "abaixo, corrija a aba e repita a ação."
        )

    for e in status["errors"]:
        st.caption(f"↳ {e}")

    if not extras:
        st.markdown(
            """
<div class="empty-state">
  <div class="icon">🧺</div>
  <p>Nenhuma cesta extra registrada ainda.</p>
</div>
""",
            unsafe_allow_html=True,
        )
        return

    total     = len(extras)
    entregues = sum(
        1 for r in extras
        if str(r.get(EXTRAS_DONE_HEADER, "")).strip().lower() == "ok"
    )
    st.markdown(
        f'<div class="result-count">{total} cesta(s) extra(s) · '
        f'{entregues} entregue(s) · {total - entregues} pendente(s)</div>',
        unsafe_allow_html=True,
    )

    for i, row in enumerate(extras):
        numero = str(row.get("NÚMERO", "")).strip()
        nome   = str(row.get("NOME", "")).strip()
        cpf    = str(row.get("CPF", "")).strip()
        tel    = str(row.get("TELEFONE", "")).strip()
        motivo = str(row.get("MOTIVO", "")).strip()
        done   = str(row.get(EXTRAS_DONE_HEADER, "")).strip().lower() == "ok"

        badge = (
            '<span class="badge b-ativo">ENTREGUE</span>' if done
            else '<span class="badge b-inativo">PENDENTE</span>'
        )

        html = f"""
        <div class="card">
          <div class="card-name">{nome or "—"}</div>
          <div class="card-num">Nº {numero or "—"}
            &nbsp;·&nbsp; CPF {cpf or "—"}
            &nbsp;·&nbsp; Tel {tel or "—"}</div>
          <div class="info-grid">
            <div class="info-row">
              <span class="info-label">Situação</span>
              <span class="info-value">{badge}</span>
            </div>
            <div class="info-row">
              <span class="info-label">Motivo</span>
              <span class="info-value">{motivo or '<span class="info-empty">—</span>'}</span>
            </div>
          </div>
        </div>
        """
        st.markdown(html, unsafe_allow_html=True)

        if done:
            st.button(
                "✅ Entregue",
                key=f"extra_done_{i}_{numero}",
                use_container_width=True,
                disabled=True,
            )
        else:
            if st.button(
                "📦 Marcar como entregue",
                key=f"extra_done_{i}_{numero}",
                use_container_width=True,
            ):
                confirm_extra_delivery_dialog(numero, nome)


# Auto-refresh where supported. The tick is short because each one is a
# pure in-memory read — no Drive roundtrip — so a write by one admin shows
# up on every other open instance within a couple of seconds.
try:
    render_extras_list = st.fragment(run_every="2s")(_render_extras_list)
except Exception:
    render_extras_list = _render_extras_list


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

# Show Drive write status — only when secrets are missing
if not _gdrive_write_enabled():
    st.warning(
        "⚠️ Salvamento automático desativado. "
        "Configure `GDRIVE_FILE_ID` e `GDRIVE_SA_CREDENTIALS` nos Secrets para habilitar.",
        icon="🔒",
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
#  SCREEN ROUTING — Cestas Extras (admin only)
# ──────────────────────────────────────────────────────────────
if "view" not in st.session_state:
    st.session_state.view = "main"

# Losing admin must never leave the user stranded on a restricted screen
if st.session_state.view == "extras" and not st.session_state.get("is_admin", False):
    st.session_state.view = "main"

if st.session_state.view == "extras":
    back, _sp = st.columns([1, 2])
    if back.button("← Voltar", key="extras_back", use_container_width=True):
        st.session_state.view = "main"
        st.rerun()

    st.markdown('<div class="sec-title">🧺 Lista de Cestas Extras</div>',
                unsafe_allow_html=True)
    extras_init()          # hydrates from Drive once per server process
    render_extras_list()
    st.stop()

# ──────────────────────────────────────────────────────────────
#  UI — SEARCH
# ──────────────────────────────────────────────────────────────
if st.session_state.get("crianca_saved"):
    st.success(f"✅ {st.session_state.pop('crianca_saved')}")

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

# ──────────────────────────────────────────────────────────────
#  UI — ADMIN FOOTER (discreet)
# ──────────────────────────────────────────────────────────────
st.markdown('<div class="admin-footer">', unsafe_allow_html=True)

if st.session_state.get("is_admin", False):
    # ── Success toast from a previous save
    if st.session_state.get("extra_saved"):
        st.success(f"✅ {st.session_state.pop('extra_saved')}")

    ecol1, ecol2 = st.columns(2)
    if ecol1.button("🧺 Dar Cesta Extra", key="btn_cesta_extra",
                    use_container_width=True, type="primary"):
        extras_init()
        cesta_extra_dialog(extras_next_numero())
    if ecol2.button("📋 Lista de Cestas Extras", key="btn_lista_extras",
                    use_container_width=True):
        st.session_state.view = "extras"
        st.rerun()

    fcol1, fcol2 = st.columns([3, 2])
    fcol1.markdown(
        '<div class="admin-badge">🔓 Modo administrador ativo</div>',
        unsafe_allow_html=True,
    )
    if fcol2.button("Sair", key="admin_logout", use_container_width=True):
        st.session_state.is_admin = False
        st.rerun()
else:
    if st.button("· acesso administrativo ·", key="admin_login_link"):
        admin_login_dialog()

st.markdown("</div>", unsafe_allow_html=True)
