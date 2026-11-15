"""
Pastoral Social — Apadrinhamento de Crianças
Mobile-first Streamlit app que lê e grava a planilha crianças.xlsx no Google Drive.

Setup (igual ao app de cadastro):
  1. Secrets:
       GITHUB_RAW_URL         → link de download da planilha crianças.xlsx
       GDRIVE_FILE_ID         → ID do arquivo crianças.xlsx no Google Drive
       GDRIVE_SA_CREDENTIALS  → JSON da service account (com acesso de edição)
  2. Deploy no Streamlit Community Cloud.

Estrutura esperada da planilha (1ª aba, cabeçalho na linha 1):
  Nº Família / Carteirinha | Beneficiário (Cesta) | Nome da criança | Sexo | Idade | Observações
  As colunas "Padrinho" e "Telefone Padrinho" são criadas automaticamente
  no primeiro salvamento, caso ainda não existam.
"""

import html as _html
import json
import re
import unicodedata
from io import BytesIO

import pandas as pd
import requests
import streamlit as st

# ──────────────────────────────────────────────────────────────
#  CONFIG
# ──────────────────────────────────────────────────────────────
try:
    GITHUB_RAW_URL = st.secrets["GITHUB_RAW_URL"]
except Exception:
    GITHUB_RAW_URL = (
        "https://github.com/kauemelo01/Cadastro-PS-PSB/raw/refs/heads/main/"
        "crian%C3%A7as.xlsx"
    )

# Nomes das colunas (a busca no cabeçalho ignora acentos e maiúsculas)
COL_NUM      = "Nº Família / Carteirinha"
COL_BENEF    = "Beneficiário (Cesta)"
COL_NOME     = "Nome da criança"
COL_SEXO     = "Sexo"
COL_IDADE    = "Idade"
COL_OBS      = "Observações"
COL_PADRINHO = "Padrinho"
COL_TEL      = "Telefone Padrinho"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ──────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Apadrinhamento",
    page_icon="🎁",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────────────────────
#  MOBILE-FIRST CSS
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
  .main .block-container { padding: 0.75rem 0.75rem 2rem; max-width: 500px; margin: auto; }

  .ps-header {
    background: linear-gradient(135deg, #1b5e20 0%, #388e3c 100%);
    border-radius: 14px; padding: 1rem 1.25rem 0.75rem; color: #fff; margin-bottom: 1rem;
  }
  .ps-header h1 { margin: 0; font-size: 1.35rem; color: #fff; }
  .ps-header p  { margin: 0.15rem 0 0; font-size: 0.82rem; opacity: 0.85; }

  .stTextInput > div > div > input {
    font-size: 16px !important; border-radius: 12px !important;
    border: 2px solid #c8e6c9 !important; padding: 0.65rem 1rem !important;
    background: #f9fbe7 !important; color: #212121 !important;
  }
  .stTextInput > div > div > input:focus {
    border-color: #388e3c !important; box-shadow: 0 0 0 3px rgba(56,142,60,.15) !important;
  }

  /* ── Família (cabeçalho do resultado) ── */
  .fam {
    font-size: 0.82rem; color: #616161; margin: 0.2rem 0 0.7rem 2px;
  }
  .fam b { color: #1b5e20; }

  /* ── Card da criança ── */
  .kid {
    background: #fff; border-radius: 14px; padding: 0.9rem 1rem;
    margin: 0.9rem 0 0.4rem; box-shadow: 0 2px 10px rgba(0,0,0,.08);
    border-left: 5px solid #9e9e9e;
  }
  .kid.boy  { border-left-color: #1e88e5; }
  .kid.girl { border-left-color: #d81b60; }
  .kid-top { display: flex; align-items: center; gap: 12px; }
  .avatar {
    position: relative; flex-shrink: 0; width: 54px; height: 54px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center; font-size: 1.9rem;
    background: #eeeeee;
  }
  .boy  .avatar { background: #e3f2fd; }
  .girl .avatar { background: #fce4ec; }
  .avatar .tag {
    position: absolute; right: -4px; bottom: -4px; font-size: 1rem;
    background: #fff; border-radius: 50%; width: 22px; height: 22px;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 1px 3px rgba(0,0,0,.2);
  }
  .kid-name { font-size: 1.2rem; font-weight: 800; color: #212121; line-height: 1.2; }
  .kid-sub  { font-size: 0.82rem; color: #757575; margin-top: 2px; }
  .kid-benef { font-size: 0.75rem; color: #9e9e9e; margin-top: 3px; }
  .kid-obs {
    background: #fff8e1; border-radius: 8px; padding: 6px 9px;
    font-size: 0.82rem; color: #5d4037; margin-top: 9px;
  }
  .pad-ok {
    background: #e8f5e9; border-radius: 8px; padding: 7px 10px;
    font-size: 0.88rem; color: #1b5e20; margin-top: 9px;
  }
  .pad-ok .tel { color: #2e7d32; font-size: 0.82rem; }
  .pad-no {
    background: #f5f5f5; border-radius: 8px; padding: 7px 10px;
    font-size: 0.85rem; color: #9e9e9e; font-style: italic; margin-top: 9px;
  }

  /* ── Estatísticas ── */
  .stats { margin-top: 2.2rem; }
  .stats h3 { font-size: 0.85rem; text-transform: uppercase; letter-spacing: .04em;
              color: #616161; margin: 0 0 .6rem; }
  .kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 6px; }
  .kpi { background: #fff; border-radius: 12px; padding: 8px 4px; text-align: center;
         box-shadow: 0 1px 6px rgba(0,0,0,.07); }
  .kpi .v { font-size: 1.25rem; font-weight: 800; color: #1b5e20; }
  .kpi .l { font-size: 0.68rem; color: #757575; text-transform: uppercase; }
  .kpi.warn .v { color: #e65100; }
  .grp { background: #fff; border-radius: 12px; padding: 9px 11px; margin-top: 8px;
         box-shadow: 0 1px 6px rgba(0,0,0,.07); }
  .grp-head { display: flex; justify-content: space-between; font-size: 0.85rem; color: #424242; }
  .grp-head b { color: #212121; }
  .bar { height: 9px; background: #eeeeee; border-radius: 6px; overflow: hidden; margin: 6px 0 4px; }
  .bar > div { height: 100%; background: linear-gradient(90deg, #43a047, #66bb6a); }
  .grp-foot { font-size: 0.74rem; color: #757575; }

  .empty-state { text-align: center; padding: 2.5rem 1rem; color: #9e9e9e; }
  .empty-state .icon { font-size: 2.8rem; }
  .empty-state p { margin: 0.5rem 0 0; font-size: 0.95rem; }

  @media (prefers-color-scheme: dark) {
    .stTextInput > div > div > input {
      background: #1a2e1a !important; color: #e8f5e9 !important; border-color: #4caf50 !important;
    }
    .kid, .kpi, .grp { background: #1e2b1e !important; }
    .kid-name, .grp-head b { color: #e8f5e9 !important; }
    .kid-sub, .grp-head, .fam { color: #bdbdbd !important; }
    .fam b, .kpi .v { color: #a5d6a7 !important; }
    .avatar .tag { background: #1e2b1e !important; }
    .boy .avatar  { background: #1a2a3a !important; }
    .girl .avatar { background: #3a1a2a !important; }
    .pad-ok { background: #1f3a1f !important; color: #c8e6c9 !important; }
    .pad-ok .tel { color: #a5d6a7 !important; }
    .pad-no { background: #2a2a2a !important; }
    .kid-obs { background: #3e2e10 !important; color: #ffe0b2 !important; }
    .bar { background: #333 !important; }
  }

  #MainMenu { visibility: hidden; } footer { visibility: hidden; } header { visibility: hidden; }
</style>
""",
    unsafe_allow_html=True,
)


# ──────────────────────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────────────────────
def esc(s: str) -> str:
    return _html.escape(s or "")


def _norm(s: object) -> str:
    """Lowercase, sem acentos, sem espaços extras (inclui espaço não-separável)."""
    s = "" if s is None else str(s)
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return " ".join(s.replace("\xa0", " ").split()).lower()


def clean(v: object) -> str:
    """Valor da célula como texto limpo ('' para vazio/NaN)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).replace("\xa0", " ").strip()
    return "" if s.lower() in ("nan", "none") else s


def norm_numero(v: object) -> str:
    """'015', '15', '15.0' → '15'."""
    s = clean(v)
    if re.fullmatch(r"\d+(\.0+)?", s):
        s = s.split(".")[0].lstrip("0") or "0"
    return s


def parse_idade(v: str) -> int | None:
    """'6 anos' → 6 · '3 meses' → 0 · '7' → 7 · vazio → None."""
    s = _norm(v)
    m = re.search(r"(\d+)", s)
    if not m:
        return None
    n = int(m.group(1))
    if "mes" in s:
        return n // 12
    return n


def faixa(idade: int | None) -> str | None:
    if idade is None:
        return None
    if idade <= 3:
        return "0 a 3 anos"
    if idade <= 8:
        return "4 a 8 anos"
    return "9 anos ou mais"


FAIXAS = ["0 a 3 anos", "4 a 8 anos", "9 anos ou mais"]


def genero(sexo: str) -> str:
    s = _norm(sexo)
    if s.startswith("f"):
        return "girl"
    if s.startswith("m"):
        return "boy"
    return ""


def avatar(sexo: str, idade: int | None) -> tuple[str, str]:
    """(emoji principal, emoji de canto) conforme sexo e faixa etária."""
    g = genero(sexo)
    f = faixa(idade)
    if f == "0 a 3 anos":
        return "👶", {"girl": "🎀", "boy": "🧢"}.get(g, "")
    if f == "4 a 8 anos":
        return {"girl": "👧", "boy": "👦"}.get(g, "🧒"), ""
    if f == "9 anos ou mais":
        return {"girl": "🙋‍♀️", "boy": "🙋‍♂️"}.get(g, "🧑"), ""
    return {"girl": "👧", "boy": "👦"}.get(g, "🧒"), ""


def find_col(headers: list[str], wanted: str) -> int | None:
    """Índice (0-based) do cabeçalho que corresponde a `wanted`, ignorando acentos."""
    w = _norm(wanted)
    for i, h in enumerate(headers):
        if _norm(h) == w:
            return i
    return None


# ──────────────────────────────────────────────────────────────
#  DATA LOADING
# ──────────────────────────────────────────────────────────────
def _fetch_gdrive(url: str) -> bytes:
    """Baixa um arquivo de um link 'uc?export=download' do Google Drive,
    tratando a página de confirmação que o Drive mostra para arquivos grandes."""
    session = requests.Session()
    resp = session.get(url, timeout=15)
    resp.raise_for_status()
    if "text/html" in resp.headers.get("Content-Type", ""):
        match = re.search(r"confirm=([0-9A-Za-z_\-]+)", resp.text)
        token = match.group(1) if match else "t"
        resp = session.get(url + f"&confirm={token}", timeout=30)
        resp.raise_for_status()
    return resp.content


@st.cache_data(ttl=300, show_spinner=False)
def load_data(url: str):
    """Retorna (df, erro). df tem colunas padronizadas + '_row' (linha no xlsx, 1-based)."""
    try:
        content = _fetch_gdrive(url)
    except Exception as exc:
        return None, str(exc)

    try:
        raw = pd.read_excel(BytesIO(content), sheet_name=0, header=None, dtype=str)
    except Exception as exc:
        return None, str(exc)

    if raw.shape[0] < 2:
        return None, "Arquivo sem dados suficientes."

    headers = [clean(h) for h in raw.iloc[0].tolist()]
    idx = {c: find_col(headers, c) for c in
           (COL_NUM, COL_BENEF, COL_NOME, COL_SEXO, COL_IDADE, COL_OBS, COL_PADRINHO, COL_TEL)}
    missing = [c for c in (COL_NUM, COL_NOME) if idx[c] is None]
    if missing:
        return None, "Coluna(s) não encontrada(s) na planilha: " + ", ".join(missing)

    body = raw.iloc[1:]
    out = pd.DataFrame({
        c: (body.iloc[:, i].map(clean) if i is not None else "")
        for c, i in idx.items()
    })
    out["_row"] = body.index + 1                       # linha 1 = cabeçalho
    out[COL_NUM] = out[COL_NUM].map(norm_numero)
    # Ignora linhas soltas sem número ou sem nome da criança
    out = out[(out[COL_NUM] != "") & (out[COL_NOME] != "")].reset_index(drop=True)
    out["_idade"] = out[COL_IDADE].map(parse_idade)
    return out, None


# ──────────────────────────────────────────────────────────────
#  GOOGLE DRIVE WRITE-BACK
# ──────────────────────────────────────────────────────────────
def _gdrive_write_enabled() -> bool:
    try:
        return "GDRIVE_FILE_ID" in st.secrets and "GDRIVE_SA_CREDENTIALS" in st.secrets
    except Exception:
        return False


def _get_drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    creds_info = json.loads(st.secrets["GDRIVE_SA_CREDENTIALS"])
    creds = Credentials.from_service_account_info(
        creds_info, scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _candidate_rows(ws, col: int) -> list[int]:
    """Linhas (>=2) com valor na coluna `col`. Lê só as células que existem,
    para não varrer ~1 milhão de linhas vazias quando a planilha tem
    formatação/sobras até o fim da folha."""
    try:
        return sorted(r for (r, c), cell in ws._cells.items()
                      if c == col and r >= 2 and cell.value not in (None, ""))
    except Exception:
        return [r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=col).value not in (None, "")]


def patch_padrinho(wb, row_hint: int, numero: str, crianca: str,
                   padrinho: str, telefone: str) -> str | None:
    """Grava Padrinho/Telefone na linha da criança dentro do workbook aberto.
    A linha é localizada por Nº + nome da criança (não só pela posição),
    para que linhas reordenadas não levem a gravar na criança errada."""
    ws = wb.worksheets[0]
    last_col = max((c.column for c in ws[1] if clean(c.value)), default=0)
    headers = [clean(ws.cell(row=1, column=c).value) for c in range(1, last_col + 1)]

    i_num, i_nome = find_col(headers, COL_NUM), find_col(headers, COL_NOME)
    if i_num is None or i_nome is None:
        return "Cabeçalho da planilha não reconhecido (Nº Família / Nome da criança)."

    # Cria as colunas de padrinho se ainda não existirem
    i_pad, i_tel = find_col(headers, COL_PADRINHO), find_col(headers, COL_TEL)
    if i_pad is None:
        last_col += 1
        ws.cell(row=1, column=last_col, value=COL_PADRINHO)
        i_pad = last_col - 1
    if i_tel is None:
        last_col += 1
        ws.cell(row=1, column=last_col, value=COL_TEL)
        i_tel = last_col - 1

    def matches(r: int) -> bool:
        return (norm_numero(ws.cell(row=r, column=i_num + 1).value) == numero
                and _norm(ws.cell(row=r, column=i_nome + 1).value) == _norm(crianca))

    if row_hint >= 2 and matches(row_hint):
        target = row_hint
    else:
        found = [r for r in _candidate_rows(ws, i_num + 1) if matches(r)]
        if not found:
            return f"{crianca} (Nº {numero}) não encontrada no arquivo atual."
        if len(found) > 1:
            return (f"Há mais de uma criança chamada {crianca} no Nº {numero}. "
                    "Nenhum dado foi alterado.")
        target = found[0]

    ws.cell(row=target, column=i_pad + 1, value=padrinho)
    ws.cell(row=target, column=i_tel + 1, value=telefone)
    return None


def save_padrinho_to_drive(row_hint: int, numero: str, crianca: str,
                           padrinho: str, telefone: str) -> str | None:
    """Baixa a versão atual do Drive, grava e reenvia. None = sucesso."""
    if not _gdrive_write_enabled():
        return "Credenciais do Google Drive não configuradas."
    try:
        from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
        import openpyxl

        file_id = st.secrets["GDRIVE_FILE_ID"]
        service = _get_drive_service()

        req = service.files().get_media(fileId=file_id)
        dl_buf = BytesIO()
        downloader = MediaIoBaseDownload(dl_buf, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        dl_buf.seek(0)

        wb = openpyxl.load_workbook(dl_buf)
        try:
            err = patch_padrinho(wb, row_hint, numero, crianca, padrinho, telefone)
            if err:
                return err
            up_buf = BytesIO()
            wb.save(up_buf)
        finally:
            wb.close()

        up_buf.seek(0)
        media = MediaIoBaseUpload(up_buf, mimetype=XLSX_MIME, resumable=False)
        service.files().update(fileId=file_id, media_body=media).execute()
        load_data.clear()
        return None
    except Exception as exc:
        return str(exc)


# ──────────────────────────────────────────────────────────────
#  SESSION STATE — edições da sessão (aparecem na hora, mesmo
#  antes do cache recarregar o arquivo)
# ──────────────────────────────────────────────────────────────
if "pad_edits" not in st.session_state:
    st.session_state.pad_edits = {}   # {(numero, nome_norm): (padrinho, telefone)}


def apply_edits(df: pd.DataFrame) -> pd.DataFrame:
    edits = st.session_state.pad_edits
    if not edits:
        return df
    df = df.copy()
    for i, r in df.iterrows():
        key = (r[COL_NUM], _norm(r[COL_NOME]))
        if key in edits:
            df.at[i, COL_PADRINHO], df.at[i, COL_TEL] = edits[key]
    return df


# ──────────────────────────────────────────────────────────────
#  DIALOG — Cadastrar / Editar Padrinho
# ──────────────────────────────────────────────────────────────
def _padrinho_form(kid: dict) -> None:
    version = st.session_state.get("pad_form_version", 0)
    prefix = f"pad_{kid['_row']}_{version}"

    st.markdown(f"Criança: **{esc(kid[COL_NOME])}**<br>"
                f"<span style='font-size:.85rem;color:#757575'>Nº {esc(kid[COL_NUM])} · "
                f"{esc(kid[COL_BENEF])}</span>", unsafe_allow_html=True)

    padrinho = st.text_input("Padrinho *", value=kid[COL_PADRINHO], key=f"{prefix}_nome")
    telefone = st.text_input("Telefone *", value=kid[COL_TEL], key=f"{prefix}_tel",
                             placeholder="(11) 91234-5678")
    st.caption("\\* campos obrigatórios")

    st.write("")
    c1, c2 = st.columns(2)
    save_clicked = c1.button("💾 Salvar", use_container_width=True, type="primary",
                             key=f"{prefix}_salvar")
    cancel_clicked = c2.button("❌ Cancelar", use_container_width=True,
                               key=f"{prefix}_cancelar")
    if cancel_clicked:
        st.rerun()

    if save_clicked:
        padrinho, telefone = " ".join(padrinho.split()), telefone.strip()
        errors = []
        if not padrinho:
            errors.append("Padrinho é obrigatório.")
        if not telefone:
            errors.append("Telefone é obrigatório.")
        elif len(re.sub(r"\D", "", telefone)) < 10:
            errors.append("Telefone deve ter DDD + número (pelo menos 10 dígitos).")
        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        if (padrinho, telefone) == (kid[COL_PADRINHO], kid[COL_TEL]):
            st.rerun()   # nada mudou

        with st.spinner("Salvando no Google Drive…"):
            err = save_padrinho_to_drive(kid["_row"], kid[COL_NUM], kid[COL_NOME],
                                         padrinho, telefone)
        if err:
            st.error(f"Erro ao salvar: {err}")
            st.stop()

        st.session_state.pad_edits[(kid[COL_NUM], _norm(kid[COL_NOME]))] = (padrinho, telefone)
        st.session_state.pad_saved = f"Padrinho de {kid[COL_NOME]} salvo: {padrinho}."
        st.rerun()


@st.dialog("Cadastrar Padrinho")
def cadastrar_padrinho_dialog(kid: dict) -> None:
    _padrinho_form(kid)


@st.dialog("Editar Padrinho")
def editar_padrinho_dialog(kid: dict) -> None:
    _padrinho_form(kid)


# ──────────────────────────────────────────────────────────────
#  RENDER
# ──────────────────────────────────────────────────────────────
def render_kid(kid: dict) -> None:
    g = genero(kid[COL_SEXO])
    emoji, tag = avatar(kid[COL_SEXO], kid["_idade"])
    sexo_lbl = {"girl": "Menina", "boy": "Menino"}.get(g, kid[COL_SEXO] or "Sexo não informado")
    idade_lbl = kid[COL_IDADE] or "idade não informada"
    tag_html = f'<span class="tag">{tag}</span>' if tag else ""

    obs = kid[COL_OBS]
    obs_html = f'<div class="kid-obs">📝 {esc(obs)}</div>' if obs else ""

    if kid[COL_PADRINHO]:
        tel = f' <span class="tel">· 📞 {esc(kid[COL_TEL])}</span>' if kid[COL_TEL] else ""
        pad_html = f'<div class="pad-ok">🤝 Padrinho: <b>{esc(kid[COL_PADRINHO])}</b>{tel}</div>'
    else:
        pad_html = '<div class="pad-no">Ainda sem padrinho</div>'

    st.markdown(
        f"""
<div class="kid {g}">
  <div class="kid-top">
    <div class="avatar">{emoji}{tag_html}</div>
    <div>
      <div class="kid-name">{esc(kid[COL_NOME])}</div>
      <div class="kid-sub">{sexo_lbl} · {esc(idade_lbl)}</div>
      <div class="kid-benef">Beneficiário: {esc(kid[COL_BENEF]) or "—"}</div>
    </div>
  </div>
  {obs_html}
  {pad_html}
</div>
""",
        unsafe_allow_html=True,
    )

    btn_key = f"pad_btn_{kid['_row']}"
    if kid[COL_PADRINHO]:
        clicked = st.button("✏️ Editar Padrinho", key=btn_key, use_container_width=True)
        dialog = editar_padrinho_dialog
    else:
        clicked = st.button("➕ Cadastrar Padrinho", key=btn_key, use_container_width=True,
                            type="primary")
        dialog = cadastrar_padrinho_dialog
    if clicked:
        # Campos novos a cada abertura; mantidos durante os reruns do popup.
        st.session_state.pad_form_version = st.session_state.get("pad_form_version", 0) + 1
        dialog(kid)


def render_stats(df: pd.DataFrame) -> None:
    def numbers(part: pd.DataFrame) -> tuple[int, int, int, float]:
        total = len(part)
        feitos = int((part[COL_PADRINHO] != "").sum())
        pct = (feitos / total * 100) if total else 0.0
        return total, feitos, total - feitos, pct

    total, feitos, faltam, pct = numbers(df)
    html = f"""
<div class="stats">
  <h3>📊 Apadrinhamento</h3>
  <div class="kpis">
    <div class="kpi"><div class="v">{feitos}</div><div class="l">Apadrinhadas</div></div>
    <div class="kpi"><div class="v">{pct:.0f}%</div><div class="l">Concluído</div></div>
    <div class="kpi"><div class="v">{total}</div><div class="l">Total</div></div>
    <div class="kpi warn"><div class="v">{faltam}</div><div class="l">Faltam</div></div>
  </div>
"""
    faixas = df["_idade"].map(faixa)
    for f, icon in zip(FAIXAS, ["👶", "🧒", "🙋"]):
        t, d, r, p = numbers(df[faixas == f])
        html += f"""
  <div class="grp">
    <div class="grp-head"><span>{icon} <b>{f}</b></span><span>{d} de {t} · {p:.0f}%</span></div>
    <div class="bar"><div style="width:{p:.1f}%"></div></div>
    <div class="grp-foot">Faltam {r}</div>
  </div>
"""
    sem_idade = int(faixas.isna().sum())
    if sem_idade:
        html += (f'<div class="grp-foot" style="margin-top:6px">'
                 f'{sem_idade} criança(s) sem idade informada (contam só no total).</div>')
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
#  UI
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
<div class="ps-header">
  <h1>🎁 Apadrinhamento</h1>
  <p>Pastoral Social · Jardim Colombo</p>
</div>
""",
    unsafe_allow_html=True,
)

if not _gdrive_write_enabled():
    st.warning(
        "⚠️ Salvamento automático desativado. "
        "Configure `GDRIVE_FILE_ID` e `GDRIVE_SA_CREDENTIALS` nos Secrets para habilitar.",
        icon="🔒",
    )

with st.spinner("Carregando crianças…"):
    df, load_error = load_data(GITHUB_RAW_URL)

if load_error or df is None:
    st.error("❌ Não foi possível carregar o arquivo crianças.xlsx.")
    st.markdown(
        """
**Possíveis causas:**
- O link do arquivo (`GITHUB_RAW_URL`) ainda não aponta para crianças.xlsx.
- O arquivo não existe ou não está compartilhado.
- Problema de rede.
"""
    )
    if load_error:
        with st.expander("Detalhes do erro"):
            st.code(load_error)
    st.stop()

df = apply_edits(df)

if st.session_state.get("pad_saved"):
    st.success(f"✅ {st.session_state.pop('pad_saved')}")

query = st.text_input(
    "Número do registro / carteirinha",
    placeholder="🔢  Nº do registro / carteirinha…",
    label_visibility="collapsed",
    key="q_num",
)
q = norm_numero(query)

if q:
    kids = df[df[COL_NUM] == q]
    if kids.empty:
        st.markdown(
            f"""
<div class="empty-state">
  <div class="icon">🔎</div>
  <p>Nenhuma criança vinculada ao Nº <b>{esc(q)}</b>.</p>
</div>
""",
            unsafe_allow_html=True,
        )
    else:
        n = len(kids)
        n_ok = int((kids[COL_PADRINHO] != "").sum())
        st.markdown(
            f'<div class="fam">Nº <b>{esc(q)}</b> · {n} criança{"s" if n > 1 else ""} · '
            f'{n_ok} com padrinho</div>',
            unsafe_allow_html=True,
        )
        for _, row in kids.iterrows():
            render_kid(row.to_dict())
else:
    st.markdown(
        """
<div class="empty-state">
  <div class="icon">🔍</div>
  <p>Digite o número do registro / carteirinha<br>para ver as crianças vinculadas.</p>
</div>
""",
        unsafe_allow_html=True,
    )

render_stats(df)
