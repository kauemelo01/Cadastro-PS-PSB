"""
Pastoral Social — Cadastro de Visitas
Aplicativo Streamlit com consulta, edição offline no aparelho e sincronização em lote.

Planilha: Visita_2026.xlsx (aba "CADASTRO GERAL")
  • Linha 1: numeração das colunas      • Linha 2: cabeçalhos      • Dados a partir da linha 3
  • Colunas 51 e 52 (criadas pelo app): ÚLTIMA EDIÇÃO POR / ÚLTIMA EDIÇÃO EM

Secrets (Streamlit Cloud → Settings → Secrets):
  GITHUB_RAW_URL         = "<url de leitura do xlsx>"       (opcional, há um padrão abaixo)
  GDRIVE_FILE_ID         = "<id do arquivo .xlsx no Drive>"  (habilita salvar no Drive)
  GDRIVE_SA_CREDENTIALS  = "<json da service account>"       (habilita salvar no Drive)
"""

import json
import re as _re
import hashlib
from pathlib import Path

import streamlit.components.v1 as components

from offline_sync import SYNC_LOCK, apply_batch, validate_operations
from datetime import date, datetime, time, timedelta, timezone
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
        "https://github.com/kauemelo01/Cadastro-PS-PSB/raw/refs/heads/main/cadastro.xlsx"
    )

try:
    from zoneinfo import ZoneInfo

    SP_TZ = ZoneInfo("America/Sao_Paulo")
except Exception:  # sem base de fusos no servidor → UTC-3 fixo
    SP_TZ = timezone(timedelta(hours=-3))

# Colunas de auditoria (posições 1-based na planilha: 51 e 52)
AUDIT_NAME_COL = "ÚLTIMA EDIÇÃO POR"
AUDIT_TS_COL = "ÚLTIMA EDIÇÃO EM"
AUDIT_NAME_POS = 51
AUDIT_TS_POS = 52
AUDIT_TS_FORMAT = "dd/mm/yyyy hh:mm:ss"

# ──────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Cadastro de Visitas",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────────────────────
#  LAYOUT DA PÁGINA (a interface offline tem seus próprios estilos)
# ──────────────────────────────────────────────────────────────
st.markdown(
    """<style>
    .main .block-container, .stMainBlockContainer { padding: 1rem .75rem 2rem; max-width: 540px; }
    .ps-header { background: linear-gradient(135deg,#1b5e20,#388e3c); border-radius: 14px;
                 padding: 1rem 1.25rem; color: white; margin-bottom: 1rem; }
    .ps-header h1 { font-size: 1.35rem; color: white; margin: 0; padding: 0; }
    .ps-header p { font-size: .82rem; margin: .25rem 0 0; opacity: .85; }
    </style>""", unsafe_allow_html=True,
)

# ──────────────────────────────────────────────────────────────
#  NORMALIZAÇÃO DE VALORES
# ──────────────────────────────────────────────────────────────
_ISO_RE = _re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}):(\d{2}))?(?:\.\d+)?$")


def norm_header(v) -> str:
    return _re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def norm_cell(v) -> str:
    """Valor de célula (pandas ou openpyxl) → texto de exibição comparável."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:  # NaN
            return ""
        if v.is_integer():
            v = int(v)
    if isinstance(v, datetime):
        v = v.strftime("%Y-%m-%d %H:%M:%S")
    elif isinstance(v, date):
        v = v.strftime("%Y-%m-%d")
    s = str(v).replace("\r\n", "\n").replace("\r", "\n").strip()
    if s.lower() in ("nan", "none", "nat"):
        return ""
    m = _ISO_RE.match(s)
    if m:
        y, mo, d, hh, mi, ss = m.groups()
        if hh is None or (hh, mi, ss) == ("00", "00", "00"):
            return f"{d}/{mo}/{y}"
        return f"{d}/{mo}/{y} {hh}:{mi}:{ss}"
    if _re.fullmatch(r"-?\d+\.0", s):
        s = s[:-2]
    return s


def parse_br_date(s: str) -> date | None:
    m = _re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", (s or "").strip())
    if not m:
        return None
    try:
        d = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None
    return d if date(1900, 1, 1) <= d <= date.today() else None


# ──────────────────────────────────────────────────────────────
#  ESQUEMA DE CAMPOS DO FORMULÁRIO (chave = cabeçalho da planilha)
# ──────────────────────────────────────────────────────────────
def F(label: str, kind: str, **kw) -> dict:
    return {"label": label, "kind": kind, **kw}


COUNT_OPTS = ["NENHUM (A)", "1", "2", "3", "4 OU MAIS"]
INT_IF_DIGITS = {  # gravados como número quando digitados só com dígitos
    "NÚMERO DE REFERÊNCIA", "Nº CASA",
    "OS MORADORES DA CASA [QUANTOS ADULTOS MORAM NA CASA?]",
    "OS MORADORES DA CASA [QUANTOS ADULTOS TRABALHAM ?]",
    "OS MORADORES DA CASA [QUANTAS CRIANÇAS MORAM NA CASA ?]",
    "OS MORADORES DA CASA [QUANTAS CRIANÇAS ESTUDAM NA CASA ?]",
}

COL_CESTAS = "PRECISA RECEBER CESTAS BÁSICAS DA PASTORAL SOCIAL? [AVALIAÇÃO]"
COL_REFEICOES = "TODOS OS MORADORES FAZEM AS 3 REFEIÇÕES DIÁRIAS ?"

FIELD_SPECS: dict[str, dict] = {
    # Identificação
    "NUMERO": F("Número do registro", "locked", help="Identificador do cadastro — não pode ser editado."),
    "NOME": F("Nome completo", "text"),
    "CPF": F("CPF", "text", ph="000.000.000-00"),
    "TIPO": F("Tipo", "select", options=["Beneficiário", "Enfermo", "Vulnerável", "Voluntário"]),
    "CID 2026": F("CID 2026", "text"),
    "APELIDO - CONHECIDO NA COMUNIDADE COMO?": F("Apelido (conhecido na comunidade como)", "text"),
    "NOME DA MÃE:": F("Nome da mãe", "text"),
    "DATA DE NASCIMENTO": F("Data de nascimento", "date"),
    "ESTADO CIVIL": F("Estado civil", "select", options=[
        "SOLTEIRO (A)", "CASADO(A)", "VIVE EM UNIÃO", "VIVE EM UNIÃO/ NAMORA",
        "VIÚVO(A)", "DIVORCIADO(A)"]),
    # Reservas
    "RESERVA 1": F("Reserva 1 — nome", "text"),
    "CPF RESERVA 1": F("Reserva 1 — CPF", "text", ph="000.000.000-00"),
    "RESERVA 2": F("Reserva 2 — nome", "text"),
    "CPF RESERVA 2": F("Reserva 2 — CPF", "text", ph="000.000.000-00"),
    # Endereço e contato
    "RUA PRINCIPAL": F("Rua principal", "text"),
    "NÚMERO DE REFERÊNCIA": F("Número de referência", "text"),
    "VIELA": F("Viela", "text"),
    "Nº CASA": F("Nº da casa", "text"),
    "REFERÊNCIA": F("Ponto de referência", "textarea"),
    "TELEFONE": F("Telefone", "text", ph="DDD + número"),
    # Trabalho e saúde
    "ESTÁ TRABALHANDO ATUALMENTE ?": F("Está trabalhando atualmente?", "select", options=[
        "SIM - EMPREGO FIXO", "SIM - REGULARMENTE", "SIM - BICOS SEM REGULARIDADE",
        "BICOS SEM REGULARIDADE", "NÃO ESTOU TRABALHANDO", "NÃO"]),
    "POSSUI ALGUMA DOENÇA GRAVE, OU OUTRO MOTIVO QUE IMPESSA DE TRABALHAR ? QUAL ?": F(
        "Possui doença grave ou outro motivo que impeça de trabalhar? Qual?", "textarea"),
    "JÁ RECEBEU CESTAS BÁSICAS DA PASTORAL SOCIAL ?": F(
        "Já recebeu cestas básicas da Pastoral Social?", "select",
        options=["RECEBE ATUALMENTE", "JÁ RECEBI NO PASSADO"]),
    # Moradores
    "OS MORADORES DA CASA [QUANTOS ADULTOS MORAM NA CASA?]": F("Quantos adultos moram na casa?", "select", options=COUNT_OPTS),
    "OS MORADORES DA CASA [QUANTOS ADULTOS TRABALHAM ?]": F("Quantos adultos trabalham?", "select", options=COUNT_OPTS),
    "OS MORADORES DA CASA [QUANTAS CRIANÇAS MORAM NA CASA ?]": F("Quantas crianças moram na casa?", "select", options=COUNT_OPTS),
    "OS MORADORES DA CASA [QUANTAS CRIANÇAS ESTUDAM NA CASA ?]": F("Quantas crianças estudam?", "select", options=COUNT_OPTS),
    # Condições da casa
    "CONDIÇÕES DA CASA [A CASA POSSUI BANHEIRO?]": F("A casa possui banheiro?", "yesno"),
    "CONDIÇÕES DA CASA [CHOVE DENTRO DA CASA?]": F("Chove dentro da casa?", "yesno"),
    "CONDIÇÕES DA CASA [EXISTE CAMA PARA TODOS?]": F("Existe cama para todos?", "yesno"),
    "CONDIÇÕES DA CASA [POSSUI LUZ ELÉTRICA?]": F("Possui luz elétrica?", "yesno"),
    "CONDIÇÕES DA CASA [POSSUI ÁGUA ENCANADA?]": F("Possui água encanada?", "yesno"),
    "CONDIÇÕES DA CASA [ALGUÉM POSSUI ALGUMA DOENÇA FÍSICA OU MENTAL GRAVE?]": F(
        "Alguém possui doença física ou mental grave?", "textarea"),
    "INFORMAÇÕES IMPORTANTES SOBRE A ESTRUTURA DA CASA E A SAÚDE DOS MORADORES": F(
        "Informações importantes sobre a estrutura da casa e a saúde dos moradores", "textarea"),
    "ELETRODOMÉSTICOS E FACILIDADES": F("Eletrodomésticos e facilidades", "multi", note="Pode marcar mais de um", options=[
        "FOGÃO", "GELADEIRA", "TELEVISÃO", "MÁQUINA DE LAVAR", "MICROONDAS", "INTERNET", "COMPUTADOR"]),
    # Contas da casa
    "AS CONTAS DA CASA [CASA PRÓPRIA?]": F("Casa própria?", "yesno"),
    "AS CONTAS DA CASA [PAGA ALUGUEL?]": F("Paga aluguel?", "yesno"),
    "AS CONTAS DA CASA [PAGA CONTA DE ÁGUA?]": F("Paga conta de água?", "yesno"),
    "AS CONTAS DA CASA [PAGA CONTA DE LUZ?]": F("Paga conta de luz?", "yesno"),
    "AS CONTAS DA CASA [RECEBE ALGUM AUXÍLIO DO GOVERNO?]": F(
        "Recebe algum auxílio do governo?", "select",
        options=["BOLSA FAMÍLIA", "APOSENTADORIA", "SIM", "NÃO RECEBO", "NÃO"]),
    "AS CONTAS DA CASA [ALTO GASTO COM MEDICAMENTOS?]": F("Alto gasto com medicamentos?", "yesno"),
    "AS CONTAS DA CASA [PAGA CONTA DE CELULAR?]": F("Paga conta de celular?", "yesno"),
    "AS CONTAS DA CASA [PAGA INTERNET OU TV?]": F("Paga internet ou TV?", "yesno"),
    "AS CONTAS DA CASA [PAGA GÁS?]": F("Paga gás?", "yesno"),
    "AS CONTAS DA CASA [PAGA INTERNET?]": F("Paga internet?", "yesno"),
    # Avaliação
    COL_REFEICOES: F("Todos os moradores fazem as 3 refeições diárias?", "yesno"),
    COL_CESTAS: F("Precisa receber cestas básicas da Pastoral Social?", "choice",
                  options=["PRECISA MUITO", "PRECISA", "NÃO PRECISA"]),
    "OBSERVAÇÕES INTERNAS": F("Observações internas", "textarea"),
    "TIPO DE BENEFICIÁRIO [TIPO]": F("Tipo de beneficiário", "select",
                                     options=["NORMAL", "ENFERMO", "VULNERÁVEL", "VOLUNTÁRIO"]),
    "AVALIAÇÃO GERAL DO ENTREVISTADOR": F("Avaliação geral do entrevistador", "textarea", h=160),
    "COMPLEMENTO ANALISE LAUDOS E CID": F("Complemento: análise de laudos e CID", "textarea"),
}

# Seções (acordeões) na ordem de exibição — tela mobile: um toque abre/fecha cada tema.
SECTIONS_TOP = [
    ("🪪 Identificação", ["NUMERO", "NOME", "CPF", "TIPO", "CID 2026",
                          "APELIDO - CONHECIDO NA COMUNIDADE COMO?", "NOME DA MÃE:",
                          "DATA DE NASCIMENTO", "ESTADO CIVIL"]),
    ("👥 Reservas", ["RESERVA 1", "CPF RESERVA 1", "RESERVA 2", "CPF RESERVA 2"]),
    ("📍 Endereço e contato", ["RUA PRINCIPAL", "NÚMERO DE REFERÊNCIA", "VIELA", "Nº CASA",
                               "REFERÊNCIA", "TELEFONE"]),
    ("💼 Trabalho e saúde", ["ESTÁ TRABALHANDO ATUALMENTE ?",
                             "POSSUI ALGUMA DOENÇA GRAVE, OU OUTRO MOTIVO QUE IMPESSA DE TRABALHAR ? QUAL ?",
                             "JÁ RECEBEU CESTAS BÁSICAS DA PASTORAL SOCIAL ?"]),
    ("👨‍👩‍👧 Moradores da casa", [
        "OS MORADORES DA CASA [QUANTOS ADULTOS MORAM NA CASA?]",
        "OS MORADORES DA CASA [QUANTOS ADULTOS TRABALHAM ?]",
        "OS MORADORES DA CASA [QUANTAS CRIANÇAS MORAM NA CASA ?]",
        "OS MORADORES DA CASA [QUANTAS CRIANÇAS ESTUDAM NA CASA ?]",
        COL_REFEICOES]),
    ("🏠 Condições da casa", [
        "CONDIÇÕES DA CASA [A CASA POSSUI BANHEIRO?]", "CONDIÇÕES DA CASA [CHOVE DENTRO DA CASA?]",
        "CONDIÇÕES DA CASA [EXISTE CAMA PARA TODOS?]", "CONDIÇÕES DA CASA [POSSUI LUZ ELÉTRICA?]",
        "CONDIÇÕES DA CASA [POSSUI ÁGUA ENCANADA?]",
        "CONDIÇÕES DA CASA [ALGUÉM POSSUI ALGUMA DOENÇA FÍSICA OU MENTAL GRAVE?]",
        "INFORMAÇÕES IMPORTANTES SOBRE A ESTRUTURA DA CASA E A SAÚDE DOS MORADORES",
        "ELETRODOMÉSTICOS E FACILIDADES"]),
    ("💰 Contas da casa", [
        "AS CONTAS DA CASA [CASA PRÓPRIA?]", "AS CONTAS DA CASA [PAGA ALUGUEL?]",
        "AS CONTAS DA CASA [PAGA CONTA DE ÁGUA?]", "AS CONTAS DA CASA [PAGA CONTA DE LUZ?]",
        "AS CONTAS DA CASA [RECEBE ALGUM AUXÍLIO DO GOVERNO?]",
        "AS CONTAS DA CASA [ALTO GASTO COM MEDICAMENTOS?]", "AS CONTAS DA CASA [PAGA CONTA DE CELULAR?]",
        "AS CONTAS DA CASA [PAGA INTERNET OU TV?]", "AS CONTAS DA CASA [PAGA GÁS?]",
        "AS CONTAS DA CASA [PAGA INTERNET?]"]),
]
SECTION_BOTTOM = ("📝 Observações e avaliação do entrevistador", [
    "OBSERVAÇÕES INTERNAS", "TIPO DE BENEFICIÁRIO [TIPO]",
    "AVALIAÇÃO GERAL DO ENTREVISTADOR", "COMPLEMENTO ANALISE LAUDOS E CID"])


# ──────────────────────────────────────────────────────────────
#  CARGA DE DADOS
# ──────────────────────────────────────────────────────────────
def _fetch_gdrive(url: str) -> bytes:
    """Baixa arquivo do Drive ('uc?export=download'), tratando a página de confirmação."""
    session = requests.Session()
    resp = session.get(url, timeout=15)
    resp.raise_for_status()
    if "text/html" in resp.headers.get("Content-Type", ""):
        match = _re.search(r"confirm=([0-9A-Za-z_\-]+)", resp.text)
        token = match.group(1) if match else "t"
        resp = session.get(url + f"&confirm={token}", timeout=30)
        resp.raise_for_status()
    return resp.content


@st.cache_data(ttl=300, show_spinner=False)
def load_data(url: str):
    """Retorna (df, erro). df: todas as colunas como texto; 52 colunas (50 da planilha + 2 de auditoria)."""
    try:
        content = (_download_drive(_get_drive_service(), st.secrets["GDRIVE_FILE_ID"])
                   if _gdrive_write_enabled() else _fetch_gdrive(url))
        return read_workbook(content), None
    except Exception as exc:
        return None, str(exc)


def read_workbook(content):
    raw = pd.read_excel(BytesIO(content), header=None, dtype=str)
    if raw.shape[0] < 3:
        raise ValueError("Arquivo sem dados suficientes.")

    # Garante 52 colunas (as de auditoria podem não existir ainda)
    while raw.shape[1] < AUDIT_TS_POS:
        raw[raw.shape[1]] = None
    raw = raw.iloc[:, :AUDIT_TS_POS]

    names = []
    for i, h in enumerate(raw.iloc[1].tolist(), start=1):
        if i == AUDIT_NAME_POS:
            names.append(AUDIT_NAME_COL)
        elif i == AUDIT_TS_POS:
            names.append(AUDIT_TS_COL)
        else:
            names.append(norm_header(h if pd.notna(h) else "") or f"COLUNA {i}")

    df = raw.iloc[2:].copy()
    df.columns = names
    df = df.apply(lambda s: s.map(norm_cell)).reset_index(drop=True)
    # remove linhas totalmente vazias
    df = df[~(df.drop(columns=[AUDIT_NAME_COL, AUDIT_TS_COL]).eq("").all(axis=1))].copy()
    return df


def _gdrive_write_enabled() -> bool:
    try:
        return "GDRIVE_FILE_ID" in st.secrets and "GDRIVE_SA_CREDENTIALS" in st.secrets
    except Exception:
        return False


def _get_drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    creds_info = st.secrets["GDRIVE_SA_CREDENTIALS"]
    if isinstance(creds_info, str):
        creds_info = json.loads(creds_info)
    else:
        creds_info = dict(creds_info)
    creds = Credentials.from_service_account_info(
        creds_info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    import httplib2
    from google_auth_httplib2 import AuthorizedHttp

    http = AuthorizedHttp(creds, http=httplib2.Http(timeout=20))
    return build("drive", "v3", http=http, cache_discovery=False)


def _suffixes():
    """A, B, …, Z, AA, AB, …"""
    n = 0
    while True:
        n, k, out = n + 1, n + 1, ""
        while k:
            k, r = divmod(k - 1, 26)
            out = chr(65 + r) + out
        yield out


def _from_display(col: str, text: str):
    """Texto de exibição → valor de célula (usado só quando não há linha atual para copiar)."""
    if text == "":
        return None
    if col in INT_IF_DIGITS and _re.fullmatch(r"\d{1,9}", text):
        return int(text)
    if col == "DATA DE NASCIMENTO" and parse_br_date(text):
        return datetime.combine(parse_br_date(text), time())
    return text


def patch_workbook(wb, numero: str, changes: dict, editor: str, ts: datetime,
                   user_row: dict | None = None) -> tuple[str | None, str | None]:
    """
    Grava as alterações de UM cadastro e, nas colunas 51/52, o nome de quem salvou e o timestamp.
    Devolve (erro, numero_novo).

    • Sem conflito: atualiza a linha do cadastro (localizada pelo NUMERO, não por posição).
    • Com conflito (outra pessoa alterou o mesmo campo, ou o cadastro sumiu/duplicou na planilha):
      NÃO recusa — acrescenta uma linha ao final com a linha atual + as alterações do usuário e
      Nº com sufixo de letra (220 → 220-A, 220-B…). Nome e timestamp também são gravados.
    """
    from copy import copy

    ws = wb.active
    hdr: dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        h = norm_header(ws.cell(2, c).value)
        if h:
            hdr.setdefault(h, c)

    ncol = hdr.get("NUMERO")
    if ncol is None:
        return "Coluna NUMERO não encontrada na planilha.", None
    for col in changes:
        if col not in hdr:
            return f"Coluna “{col}” não encontrada na planilha.", None

    rows = [r for r in range(3, ws.max_row + 1) if norm_cell(ws.cell(r, ncol).value) == numero]
    conflict = len(rows) != 1
    if not conflict:
        for col, ch in changes.items():
            atual = norm_cell(ws.cell(rows[0], hdr[col]).value)
            if atual != ch["old"] and atual != ch["new"]:
                conflict = True
                break

    novo = None
    if not conflict:
        target = rows[0]
    else:
        base = _re.sub(r"-[A-Z]+$", "", numero)
        existentes = {norm_cell(ws.cell(r, ncol).value) for r in range(3, ws.max_row + 1)}
        novo = next(f"{base}-{x}" for x in _suffixes() if f"{base}-{x}" not in existentes)

        nome_col = hdr.get("NOME", 5)
        last = max((r for r in range(3, ws.max_row + 1)
                    if norm_cell(ws.cell(r, ncol).value) or norm_cell(ws.cell(r, nome_col).value)),
                   default=2)
        target = last + 1
        src = rows[0] if len(rows) == 1 else None
        for c in range(1, AUDIT_TS_POS + 1):  # estilo da última linha
            if ws.cell(last, c).has_style:
                ws.cell(target, c)._style = copy(ws.cell(last, c)._style)
        for col, c in hdr.items():  # base: linha atual da planilha (ou a visão do usuário)
            if c > 50:
                continue
            if src is not None:
                ws.cell(target, c).value = ws.cell(src, c).value
            elif user_row and col in user_row:
                ws.cell(target, c).value = _from_display(col, user_row[col])
        ws.cell(target, ncol).value = novo

    for col, ch in changes.items():
        cell = ws.cell(target, hdr[col])
        cell.value = ch["value"]
        if isinstance(ch["value"], datetime):
            cell.number_format = "dd/mm/yyyy"

    # ── Colunas 51 e 52: cabeçalhos (se ainda não existirem) + valores
    for pos, title in ((AUDIT_NAME_POS, AUDIT_NAME_COL), (AUDIT_TS_POS, AUDIT_TS_COL)):
        for r_hdr, val in ((1, pos), (2, title)):
            cell = ws.cell(r_hdr, pos)
            if cell.value in (None, ""):
                cell.value = val
                model = ws.cell(r_hdr, 50)
                if model.has_style:
                    cell.font, cell.fill = copy(model.font), copy(model.fill)
                    cell.border, cell.alignment = copy(model.border), copy(model.alignment)
    ws.cell(target, AUDIT_NAME_POS).value = editor
    ws.cell(target, AUDIT_TS_POS).value = ts
    ws.cell(target, AUDIT_TS_POS).number_format = AUDIT_TS_FORMAT

    # Na planilha original as colunas ≥ 51 estão ocultas (AY:XFD) → exibe 51 e 52.
    from openpyxl.worksheet.dimensions import ColumnDimension

    dims = ws.column_dimensions
    first = dims["AY"] if "AY" in dims else None
    if first is not None and first.hidden and (first.max or 51) > AUDIT_TS_POS:
        dims["BA"] = ColumnDimension(ws, index="BA", min=AUDIT_TS_POS + 1, max=first.max, hidden=True)
    for letter, idx, width in (("AY", AUDIT_NAME_POS, 26), ("AZ", AUDIT_TS_POS, 20)):
        d = dims[letter]
        d.min, d.max, d.hidden, d.width = idx, idx, False, width
    return None, novo


def _download_drive(service, file_id):
    from googleapiclient.http import MediaIoBaseDownload

    buf = BytesIO()
    downloader = MediaIoBaseDownload(buf, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def save_batch_to_drive(operations):
    """Uma leitura, todas as alterações em ordem, um único upload do XLSX."""
    import openpyxl
    from googleapiclient.http import MediaIoBaseUpload

    with SYNC_LOCK:
        service = _get_drive_service()
        file_id = st.secrets["GDRIVE_FILE_ID"]
        wb = openpyxl.load_workbook(BytesIO(_download_drive(service, file_id)))
        columns = {norm_header(c.value) for c in wb.active[2][:50] if c.value}
        validated = validate_operations(operations, columns, _from_display, SP_TZ)
        results = apply_batch(wb, validated, patch_workbook)
        out = BytesIO()
        wb.save(out)
        content = out.getvalue()
        out.seek(0)
        media = MediaIoBaseUpload(
            out, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False,
        )
        service.files().update(fileId=file_id, media_body=media).execute()
        load_data.clear()
        return results, read_workbook(content).to_dict(orient="records")


# A interface de consulta/edição executa no navegador, inclusive sem WebSocket.
offline_app = components.declare_component(
    "offline_visitas", path=str(Path(__file__).with_name("offline_frontend"))
)

st.markdown(
    "<div class='ps-header'><h1>📋 Cadastro de Visitas</h1>"
    "<p>Pastoral Social · Jardim Colombo</p></div>", unsafe_allow_html=True,
)
with st.spinner("Carregando cadastro…"):
    _base_df, load_error = load_data(GITHUB_RAW_URL)
if load_error:
    # Não guarda falhas por cinco minutos; a próxima tentativa deve consultar a rede.
    load_data.clear()
    st.warning("Sem acesso à planilha. Usando os dados guardados neste aparelho, se disponíveis.")
if not _gdrive_write_enabled():
    st.warning("Drive não configurado. As edições ficam pendentes neste aparelho até ser possível enviar.")

specs = dict(FIELD_SPECS)
if _base_df is not None:
    for col in _base_df.columns:
        if col not in specs and col not in (AUDIT_NAME_COL, AUDIT_TS_COL):
            specs[col] = F(col.title(), "text")
source_id = GITHUB_RAW_URL
dataset = hashlib.sha256(source_id.encode()).hexdigest()[:24]
response = st.session_state.get("sync_response")
event = offline_app(
    dataset=dataset,
    records=_base_df.to_dict(orient="records") if _base_df is not None else None,
    specs=specs, sections=SECTIONS_TOP + [("⭐ Avaliação de cestas", [COL_CESTAS]), SECTION_BOTTOM],
    audit_name=AUDIT_NAME_COL, audit_ts=AUDIT_TS_COL,
    write_enabled=_gdrive_write_enabled(), response=response,
    default=None, key="offline_visitas_main",
)
if isinstance(event, dict) and event.get("type") == "sync":
    request_id = event.get("request_id")
    if request_id and (not response or response.get("request_id") != request_id):
        try:
            if not _gdrive_write_enabled():
                raise ValueError("Configure as credenciais do Drive para enviar. As pendências continuam neste aparelho.")
            with st.spinner("Enviando todas as alterações em um único lote…"):
                targets, records = save_batch_to_drive(event.get("operations"))
            st.session_state.sync_response = {
                "request_id": request_id, "ok": True, "targets": targets, "records": records,
            }
        except Exception as exc:
            st.session_state.sync_response = {
                "request_id": request_id, "ok": False,
                "error": f"Não foi possível enviar: {exc}. As alterações continuam no cache deste aparelho.",
            }
        st.rerun()
