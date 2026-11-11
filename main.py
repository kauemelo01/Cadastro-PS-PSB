"""
Pastoral Social — Cadastro de Visitas
Mobile-first Streamlit app: busca por Número / Nome / CPF e edição completa do cadastro.

Planilha: Visita_2026.xlsx (aba "CADASTRO GERAL")
  • Linha 1: numeração das colunas      • Linha 2: cabeçalhos      • Dados a partir da linha 3
  • Colunas 51 e 52 (criadas pelo app): ÚLTIMA EDIÇÃO POR / ÚLTIMA EDIÇÃO EM

Secrets (Streamlit Cloud → Settings → Secrets):
  GITHUB_RAW_URL         = "<url de leitura do xlsx>"       (opcional, há um padrão abaixo)
  GDRIVE_FILE_ID         = "<id do arquivo .xlsx no Drive>"  (habilita salvar no Drive)
  GDRIVE_SA_CREDENTIALS  = "<json da service account>"       (habilita salvar no Drive)
"""

import html as _html
import hashlib
import json
import os
import re as _re
import sqlite3
import threading
import unicodedata as _ud
import uuid
from contextlib import closing
from datetime import date, datetime, time, timedelta, timezone
from io import BytesIO
from pathlib import Path

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

CPF_COLS = {"CPF", "CPF RESERVA 1", "CPF RESERVA 2"}

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
#  MOBILE-FIRST CSS
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
  .main .block-container { padding: 0.75rem 0.75rem 2rem; max-width: 500px; margin: auto; }

  /* Header */
  .ps-header {
    background: linear-gradient(135deg, #1b5e20 0%, #388e3c 100%);
    border-radius: 14px; padding: 1rem 1.25rem 0.75rem; color: #fff; margin-bottom: 1rem;
  }
  .ps-header h1 { margin: 0; font-size: 1.35rem; color: #fff; }
  .ps-header p  { margin: 0.15rem 0 0; font-size: 0.82rem; opacity: 0.85; }

  /* Search box */
  .stTextInput > div > div > input {
    font-size: 16px !important; border-radius: 12px !important;
    border: 2px solid #c8e6c9 !important; padding: 0.65rem 1rem !important;
    background: #f9fbe7 !important; color: #212121 !important;
  }
  .stTextInput > div > div > input:focus {
    border-color: #388e3c !important; box-shadow: 0 0 0 3px rgba(56,142,60,.15) !important;
  }

  /* Result card */
  .card {
    background: #fff; border-radius: 14px; padding: 1rem 1.1rem; margin-bottom: 0.5rem;
    box-shadow: 0 2px 10px rgba(0,0,0,.08); border-left: 5px solid #388e3c;
  }
  .card-name { font-size: 1.05rem; font-weight: 700; color: #1b5e20; margin-bottom: 5px; }
  .card-num  { font-size: 0.78rem; color: #757575; margin-bottom: 8px; }
  .card-audit { font-size: 0.72rem; color: #9e9e9e; margin-top: 10px; }

  /* Badges */
  .badge { display: inline-block; padding: 3px 11px; border-radius: 20px;
           font-size: 0.77rem; font-weight: 700; margin: 2px 3px 2px 0; }
  .b-tipo   { background: #e3f2fd; color: #1565c0; }
  .b-need   { background: #fff3e0; color: #e65100; }
  .b-need2  { background: #ffebee; color: #b71c1c; }
  .b-noneed { background: #eeeeee; color: #616161; }

  /* Info rows */
  .info-grid { margin-top: 10px; border-top: 1px solid #f0f0f0; padding-top: 8px; }
  .info-row  { display: flex; align-items: flex-start; padding: 5px 0;
               border-bottom: 1px solid #f9f9f9; gap: 8px; }
  .info-label { font-size: 0.75rem; font-weight: 700; text-transform: uppercase;
                letter-spacing: 0.03em; color: #9e9e9e; min-width: 80px; flex-shrink: 0; padding-top: 3px; }
  .info-value { font-size: 0.9rem; color: #212121; flex: 1; word-break: break-word; }
  .info-empty { color: #bdbdbd; font-style: italic; }

  mark.hl { background: #fff176; color: #212121; border-radius: 3px; padding: 0 2px; font-weight: 700; }

  .empty-state { text-align: center; padding: 3rem 1rem; color: #9e9e9e; }
  .empty-state .icon { font-size: 2.8rem; }
  .empty-state p { margin: 0.5rem 0 0; font-size: 0.95rem; }
  .result-count { font-size: 0.82rem; color: #757575; margin-bottom: 0.6rem; padding-left: 2px; }

  /* ── Popup de edição ── */
  div[data-testid="stDialog"] div[role="dialog"] { max-width: 640px; }
  div[data-testid="stDialog"] details summary p { font-weight: 700; font-size: 1.02rem; }
  div[data-testid="stDialog"] input, div[data-testid="stDialog"] textarea { font-size: 16px !important; }
  div[data-testid="stDialog"] [data-testid="stButtonGroup"] button { min-height: 2.6rem; font-weight: 700; }

  /* Destaque: precisa receber cestas? */
  .st-key-cestas_box {
    background: #fff8e1; border: 2px solid #f9a825; border-radius: 14px; padding: 0.9rem 1rem 0.6rem;
  }
  .cestas-title { font-size: 0.8rem; font-weight: 800; letter-spacing: .04em; color: #e65100; margin-bottom: 4px; }
  .st-key-cestas_box [role="radiogroup"] { gap: 0.45rem; width: 100%; }
  .st-key-cestas_box [role="radiogroup"] label {
    width: 100%; box-sizing: border-box; background: #fff; border: 2px solid #f9a825;
    border-radius: 12px; padding: 0.7rem 0.9rem; margin: 0;
  }
  .st-key-cestas_box [role="radiogroup"] label p { font-weight: 800; font-size: 1.02rem; }

  /* Popup de confirmação */
  .chg-list { max-height: 38vh; overflow-y: auto; border: 1px solid #e0e0e0; border-radius: 10px; padding: 4px 10px; }
  .chg-item { padding: 7px 0; border-bottom: 1px solid #f0f0f0; font-size: 0.85rem; }
  .chg-item:last-child { border-bottom: none; }
  .chg-label { font-weight: 700; color: #1b5e20; font-size: 0.78rem; text-transform: uppercase; letter-spacing: .03em; }
  .chg-old { color: #b71c1c; text-decoration: line-through; word-break: break-word; }
  .chg-new { color: #2e7d32; font-weight: 600; word-break: break-word; }

  /* Dark mode */
  @media (prefers-color-scheme: dark) {
    .stTextInput > div > div > input { background: #1a2e1a !important; color: #e8f5e9 !important; border-color: #4caf50 !important; }
    .stTextInput > div > div > input::placeholder { color: #81c784 !important; opacity: 0.7; }
    .card { background: #1e2b1e !important; }
    .info-value { color: #e0e0e0 !important; }
    .card-name { color: #a5d6a7 !important; }
    .card-num, .card-audit { color: #9e9e9e !important; }
    .info-grid { border-top-color: #2e3e2e !important; }
    .info-row  { border-bottom-color: #2a3a2a !important; }
    .st-key-cestas_box { background: #3e2e10 !important; border-color: #f9a825 !important; }
    .cestas-title { color: #ffcc80 !important; }
    .st-key-cestas_box [role="radiogroup"] label { background: #2a2110 !important; }
    .chg-list, .chg-item { border-color: #2e3e2e !important; }
    .chg-label { color: #a5d6a7 !important; }
    .chg-old { color: #ef9a9a !important; }
    .chg-new { color: #a5d6a7 !important; }
    .b-noneed { background: #2a352a !important; color: #bdbdbd !important; }
  }

  #MainMenu { visibility: hidden; }
  footer    { visibility: hidden; }
  header    { visibility: hidden; }
</style>
""",
    unsafe_allow_html=True,
)


def show_offline_map() -> None:
    """Exibe o HTML versionado junto ao app, sem carregar os cadastros ou acessar o Drive."""
    import streamlit.components.v1 as components

    st.markdown('[← Voltar ao cadastro](./)')
    map_path = Path(__file__).resolve().parent / "mapa_enderecos_openstreetmap.html"
    try:
        map_html = map_path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        st.error("Mapa offline não encontrado. Coloque o arquivo "
                 "mapa_enderecos_openstreetmap.html na mesma pasta do main.py no repositório.")
    except (OSError, UnicodeError):
        st.error("Não foi possível abrir o mapa offline. Verifique o arquivo "
                 "mapa_enderecos_openstreetmap.html no repositório.")
    else:
        components.html(map_html, height=850, scrolling=True)


if st.query_params.get("mapa") == "offline":
    show_offline_map()
    st.stop()


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


def format_cpf_if_plain(text: str) -> str:
    """'12345678901' ou '123.456.789-01' → '123.456.789-01'. Outros formatos ficam como digitados."""
    m = _re.fullmatch(r"(\d{3})\.?(\d{3})\.?(\d{3})-?(\d{2})", text.strip())
    return f"{m.group(1)}.{m.group(2)}.{m.group(3)}-{m.group(4)}" if m else text


def md_esc(s: str) -> str:
    return _re.sub(r"([\\`*_{}\[\]()#+!|~$<>])", r"\\\1", s)


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
#  CACHE PERSISTENTE DE ALTERAÇÕES PENDENTES (no servidor do app)
# ──────────────────────────────────────────────────────────────
def _drive_file_id() -> str:
    try:
        return str(st.secrets.get("GDRIVE_FILE_ID", "")).strip()
    except Exception:
        return ""


def _cache_path() -> Path:
    folder = Path(os.environ.get("CADASTRO_CACHE_DIR") or
                  Path(__file__).resolve().parent / ".cadastro_cache")
    source = hashlib.sha256(GITHUB_RAW_URL.encode("utf-8")).hexdigest()[:24]
    return folder / f"pendencias-{source}.sqlite3"


def _cache_json_default(value):
    if isinstance(value, (datetime, date)):
        return {"__cache_type__": type(value).__name__, "value": value.isoformat()}
    raise TypeError(f"Tipo não suportado no cache: {type(value).__name__}")


def _cache_json_hook(value):
    if value.get("__cache_type__") == "datetime":
        return datetime.fromisoformat(value["value"])
    if value.get("__cache_type__") == "date":
        return date.fromisoformat(value["value"])
    return value


def _open_cache():
    path = _cache_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), timeout=30)
    try:
        conn.execute("""CREATE TABLE IF NOT EXISTS pending (
            seq INTEGER PRIMARY KEY AUTOINCREMENT,
            id TEXT NOT NULL UNIQUE,
            file_id TEXT NOT NULL,
            payload TEXT NOT NULL
        )""")
        conn.commit()
        return conn
    except Exception:
        conn.close()
        raise


def pending_changes() -> list[dict]:
    """Fila compartilhada pelas sessões deste app, isolada por origem e arquivo do Drive."""
    if not _cache_path().exists():
        return []
    with closing(_open_cache()) as conn:
        rows = conn.execute(
            "SELECT payload FROM pending WHERE file_id IN ('', ?) ORDER BY seq",
            (_drive_file_id(),),
        ).fetchall()
    return [json.loads(row[0], object_hook=_cache_json_hook) for row in rows]


def cache_change(operation: dict) -> None:
    payload = json.dumps(operation, ensure_ascii=False, default=_cache_json_default)
    with closing(_open_cache()) as conn, conn:
        conn.execute("INSERT OR IGNORE INTO pending (id, file_id, payload) VALUES (?, ?, ?)",
                     (operation["id"], _drive_file_id(), payload))


def _bind_cached_changes(operations: list[dict]) -> None:
    # Vincula pendências criadas sem credenciais antes da primeira tentativa de envio.
    # Assim, uma troca posterior do arquivo de destino não reaplica esse lote em outro Drive.
    with closing(_open_cache()) as conn, conn:
        conn.executemany("UPDATE pending SET file_id = ? WHERE id = ? AND file_id = ''",
                         [(_drive_file_id(), op["id"]) for op in operations])


def _remove_cached_changes(operations: list[dict]) -> None:
    # Só os IDs confirmados são removidos; inclusões feitas durante o envio ficam na fila.
    with closing(_open_cache()) as conn, conn:
        conn.executemany("DELETE FROM pending WHERE id = ?", [(op["id"],) for op in operations])


def _cached_parent(operations: list[dict], numero: str, user_row: dict) -> str | None:
    """Liga edições sucessivas da mesma versão, inclusive se ela gerar uma linha com sufixo."""
    for op in reversed(operations):
        if (op["numero"] == numero and
                user_row.get(AUDIT_NAME_COL) == op["editor"] and
                user_row.get(AUDIT_TS_COL) == norm_cell(op["ts"]) and
                all(user_row.get(col) == ch["new"] for col, ch in op["changes"].items())):
            return op["id"]
    return None


@st.cache_resource
def _drive_lock():
    # Streamlit executa sessões em threads; serializa leitura/alteração/envio do xlsx.
    return threading.RLock()


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
        content = _fetch_gdrive(url)
        raw = pd.read_excel(BytesIO(content), header=None, dtype=str)
    except Exception as exc:
        return None, str(exc)

    if raw.shape[0] < 3:
        return None, "Arquivo sem dados suficientes."

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
    return df, None


with st.spinner("Carregando cadastro…"):
    _base_df, load_error = load_data(GITHUB_RAW_URL)

# ──────────────────────────────────────────────────────────────
#  SESSÃO: edições em memória (cobre o intervalo até o arquivo atualizar)
# ──────────────────────────────────────────────────────────────
if "edits" not in st.session_state:
    st.session_state.edits = {}  # {numero_original: {coluna: novo_texto}}


def with_overlay(base: pd.DataFrame) -> pd.DataFrame:
    """Aplica as edições feitas nesta sessão. Descarta as que o arquivo já reflete."""
    edits = st.session_state.edits
    if not edits or base is None:
        return base
    out = base.copy()
    pos = {v: i for i, v in out["NUMERO"].items()}
    for numero in list(edits):
        i = pos.get(numero)
        if i is None:
            del edits[numero]
            continue
        for col in list(edits[numero]):
            if col not in out.columns or out.at[i, col] == edits[numero][col]:
                del edits[numero][col]
            else:
                out.at[i, col] = edits[numero][col]
        if not edits[numero]:
            del edits[numero]
    return out


def with_pending_overlay(base: pd.DataFrame, operations: list[dict]) -> pd.DataFrame:
    """Mostra as pendências sem descartá-las até a confirmação do Drive."""
    if base is None or not operations:
        return base
    out = base.copy()
    for op in operations:
        mask = out["NUMERO"] == op["numero"]
        for col, ch in op["changes"].items():
            if col in out.columns:
                out.loc[mask, col] = ch["new"]
        out.loc[mask, AUDIT_NAME_COL] = op["editor"]
        out.loc[mask, AUDIT_TS_COL] = norm_cell(op["ts"])
    return out


try:
    _cached_changes, _cache_error = pending_changes(), None
except Exception as exc:
    _cached_changes, _cache_error = [], str(exc)

df = with_pending_overlay(with_overlay(_base_df), _cached_changes)

# ──────────────────────────────────────────────────────────────
#  GOOGLE DRIVE — GRAVAÇÃO
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
        creds_info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


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


_SYNC_SHEET = "_PS_SYNC_CACHE"
_SYNC_HEADERS = ("OPERATION_ID", "NUMERO", "DESTINO", "EDITOR", "TIMESTAMP")


def _patch_cached_batch(wb, operations: list[dict]) -> tuple[str | None, dict, bool]:
    """Aplica a fila em ordem e registra IDs para retomar envios com resposta perdida."""
    if _SYNC_SHEET in wb.sheetnames:
        journal = wb[_SYNC_SHEET]
        if tuple(journal.cell(1, c).value for c in range(1, 6)) != _SYNC_HEADERS:
            return f"A aba {_SYNC_SHEET} já existe com outro formato.", {}, False
    else:
        journal = wb.create_sheet(_SYNC_SHEET)
        journal.append(_SYNC_HEADERS)
        journal.sheet_state = "veryHidden"

    applied = {row[0]: row[2] for row in journal.iter_rows(min_row=2, values_only=True)}
    targets, modified = {}, False
    for op in operations:
        if op["id"] in applied:
            targets[op["id"]] = applied[op["id"]]
            continue
        numero = applied.get(op.get("parent_id"), op["numero"])
        err, novo = patch_workbook(wb, numero, op["changes"], op["editor"],
                                   op["ts"], op.get("row"))
        if err:
            return f"Cadastro Nº {op['numero']}: {err}", {}, False
        target = novo or numero
        journal.append((op["id"], op["numero"], target, op["editor"], op["ts"]))
        journal.cell(journal.max_row, 5).number_format = AUDIT_TS_FORMAT
        targets[op["id"]] = applied[op["id"]] = target
        modified = True
    return None, targets, modified


def save_changes_to_drive(operations: list[dict]) -> tuple[str | None, dict]:
    """Um download e um upload para TODO o lote; nenhuma exclusão local em caso de erro."""
    try:
        import openpyxl
        from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

        file_id = _drive_file_id()
        service = _get_drive_service()

        buf = BytesIO()
        downloader = MediaIoBaseDownload(buf, service.files().get_media(fileId=file_id))
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)

        wb = openpyxl.load_workbook(buf)
        err, targets, modified = _patch_cached_batch(wb, operations)
        if err:
            load_data.clear()
            return err, {}

        if modified:
            out = BytesIO()
            wb.save(out)
            out.seek(0)
            media = MediaIoBaseUpload(
                out,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                resumable=False,
            )
            service.files().update(fileId=file_id, media_body=media).execute()
        load_data.clear()
        return None, targets
    except Exception as exc:
        return str(exc), {}


def save_or_cache(pending: dict, editor: str) -> tuple[bool, str | None, str | None]:
    """Devolve (ficou_em_cache, erro_de_envio, destino). Erros locais mantêm o formulário aberto."""
    with _drive_lock():
        queued = pending_changes()
        if "operation" not in pending:
            pending["operation"] = {
                "id": uuid.uuid4().hex,
                "numero": pending["numero"],
                "changes": pending["changes"],
                "editor": editor,
                "ts": datetime.now(SP_TZ).replace(tzinfo=None, microsecond=0),
                "row": pending.get("row") or {},
                "parent_id": _cached_parent(queued, pending["numero"], pending.get("row") or {}),
            }
        operation = pending["operation"]
        # Grava antes da rede: uma queda do processo durante o upload não perde a edição.
        cache_change(operation)
        if any(op["id"] != operation["id"] for op in queued) or not _gdrive_write_enabled():
            return True, None, None
        _bind_cached_changes([operation])
        err, targets = save_changes_to_drive([operation])
        if err:
            return True, err, None
        _remove_cached_changes([operation])
        return False, None, targets[operation["id"]]


def sync_cached_changes() -> tuple[str | None, list[dict], dict]:
    """Envia a fila inteira preservando o autor, os tipos e o timestamp de cada confirmação."""
    with _drive_lock():
        try:
            operations = pending_changes()
            if not operations:
                return None, [], {}
            if not _gdrive_write_enabled():
                return "Configure as credenciais do Google Drive para enviar as pendências.", [], {}
            _bind_cached_changes(operations)
            err, targets = save_changes_to_drive(operations)
            if err:
                return err, [], {}
            _remove_cached_changes(operations)
            return None, operations, targets
        except Exception as exc:
            return str(exc), [], {}


def _overlay_saved_change(operation: dict, target: str | None = None) -> None:
    numero = operation["numero"]
    if target and target != numero:
        # Como na gravação original, a linha em conflito permanece intacta.
        st.session_state.edits.pop(numero, None)
        return
    ed = {col: ch["new"] for col, ch in operation["changes"].items()}
    ed[AUDIT_NAME_COL] = operation["editor"]
    ed[AUDIT_TS_COL] = norm_cell(operation["ts"])
    st.session_state.edits.setdefault(numero, {}).update(ed)


# ──────────────────────────────────────────────────────────────
#  FORMULÁRIO DE EDIÇÃO
# ──────────────────────────────────────────────────────────────
def _field_widget(col: str, spec: dict, initial: str, key: str, draft: dict):
    """Desenha o campo. Retorna (valor_inicial_do_widget, valor_atual_do_widget, texto_novo)."""
    kind, label = spec["kind"], spec["label"]
    help_ = spec.get("help")

    if kind == "locked":  # único campo bloqueado
        st.text_input(label, value=initial, key=key, disabled=True, help=help_)
        return initial, initial, initial

    if kind in ("text", "id"):
        val = st.text_input(label, value=draft.get(col, initial), key=key,
                            placeholder=spec.get("ph"), help=help_)
        return initial, val, val.strip()

    if kind == "textarea":
        val = st.text_area(label, value=draft.get(col, initial), key=key,
                           height=spec.get("h", 100), help=help_)
        return initial, val, val.strip()

    if kind == "select":
        opts = list(spec["options"])
        if initial and initial not in opts:
            opts.append(initial)  # preserva valores antigos fora da lista
        w_init = initial or None
        start = draft.get(col, w_init)
        idx = opts.index(start) if start in opts else None
        val = st.selectbox(label, opts, index=idx, key=key, placeholder="Selecione…")
        return w_init, val, val or ""

    if kind == "choice":  # lista vertical (cabe inteira em telas estreitas)
        opts = list(spec["options"])
        if initial and initial not in opts:
            opts.append(initial)
        w_init = initial or None
        start = draft.get(col, w_init)
        idx = opts.index(start) if start in opts else None
        val = st.radio(label, opts, index=idx, key=key, label_visibility="collapsed")
        return w_init, val, val or ""

    if kind == "yesno":
        opts = ["SIM", "NÃO"]
        if initial and initial not in opts:
            opts.append(initial)  # preserva valores antigos fora da lista
        w_init = initial or None
        start = draft.get(col, w_init)
        st.markdown(f"<div style='font-size:.875rem;margin-bottom:2px'>{_html.escape(label)}</div>",
                    unsafe_allow_html=True)
        val = st.segmented_control(label, opts, selection_mode="single", default=start,
                                   key=key, label_visibility="collapsed")
        return w_init, val, val or ""

    if kind == "date":
        w_init = parse_br_date(initial)
        val = st.date_input(label, value=draft.get(col, w_init), min_value=date(1900, 1, 1),
                            max_value=date.today(), format="DD/MM/YYYY", key=key)
        if initial and w_init is None:
            st.caption(f"Valor atual na planilha (formato não reconhecido): **{md_esc(initial)}**")
        return w_init, val, val.strftime("%d/%m/%Y") if val else ""

    if kind == "multi":
        toks = [t.strip() for t in _re.split(r"[,;]", initial) if t.strip()]
        opts = list(spec["options"]) + [t for t in toks if t not in spec["options"]]
        val = st.multiselect(label, opts, default=draft.get(col, toks), key=key,
                             placeholder="Selecione…")
        if spec.get("note"):
            st.caption(spec["note"])
        ordered = [o for o in opts if o in val]
        return toks, val, ", ".join(ordered)

    raise ValueError(kind)


def _differs(kind: str, w_init, w_val) -> bool:
    if kind == "multi":
        return set(w_init or []) != set(w_val or [])
    if kind in ("text", "id", "textarea"):
        return (w_init or "").strip() != (w_val or "").strip()
    return w_init != w_val


def _cell_value(col: str, spec: dict, new_text: str, widget_val):
    """Texto/valor do widget → valor a gravar na célula (preserva tipos numéricos e datas)."""
    if new_text == "":
        return None
    if spec["kind"] == "date" and isinstance(widget_val, date):
        return datetime.combine(widget_val, time())
    if col in INT_IF_DIGITS and _re.fullmatch(r"\d{1,9}", new_text):
        return int(new_text)
    return new_text


def _row_dict(numero: str) -> dict | None:
    hit = df[df["NUMERO"] == numero]
    return None if hit.empty else hit.iloc[0].to_dict()


@st.dialog("✏️ Editar Cadastro", width="large")
def edit_dialog(numero: str, row: dict, nonce: str, draft: dict | None = None):
    draft = draft or {}
    st.markdown(f"**{md_esc(row.get('NOME') or '—')}**  \nNº {md_esc(numero)}")
    st.caption("Toque em cada tema para abrir. Só os campos que você alterar serão gravados.")

    changes: dict[str, dict] = {}
    widgets: dict[str, object] = {}

    def draw(col: str):
        if col not in row:
            return
        spec = FIELD_SPECS[col]
        initial = row.get(col, "")
        w_init, w_val, new_text = _field_widget(col, spec, initial, f"ed_{nonce}_{col}", draft)
        widgets[col] = w_val
        if _differs(spec["kind"], w_init, w_val):
            if col in CPF_COLS:
                new_text = format_cpf_if_plain(new_text)
            changes[col] = {"old": initial, "new": new_text, "label": spec["label"],
                            "value": _cell_value(col, spec, new_text, w_val)}

    for i, (title, cols) in enumerate(SECTIONS_TOP):
        with st.expander(title, expanded=(i == 0)):
            for c in cols:
                draw(c)

    # Outras colunas que existam na planilha e não estejam no esquema acima
    extras = [c for c in row if c not in FIELD_SPECS and c not in (AUDIT_NAME_COL, AUDIT_TS_COL)]
    for c in extras:
        FIELD_SPECS.setdefault(c, F(c.title(), "text"))
    if extras:
        with st.expander("📎 Outros campos"):
            for c in extras:
                draw(c)

    # ── Destaque: avaliação de necessidade de cestas (entre os últimos campos)
    with st.container(key="cestas_box"):
        st.markdown("<div class='cestas-title'>⭐ AVALIAÇÃO — CESTAS BÁSICAS</div>", unsafe_allow_html=True)
        st.markdown("**Precisa receber cestas básicas da Pastoral Social?**")
        draw(COL_CESTAS)

    with st.expander(SECTION_BOTTOM[0], expanded=False):
        for c in SECTION_BOTTOM[1]:
            draw(c)

    st.write("")
    n = len(changes)
    st.caption(f"{n} campo{'s' if n != 1 else ''} alterado{'s' if n != 1 else ''}")
    c1, c2 = st.columns(2)
    if c1.button("💾 Salvar", type="primary", width="stretch", key=f"save_{nonce}"):
        errors = []
        if "NOME" in changes and not changes["NOME"]["new"]:
            errors.append("O nome não pode ficar em branco.")
        if not n:
            errors.append("Nenhuma alteração para salvar.")
        if errors:
            for e in errors:
                st.error(e)
        else:
            st.session_state.pending = {"numero": numero, "nome": row.get("NOME", ""),
                                        "changes": changes, "widgets": widgets, "row": row}
            st.session_state.open_dialog = "confirm"
            st.rerun()  # fecha este popup; o de confirmação abre em seguida
    if c2.button("Cancelar", width="stretch", key=f"cancel_{nonce}"):
        st.session_state.pop("pending", None)
        st.rerun()


@st.dialog("Confirmar alterações")
def confirm_dialog(pending: dict):
    numero, nome, changes = pending["numero"], pending["nome"], pending["changes"]
    n = len(changes)
    st.markdown(f"Você alterou **{n} campo{'s' if n != 1 else ''}** de "
                f"**{md_esc(nome or '—')}** · Nº {md_esc(numero)}:")

    def short(s: str) -> str:
        s = s if s else "(vazio)"
        return _html.escape(s if len(s) <= 140 else s[:137] + "…").replace("\n", " ⏎ ")

    items = "".join(
        f"<div class='chg-item'><div class='chg-label'>{_html.escape(ch['label'])}</div>"
        f"<span class='chg-old'>{short(ch['old'])}</span> → <span class='chg-new'>{short(ch['new'])}</span></div>"
        for ch in changes.values()
    )
    st.markdown(f"<div class='chg-list'>{items}</div>", unsafe_allow_html=True)
    st.write("")
    editor = st.text_input("Seu nome (obrigatório) *", key="confirm_editor",
                           placeholder="Quem está salvando estas alterações")
    st.caption("Seu nome e a data/hora ficam gravados no cadastro.")

    c1, c2 = st.columns(2)
    if c1.button("✅ Confirmar e salvar", type="primary", width="stretch"):
        editor = " ".join(editor.split())
        if len(editor) < 2:
            st.error("Informe seu nome para confirmar.")
            return
        try:
            with st.spinner("Salvando no Google Drive…"):
                cached, err, target = save_or_cache(pending, editor)
        except Exception as exc:
            st.error(f"Não foi possível registrar a alteração no cache: {exc}")
            return
        if cached:
            flash = ("warning", f"Cadastro Nº {numero} salvo no cache. "
                     "Use a opção de enviar alterações pendentes na tela principal quando a conexão voltar.")
        else:
            flash = ("info", f"✅ Cadastro Nº {numero} salvo.")

        _overlay_saved_change(pending["operation"], target)
        st.session_state.pop("pending", None)
        st.session_state.flash = flash
        st.rerun()
    if c2.button("↩️ Voltar", width="stretch"):
        st.session_state.open_dialog = "edit_back"
        st.rerun()


# ──────────────────────────────────────────────────────────────
#  BUSCA
# ──────────────────────────────────────────────────────────────
_ACCENT_CLASS: dict[str, str] = {}
for _base, _variants in [("a", "aàáâãäå"), ("e", "eèéêë"), ("i", "iìíîï"), ("o", "oòóôõö"),
                         ("u", "uùúûü"), ("c", "cç"), ("n", "nñ")]:
    _cls = f"[{_variants}]"
    for _ch in _variants:
        _ACCENT_CLASS[_ch] = _cls


def _normalize(s: str) -> str:
    return "".join(c for c in _ud.normalize("NFD", s) if _ud.category(c) != "Mn").lower()


def _strip_cpf(s: str) -> str:
    return s.replace(".", "").replace("-", "")


def search_numero_exact(query: str) -> pd.DataFrame:
    """Linha cujo NUMERO é exatamente o digitado (ignora zeros à esquerda)."""
    q = query.strip()
    if not q or df is None:
        return pd.DataFrame()
    if q.isdigit():
        q = str(int(q))
    return df[df["NUMERO"].str.strip() == q].copy()


# ──────────────────────────────────────────────────────────────
#  CARD DO REGISTRO
# ──────────────────────────────────────────────────────────────
def cell(row: pd.Series, col: str) -> str:
    v = str(row.get(col, "")).strip()
    return "" if v in ("nan", "None") else v


def _query_to_pattern(query: str) -> str:
    return "".join(_ACCENT_CLASS.get(ch, _re.escape(ch)) for ch in query.lower())


def highlight(text: str, query: str) -> str:
    """Escapa HTML e marca (sem diferenciar acentos) as ocorrências da busca."""
    if not text:
        return ""
    query = (query or "").strip()
    if not query:
        return _html.escape(text)
    try:
        pat = _re.compile(_query_to_pattern(query), _re.IGNORECASE)
    except _re.error:
        return _html.escape(text)
    out, last = [], 0
    for m in pat.finditer(text):
        if m.end() == m.start():
            continue
        out.append(_html.escape(text[last:m.start()]))
        out.append(f'<mark class="hl">{_html.escape(m.group(0))}</mark>')
        last = m.end()
    out.append(_html.escape(text[last:]))
    return "".join(out)


def _row(label: str, value_html: str) -> str:
    return (f'<div class="info-row"><span class="info-label">{label}</span>'
            f'<span class="info-value">{value_html or "<span class=info-empty>—</span>"}</span></div>')


def render_record(row: pd.Series, numero_query: str = "") -> None:
    q_num = numero_query.strip()
    nome, numero, cpf = cell(row, "NOME"), cell(row, "NUMERO"), cell(row, "CPF")
    tipo, cid, tel = cell(row, "TIPO"), cell(row, "CID 2026"), cell(row, "TELEFONE")
    benef = cell(row, "TIPO DE BENEFICIÁRIO [TIPO]")
    cestas = cell(row, COL_CESTAS)

    endereco = ", ".join(p for p in (cell(row, "RUA PRINCIPAL"), cell(row, "Nº CASA")) if p)
    viela = cell(row, "VIELA")
    if viela:
        endereco = f"{endereco} — viela {viela}" if endereco else f"viela {viela}"

    cestas_up = cestas.upper()
    cestas_cls = ("b-need2" if cestas_up == "PRECISA MUITO"
                  else "b-noneed" if cestas_up in ("NÃO PRECISA", "NÃO")
                  else "b-need")
    cestas_html = f'<span class="badge {cestas_cls}">{_html.escape(cestas)}</span>' if cestas else ""
    tipo_html = f'<span class="badge b-tipo">{_html.escape(tipo)}</span>' if tipo else ""

    numero_hl = highlight(numero, q_num)
    h = f"""
    <div class="card">
      <div class="card-name">{_html.escape(nome) or "—"}</div>
      <div class="card-num">Nº {numero_hl} &nbsp;·&nbsp; CPF {_html.escape(cpf) if cpf else "—"}</div>
      <div class="info-grid">
        {_row("Tipo", tipo_html)}
        {_row("CID 2026", _html.escape(cid))}
        {_row("Telefone", _html.escape(tel))}
        {_row("Endereço", _html.escape(endereco))}
        {_row("Precisa cesta", cestas_html)}
        {_row("Benefíc.", _html.escape(benef))}
      </div>
    """
    por, em = cell(row, AUDIT_NAME_COL), cell(row, AUDIT_TS_COL)
    if por or em:
        h += f'<div class="card-audit">Última edição: {_html.escape(por)} · {_html.escape(em)}</div>'
    h += "</div>"
    st.markdown(h, unsafe_allow_html=True)

    # Único botão: sempre disponível
    if st.button("✏️ Editar Cadastro", key=f"edit_{row.name}", width="stretch"):
        edit_dialog(numero, row.to_dict(), nonce=uuid.uuid4().hex[:8])


def record_from_edit_link() -> pd.Series | None:
    """Consome ?id=NUMERO uma vez por link na sessão, sem reabrir o popup nos reruns."""
    ids = st.query_params.get_all("id")
    if not ids:
        st.session_state.pop("_edit_link_handled", None)
        return None
    signature = tuple(ids)
    if st.session_state.get("_edit_link_handled") == signature:
        return None
    # O encadeamento de uma edição em andamento tem prioridade sobre um novo link.
    if (st.session_state.get("open_dialog") in ("confirm", "edit_back") and
            st.session_state.get("pending")):
        return None
    st.session_state._edit_link_handled = signature
    if (len(ids) != 1 or len(ids[0]) > 100 or
            not _re.fullmatch(r"[0-9]+(?:-[A-Z]+)?", ids[0].strip())):
        st.warning("Link de edição inválido. Informe um único número no parâmetro id.")
        return None
    numero = ids[0].strip()
    # Executado antes de criar o widget; depois disso a busca continua funcionando normalmente.
    st.session_state.q_num = numero
    results = search_numero_exact(numero)
    if results.empty:
        st.warning("O cadastro indicado no link não foi encontrado.")
        return None
    if len(results) != 1:
        st.warning("Há mais de um cadastro com o número indicado no link. "
                   "Escolha o cadastro abaixo para editar.")
        return None
    return results.iloc[0]


# ──────────────────────────────────────────────────────────────
#  UI
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
<div class="ps-header">
  <h1>📋 Cadastro de Visitas</h1>
  <p>Pastoral Social · Jardim Colombo</p>
</div>
""",
    unsafe_allow_html=True,
)

_flash = st.session_state.pop("flash", None)
if _flash:
    if _flash[0] == "warning":
        st.warning(_flash[1], icon="⚠️")
    else:
        st.toast(_flash[1])

if not _gdrive_write_enabled():
    st.warning(
        "⚠️ Salvamento no Drive desativado. Configure `GDRIVE_FILE_ID` e "
        "`GDRIVE_SA_CREDENTIALS` nos Secrets para gravar na planilha.",
        icon="🔒",
    )

if _cache_error:
    st.error(f"Não foi possível ler as alterações pendentes do cache: {_cache_error}")
elif _cached_changes:
    _n_pending = len(_cached_changes)
    _pending_label = "alteração" if _n_pending == 1 else "alterações"
    st.warning(f"{_n_pending} {_pending_label} em cache aguardando envio ao Google Drive.", icon="⚠️")
    if st.button("☁️ Enviar alterações pendentes", width="stretch", key="sync_cached_changes"):
        with st.spinner("Enviando todas as alterações pendentes ao Google Drive…"):
            _sync_error, _sent, _targets = sync_cached_changes()
        if _sync_error:
            st.error(f"Não foi possível enviar: {_sync_error}. As alterações continuam no cache.")
        else:
            for _op in _sent:
                _overlay_saved_change(_op, _targets[_op["id"]])
            _sent_label = "alteração enviada" if len(_sent) == 1 else "alterações enviadas"
            st.session_state.flash = ("info", f"✅ {len(_sent)} {_sent_label} ao Google Drive.")
            st.rerun()

if load_error or df is None:
    st.error("❌ Não foi possível carregar o arquivo.")
    st.markdown(
        """
**Possíveis causas:**
- A URL do arquivo ainda não foi configurada (`GITHUB_RAW_URL`).
- O arquivo não existe / é privado.
- Problema de rede.
"""
    )
    if load_error:
        with st.expander("Detalhes do erro"):
            st.code(load_error)
    st.stop()

_edit_link_row = record_from_edit_link()

query_num = st.text_input("Buscar por Número", placeholder="🔢  Número do registro…",
                          label_visibility="collapsed", key="q_num")


def show_results(results: pd.DataFrame, num_query: str) -> None:
    if results.empty:
        st.markdown(
            """
<div class="empty-state">
  <div class="icon">🔎</div>
  <p>Nenhum registro encontrado.<br>Confira o número e tente de novo.</p>
</div>
""",
            unsafe_allow_html=True,
        )
        return

    count = len(results)
    st.markdown(
        f'<div class="result-count">{count} registro{"s" if count > 1 else ""} '
        f'encontrado{"s" if count > 1 else ""}</div>',
        unsafe_allow_html=True,
    )
    for _, row in results.iterrows():
        render_record(row, numero_query=num_query)


if query_num.strip():
    show_results(search_numero_exact(query_num), query_num)
else:
    st.markdown(
        f"""
<div class="empty-state">
  <div class="icon">🔍</div>
  <p>Digite o número de registro conforme o
  <a href="https://www.google.com/maps/d/viewer?mid=10OImJnmvNDhH2FrCKoVIjK0DUt0XuVE"
     target="_blank" rel="noopener noreferrer">Google Maps</a> ou o
  <a href="?mapa=offline" target="_blank" rel="noopener noreferrer">OpenStreetMap (offline)</a></p>
  <p style="margin-top:1.5rem;font-size:0.8rem;color:#bdbdbd">{len(df)} cadastros</p>
</div>
""",
        unsafe_allow_html=True,
    )

# ──────────────────────────────────────────────────────────────
#  ENCADEAMENTO DOS POPUPS (editar → confirmar → [voltar])
#  O Streamlit não permite popup dentro de popup; por isso o primeiro fecha
#  (st.rerun) e o segundo abre logo depois. O flag é consumido (pop) ao abrir,
#  então fechar o popup no "X" não o reabre na próxima interação.
# ──────────────────────────────────────────────────────────────
_next = st.session_state.pop("open_dialog", None)
_pending = st.session_state.get("pending")
if _next == "confirm" and _pending:
    confirm_dialog(_pending)
elif _next == "edit_back" and _pending:
    _row = _row_dict(_pending["numero"])
    if _row:
        edit_dialog(_pending["numero"], _row, nonce=uuid.uuid4().hex[:8], draft=_pending["widgets"])
elif _edit_link_row is not None and not st.session_state.get(f"edit_{_edit_link_row.name}"):
    # O botão manual pode já ter aberto o diálogo nesta execução.
    edit_dialog(cell(_edit_link_row, "NUMERO"), _edit_link_row.to_dict(), nonce=uuid.uuid4().hex[:8])
