import base64
import os
import uuid
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

try:
    import plotly.express as px
except Exception:
    px = None

try:
    import requests
except Exception:
    requests = None

DATA_FILE = "Dados.xlsx"
SHEET = "respostas"
TZ = "America/Sao_Paulo"
FOOTER_LOGOS_FILE = os.path.join("assets", "logos-rodape.svg")
STYLE_FILE = "styles.css"

# Quando o app roda no Streamlit Community Cloud, o disco é apagado a cada
# reinício/redeploy — por isso as respostas são gravadas em uma base do
# Airtable sempre que houver credenciais configuradas em st.secrets. Sem
# essas credenciais (ex.: rodando localmente sem configurar nada), o app
# volta a gravar no arquivo Dados.xlsx local, como antes.
AIRTABLE_TABLE = "respostas"

# ---- Metas (ajuste livre) ----
TARGET_NOTA_GERAL = 8.0  # média da nota geral 0-10
TARGET_PCT_POSITIVAS = 70.0  # % mínimo desejado de avaliações positivas (nota 8-10)
TARGET_PCT_NEGATIVAS_MAX = 15.0  # % máximo aceitável de avaliações negativas (nota 0-4)

LIKERT = {
    1: "Muito insatisfeito(a)",
    2: "Insatisfeito(a)",
    3: "Neutro",
    4: "Satisfeito(a)",
    5: "Muito satisfeito(a)",
}
LIKERT_LEGENDA = "  ·  ".join(f"{n} = {LIKERT[n]}" for n in range(1, 6))

CANAIS_EXTERNO = [
    "Balcão Virtual",
    "E-mail",
    "Telefone",
    "Presencial",
    "PJe",
    "WhatsApp (institucional)",
    "Outro",
]
CANAIS_INTERNO = [
    "E-mail institucional",
    "Sistema interno (PJe/SEI)",
    "Intranet",
    "Reunião presencial",
    "Telefone/ramal",
    "Outro",
]

UNIDADES = [
    "Fortaleza",
    "Juazeiro do Norte",
    "Outra (informar)",
]

PERFIL_EXTERNO = "Usuário(a) externo(a) / Parte"
PERFIL_ADVOGADO = "Advogado(a)"
PERFIL_SERVIDOR = "Servidor(a)"
PERFIL_MAGISTRADO = "Magistrado(a)"

TIPOS_USUARIO = [PERFIL_EXTERNO, PERFIL_ADVOGADO, PERFIL_SERVIDOR, PERFIL_MAGISTRADO]

# Respostas antigas (registradas antes deste perfil de perguntas existir) usam
# nomes diferentes para o mesmo perfil — mapeamos para o nome atual para que o
# Painel consiga cruzar os dados antigos com os indicadores específicos de
# cada perfil. "Outro" fica de fora por ser um valor genérico demais para
# associar com segurança a um perfil específico.
LEGACY_PERFIL_MAP = {
    "Jurisdicionado": PERFIL_EXTERNO,
    "Advogado": PERFIL_ADVOGADO,
    "Servidor/Colaborador": PERFIL_SERVIDOR,
}

PERFIL_INFO = {
    PERFIL_EXTERNO: {"icone": "👤", "descricao": "Sou parte, jurisdicionado(a) ou público em geral", "cor": "#2563eb", "classe": "perfil-externo"},
    PERFIL_ADVOGADO: {"icone": "⚖️", "descricao": "Atuo representando partes em processos", "cor": "#b45309", "classe": "perfil-advogado"},
    PERFIL_SERVIDOR: {"icone": "🏛️", "descricao": "Trabalho na JFCE (servidor/colaborador)", "cor": "#1f6f5c", "classe": "perfil-servidor"},
    PERFIL_MAGISTRADO: {"icone": "🧑‍⚖️", "descricao": "Sou juiz(a) federal ou desembargador(a)", "cor": "#7c3aed", "classe": "perfil-magistrado"},
}

AREAS_ATUACAO = [
    "Cível",
    "Tributário",
    "Previdenciário",
    "Trabalhista",
    "Criminal",
    "Execução fiscal",
    "Outro",
]

LOTACOES = [
    "Gabinete",
    "Secretaria de Vara",
    "Setor administrativo",
    "Tecnologia da Informação",
    "Atendimento ao público",
    "Outro",
]

FAIXAS_IDADE = ["18-30", "31-40", "41-50", "51-60", "61+"]
GENEROS = ["Feminino", "Masculino", "Outro", "Prefiro não informar"]

# Colunas comuns a todos os perfis, perguntadas na "Avaliação geral" (mantidas
# só 2 para deixar essa seção com apenas 3 perguntas ao todo — as 2 abaixo +
# a nota geral 0-10). O texto de cada uma foi ampliado para cobrir, num único
# enunciado, aspectos que antes eram perguntas separadas (facilidade_contato,
# tempo_resposta, acessibilidade, cordialidade_respeito) — essas colunas
# continuam no schema por compatibilidade com respostas antigas, mas não são
# mais perguntadas isoladamente.
COMUNS_DIMENSOES = [
    (
        "clareza_informacoes",
        "Clareza e acessibilidade das informações recebidas (linguagem compreensível, "
        "canais fáceis de encontrar e usar)",
    ),
    (
        "resolutividade",
        "Resolutividade e agilidade no atendimento (sua demanda foi resolvida de forma "
        "eficaz, com cordialidade e em tempo razoável)",
    ),
]

# Colunas específicas por perfil (avaliadas na etapa 2)
DIMENSOES_PERFIL = {
    PERFIL_EXTERNO: [
        ("usabilidade_balcao_virtual", "Usabilidade do Balcão Virtual"),
        ("experiencia_audiencia", "Experiência em audiência"),
    ],
    PERFIL_ADVOGADO: [
        ("usabilidade_pje", "Usabilidade do PJe/peticionamento eletrônico"),
        ("cumprimento_prazos", "Cumprimento de prazos processuais pela vara"),
        ("comunicacao_secretaria", "Comunicação com a secretaria/servidores"),
        ("experiencia_audiencia", "Experiência em audiência"),
    ],
    PERFIL_SERVIDOR: [
        ("suporte_ti", "Suporte de TI/sistemas internos"),
        ("comunicacao_interna", "Comunicação interna/gestão"),
        ("condicoes_trabalho", "Condições de trabalho/infraestrutura"),
        ("capacitacao", "Capacitação e desenvolvimento profissional"),
    ],
    PERFIL_MAGISTRADO: [
        ("suporte_assessoria", "Suporte da assessoria/gabinete"),
        ("infraestrutura_tecnologica", "Infraestrutura tecnológica (sistemas, PJe)"),
        ("suporte_administrativo", "Suporte administrativo (secretaria/administração)"),
        ("comunicacao_gestao", "Comunicação com a gestão do tribunal"),
    ],
}

COLUMNS = [
    "timestamp",
    "respondent_id",
    "unidade",
    "tipo_usuario",
    "faixa_idade",
    "genero",
    "canal_contato",
    "area_atuacao",
    "lotacao",
    "ja_usou_balcao_virtual",
    "usabilidade_balcao_virtual",
    "ja_participou_audiencia",
    "experiencia_audiencia",
    "usabilidade_pje",
    "cumprimento_prazos",
    "comunicacao_secretaria",
    "suporte_ti",
    "comunicacao_interna",
    "condicoes_trabalho",
    "capacitacao",
    "suporte_assessoria",
    "infraestrutura_tecnologica",
    "suporte_administrativo",
    "comunicacao_gestao",
    "clareza_informacoes",
    "cordialidade_respeito",
    "facilidade_contato",
    "tempo_resposta",
    "resolutividade",
    "acessibilidade",
    "satisfacao_geral",
    "recomendacao_0_10",
    "comentario_aberto",
]


# ---------------- Dados ----------------
def _airtable_configured() -> bool:
    """True quando há token + base do Airtable configurados em st.secrets
    (ver README para o passo a passo)."""
    if requests is None:
        return False
    try:
        return bool(st.secrets.get("airtable_token")) and bool(st.secrets.get("airtable_base_id"))
    except Exception:
        return False


def _airtable_headers() -> dict:
    return {
        "Authorization": f"Bearer {st.secrets['airtable_token']}",
        "Content-Type": "application/json",
    }


def _airtable_url(suffix: str = "") -> str:
    return f"https://api.airtable.com/v0/{st.secrets['airtable_base_id']}/{AIRTABLE_TABLE}{suffix}"


def _serialize_valor(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v)


def ensure_data_file():
    if _airtable_configured():
        return
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=SHEET, index=False)


@st.cache_data(show_spinner=False)
def _read_excel_cached(path: str, mtime: float) -> pd.DataFrame:
    # mtime faz parte da chave de cache: muda quando o arquivo é regravado
    return pd.read_excel(path, sheet_name=SHEET)


def _airtable_fetch_all() -> list:
    records, offset, params = [], None, {"pageSize": 100}
    while True:
        p = dict(params)
        if offset:
            p["offset"] = offset
        resp = requests.get(_airtable_url(), headers=_airtable_headers(), params=p, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
    return records


@st.cache_data(ttl=20, show_spinner=False)
def _read_airtable_cached(versao: int) -> pd.DataFrame:
    # versao muda a cada gravação (append_row incrementa o contador em
    # session_state), e o ttl=20s cobre o caso de outra pessoa ter
    # respondido a pesquisa em uma sessão diferente.
    registros = _airtable_fetch_all()
    linhas = [r.get("fields", {}) for r in registros]
    return pd.DataFrame(linhas)


def load_data() -> pd.DataFrame:
    ensure_data_file()
    if _airtable_configured():
        versao = st.session_state.get("_airtable_versao", 0)
        df = _read_airtable_cached(versao)
    else:
        df = _read_excel_cached(DATA_FILE, os.path.getmtime(DATA_FILE))
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = pd.NA
    df = df[COLUMNS].copy()
    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df["tipo_usuario"] = df["tipo_usuario"].replace(LEGACY_PERFIL_MAP)
    return df


def append_row(row: dict):
    ensure_data_file()
    if _airtable_configured():
        fields = {c: _serialize_valor(row.get(c)) for c in COLUMNS}
        resp = requests.post(
            _airtable_url(), headers=_airtable_headers(),
            json={"records": [{"fields": fields}]}, timeout=30,
        )
        resp.raise_for_status()
        st.session_state["_airtable_versao"] = st.session_state.get("_airtable_versao", 0) + 1
        _read_airtable_cached.clear()
        return
    df = pd.read_excel(DATA_FILE, sheet_name=SHEET)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = pd.NA
    df = df[COLUMNS]
    df2 = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl", mode="w") as writer:
        df2.to_excel(writer, sheet_name=SHEET, index=False)
    _read_excel_cached.clear()


def distribuicao_satisfacao(series_0_10: pd.Series) -> dict:
    """% de avaliações positivas (8-10), neutras (5-7) e negativas (0-4), e o n considerado.

    Faixas pensadas para uma pesquisa de satisfação (não é uma métrica de NPS/
    recomendação — por isso não usamos os termos "promotor/detrator")."""
    s = pd.to_numeric(series_0_10, errors="coerce").dropna()
    if len(s) == 0:
        return {"positivas": float("nan"), "neutras": float("nan"), "negativas": float("nan"), "n": 0}
    return {
        "positivas": (s >= 8).mean() * 100,
        "neutras": ((s >= 5) & (s <= 7)).mean() * 100,
        "negativas": (s <= 4).mean() * 100,
        "n": int(len(s)),
    }


def period_floor(ts: pd.Series, freq: str) -> pd.Series:
    if freq == "W":
        return ts.dt.to_period("W").dt.start_time
    if freq == "M":
        return ts.dt.to_period("M").dt.start_time
    return ts.dt.floor("D")


@st.cache_data(show_spinner=False)
def _logo_base64(path: str):
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def _asset_data_uri(path: str, mime_type: str) -> str | None:
    encoded = _logo_base64(path)
    return f"data:{mime_type};base64,{encoded}" if encoded else None


# ---------------- Estilo ----------------
def inject_css():
    st.markdown(
        """
        <style>
        /* Fundo geral do app: gradiente suave em vez do branco padrão */
        [data-testid="stAppViewContainer"] {
            background: linear-gradient(160deg, #eef2f9 0%, #e6ecf5 45%, #eef5f2 100%);
        }
        [data-testid="stHeader"] {
            background: transparent;
        }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0b2545 0%, #13315c 100%);
        }
        /* Texto solto (rótulos, legendas, títulos) direto sobre o fundo escuro
           da sidebar fica claro para ler. */
        [data-testid="stSidebar"] * {
            color: #eef2f9 !important;
        }
        [data-testid="stSidebar"] div[role="radiogroup"] > label {
            border-color: rgba(255, 255, 255, 0.35) !important;
            background: rgba(255, 255, 255, 0.06);
        }
        /* As caixas dos filtros (selectbox, período) têm fundo claro por padrão —
           manter o texto escuro *dentro* delas, senão fica claro-sobre-claro e
           ilegível (o "!important" acima venceria e apagaria o texto). */
        [data-testid="stSidebar"] div[data-baseweb="select"],
        [data-testid="stSidebar"] div[data-baseweb="select"] > div,
        [data-testid="stSidebar"] div[data-testid="stDateInput"] input {
            background-color: #ffffff !important;
            border-radius: 8px;
        }
        [data-testid="stSidebar"] div[data-baseweb="select"] *,
        [data-testid="stSidebar"] div[data-testid="stDateInput"] input,
        [data-testid="stSidebar"] div[data-testid="stDateInput"] svg {
            color: #1f2937 !important;
            fill: #1f2937;
        }

        .block-container {padding-top: 4.5rem; padding-left: 2rem; padding-right: 2rem; max-width: 100%;}
        div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) {
            border: 1px solid rgba(49, 51, 63, 0.10);
            border-radius: 16px;
            padding: 1.8rem 1.8rem 1.2rem 1.8rem;
            background: #ffffff;
            box-shadow: 0 10px 28px rgba(11, 37, 69, 0.08);
        }
        /* Painel: todo gráfico e toda tabela viram um "cartão" branco com cantos
           arredondados e sombra leve — mesmo tratamento visual em ambos. */
        div[data-testid="stElementContainer"]:has(> div[data-testid="stFullScreenFrame"] div[data-testid="stPlotlyChart"]),
        div[data-testid="stElementContainer"]:has(> div[data-testid="stDataFrame"]) {
            background: #ffffff;
            border: 1px solid rgba(49, 51, 63, 0.10);
            border-radius: 16px;
            padding: 1rem 1rem 0.4rem 1rem;
            box-shadow: 0 8px 20px rgba(11, 37, 69, 0.06);
            margin-bottom: 0.6rem;
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stFullScreenFrame"] div[data-testid="stPlotlyChart"]) div[data-testid="stPlotlyChart"],
        div[data-testid="stElementContainer"]:has(> div[data-testid="stDataFrame"]) div[data-testid="stDataFrame"] {
            border-radius: 12px;
            overflow: hidden;
        }
        /* Grupos de opções (Likert, sim/não, nota 0-10): tudo numa única linha, encolhendo
           o necessário para caber — pensado para uso em tablet — com mais respiro entre elas.
           Mantém o círculo de seleção do radio, só sem a borda/pílula de fundo ao redor
           de cada opção. */
        div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) div[role="radiogroup"] {
            flex-wrap: nowrap;
            display: flex;
            gap: 3.6rem;
            width: 100%;
            margin-top: 0.2rem;
            margin-bottom: 1.1rem;
        }
        div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) div[role="radiogroup"] > label {
            border: none;
            padding: 0.3rem 0.4rem;
            border-radius: 10px;
            margin-right: 0;
            flex: 1 1 0;
            justify-content: center;
            min-width: 2.7rem;
            white-space: nowrap;
            background: transparent;
        }
        div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) div[role="radiogroup"] > label:has(input:checked) p {
            font-weight: 700;
            color: #1d4ed8;
        }
        div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) div[role="radiogroup"] > label p {
            font-size: 0.92rem;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            margin: 0;
        }
        .likert-legenda {
            color: #666;
            font-size: 0.8rem;
            margin: -0.3rem 0 0.9rem 0;
        }
        /* A pergunta de 0 a 10 tem 11 opções — precisa de um espaçamento menor
           que o das demais (que têm só 2 ou 5 opções) para caber tudo numa linha só. */
        .st-key-q_nps div[role="radiogroup"] {
            gap: 0.6rem;
        }
        .st-key-q_nps div[role="radiogroup"] > label {
            padding: 0.5rem 0.3rem;
            min-width: 2.4rem;
        }
        @media (max-width: 900px) {
            .block-container {padding-left: 0.7rem; padding-right: 0.7rem; padding-top: 3.6rem;}
            div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) div[role="radiogroup"] {gap: 1.8rem;}
            div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) div[role="radiogroup"] > label {padding: 0.4rem 0.35rem; min-width: 2.3rem;}
            div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) div[role="radiogroup"] > label p {font-size: 0.82rem;}
            div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] > div[data-testid="stMarkdown"] .questionario-marker) {padding: 1.1rem 0.9rem 0.6rem 0.9rem;}
            .st-key-q_nps div[role="radiogroup"] {gap: 0.35rem;}
            .st-key-q_nps div[role="radiogroup"] > label {padding: 0.35rem 0.2rem; min-width: 2rem;}
        }
        .jfce-header {
            background: linear-gradient(90deg, #0b2545 0%, #13315c 55%, #1f6f5c 100%);
            padding: 0.9rem 1.6rem;
            border-radius: 14px;
            margin-bottom: 1.1rem;
            display: flex;
            align-items: center;
            gap: 1rem;
            box-sizing: border-box;
        }
        .jfce-header img {
            height: 56px;
            width: auto;
            max-width: 90px;
            object-fit: contain;
            display: block;
            flex-shrink: 0;
        }
        .jfce-header .textos {display: flex; flex-direction: column; justify-content: center; gap: 0.15rem;}
        .jfce-header .titulo {color: #fff; font-size: 1.1rem; font-weight: 600; line-height: 1.3; margin: 0;}
        .jfce-header .subtitulo {color: #d7e3f0; font-size: 0.85rem; line-height: 1.3; margin: 0;}
        .step-caption {color: #4a4a4a; font-size: 0.9rem; margin-bottom: 0.4rem;}
        .intro-titulo {font-size: 1.6rem; font-weight: 700; margin-bottom: 0.2rem;}
        .intro-sub {color: #555; font-size: 1rem; margin-bottom: 1.4rem;}

        /* Cards de perfil = o próprio botão (o card inteiro é clicável, ótimo para tablet).
           O marcador (.perfil-marker) fica num elemento acima do botão; usamos :has() para
           "pintar" o botão do elemento seguinte com base na classe do marcador. */
        .perfil-marker {display: none;}
        .questionario-marker {display: none;}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] {
            height: 100%;
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button {
            width: 100%;
            min-height: 30vh;
            height: 100%;
            border: 3px solid #ccc;
            border-radius: 28px;
            background: linear-gradient(160deg, #ffffff 0%, #f5f7fa 100%);
            box-shadow: 0 8px 22px rgba(11, 37, 69, 0.12);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 0.3rem;
            padding: 1.5rem 1rem;
            transition: transform .15s ease, box-shadow .15s ease;
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button:hover,
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button:focus {
            transform: translateY(-5px);
            box-shadow: 0 14px 30px rgba(11, 37, 69, 0.2);
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button p:nth-child(1) {
            font-size: 6rem;
            line-height: 1;
            margin-bottom: 0.4rem;
            width: 100%;
            text-align: center;
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button p:nth-child(2) {
            font-size: 1.5rem;
            font-weight: 700;
            margin: 0.2rem 0;
            text-align: center;
            width: 100%;
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button p:nth-child(3) {
            font-size: 1.05rem;
            font-weight: 400;
            color: #555;
            width: 100%;
            text-align: center;
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button div[data-testid="stMarkdownContainer"] {
            width: 100%;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-externo) + div[data-testid="stElementContainer"] button {border-color: #2563eb; background: linear-gradient(160deg, #ffffff 0%, #2563eb1c 100%);}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-externo) + div[data-testid="stElementContainer"] button p:nth-child(2) {color: #2563eb;}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-advogado) + div[data-testid="stElementContainer"] button {border-color: #b45309; background: linear-gradient(160deg, #ffffff 0%, #b453091c 100%);}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-advogado) + div[data-testid="stElementContainer"] button p:nth-child(2) {color: #b45309;}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-servidor) + div[data-testid="stElementContainer"] button {border-color: #1f6f5c; background: linear-gradient(160deg, #ffffff 0%, #1f6f5c1c 100%);}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-servidor) + div[data-testid="stElementContainer"] button p:nth-child(2) {color: #1f6f5c;}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-magistrado) + div[data-testid="stElementContainer"] button {border-color: #7c3aed; background: linear-gradient(160deg, #ffffff 0%, #7c3aed1c 100%);}
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker.perfil-magistrado) + div[data-testid="stElementContainer"] button p:nth-child(2) {color: #7c3aed;}

        @media (max-width: 900px) {
            div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button {min-height: 26vh; padding: 1.1rem 0.7rem;}
            div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button p:nth-child(1) {font-size: 4.2rem;}
            div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button p:nth-child(2) {font-size: 1.15rem;}
            div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdown"] .perfil-marker) + div[data-testid="stElementContainer"] button p:nth-child(3) {font-size: 0.85rem;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    if os.path.exists(STYLE_FILE):
        with open(STYLE_FILE, "r", encoding="utf-8") as css_file:
            st.markdown(f"<style>{css_file.read()}</style>", unsafe_allow_html=True)


def render_header(subtitle: str):
    st.markdown(
        f"""
        <div class="jfce-header">
            <div class="jfce-header__content">
                <div class="jfce-header__eyebrow">Justiça Federal no Ceará</div>
                <div class="jfce-header__title">Pesquisa de Satisfação</div>
                <div class="jfce-header__subtitle">{subtitle}</div>
            </div>
            <div class="jfce-header__badge">Pesquisa anônima</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_brand():
    st.sidebar.markdown(
        """
        <div class="sidebar-brand">
            <div class="sidebar-brand__eyebrow">JFCE</div>
            <div class="sidebar-brand__title">Pesquisa de<br>Satisfação</div>
            <div class="sidebar-brand__subtitle">Escuta ativa para melhorar nossos serviços</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_footer():
    footer_uri = _asset_data_uri(FOOTER_LOGOS_FILE, "image/svg+xml")
    if footer_uri:
        st.sidebar.markdown(
            f'<div class="sidebar-footer"><img src="{footer_uri}" alt="Logos institucionais" /></div>',
            unsafe_allow_html=True,
        )


def style_plotly(fig, height: int | None = None):
    fig.update_layout(
        font={"family": "Montserrat, Inter, Segoe UI, sans-serif", "color": "#3f4347", "size": 12},
        title={"font": {"color": "#25282b", "size": 16}, "x": 0.02, "xanchor": "left"},
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        colorway=["#006f9f", "#36a8c7", "#21855b", "#e59a3a", "#c33c3c"],
        margin={"l": 28, "r": 20, "t": 58, "b": 36},
        hoverlabel={"font": {"family": "Montserrat, Inter, Segoe UI, sans-serif"}},
        legend={"title": None, "orientation": "h", "y": -0.18},
    )
    if height:
        fig.update_layout(height=height)
    fig.update_xaxes(showgrid=False, linecolor="#e8eaec", tickfont={"color": "#6d7378"})
    fig.update_yaxes(gridcolor="#eef0f1", zeroline=False, tickfont={"color": "#6d7378"})
    return fig


def likert_q(label: str, key: str, help_txt: str = "") -> int:
    return st.radio(
        label,
        options=[1, 2, 3, 4, 5],
        format_func=lambda x: f"{x}",
        key=key,
        horizontal=True,
        help=help_txt or LIKERT_LEGENDA,
    )


def reset_wizard():
    for k in list(st.session_state.keys()):
        if k.startswith("q_") or k in ("step", "respostas", "enviado"):
            del st.session_state[k]
    st.session_state.step = 0
    st.session_state.respostas = {}
    st.session_state.enviado = False


def escolher_perfil(perfil: str):
    st.session_state.respostas = {"tipo_usuario": perfil}
    st.session_state.step = 1


# ---------------- App ----------------
st.set_page_config(
    page_title="Pesquisa de Satisfação — JFCE",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="auto",
)
inject_css()

if "step" not in st.session_state:
    st.session_state.step = 0
if "respostas" not in st.session_state:
    st.session_state.respostas = {}
if "enviado" not in st.session_state:
    st.session_state.enviado = False

render_sidebar_brand()
page = st.sidebar.radio("Navegação", ["Responder pesquisa", "Painel (análises)"])
st.sidebar.markdown(
    '<div class="sidebar-privacy">As respostas são registradas de forma anonimizada, sem nome, e-mail ou CPF.</div>',
    unsafe_allow_html=True,
)

# =========================================================
# PÁGINA: RESPONDER
# =========================================================
if page == "Responder pesquisa":
    render_sidebar_footer()
    render_header("Sua opinião ajuda a melhorar nossos serviços — leva menos de 2 minutos.")

    if st.session_state.enviado:
        st.success("✅ Resposta registrada com sucesso. Muito obrigado(a) pela contribuição!")
        st.button("Responder novamente", on_click=reset_wizard)
        st.stop()

    step = st.session_state.step

    # ---------- ETAPA 0: tela inicial — escolha do perfil ----------
    if step == 0:
        st.markdown('<div class="intro-titulo">Para começar, selecione o seu perfil</div>', unsafe_allow_html=True)
        st.markdown('<div class="intro-sub">As perguntas seguintes serão adaptadas ao perfil escolhido.</div>', unsafe_allow_html=True)

        # Grade 2x2: cards grandes e o card inteiro é clicável (bom para uso em tablet)
        linhas = [TIPOS_USUARIO[0:2], TIPOS_USUARIO[2:4]]
        for linha in linhas:
            cols = st.columns(2)
            for col, perfil in zip(cols, linha):
                info = PERFIL_INFO[perfil]
                with col:
                    st.markdown(f'<div class="perfil-marker {info["classe"]}"></div>', unsafe_allow_html=True)
                    st.button(
                        f"{info['icone']}\n\n**{perfil}**\n\n{info['descricao']}",
                        key=f"btn_perfil_{perfil}",
                        use_container_width=True,
                        on_click=escolher_perfil,
                        args=(perfil,),
                    )

    # ---------- QUESTIONÁRIO COMPLETO (uma única página) ----------
    elif step == 1:
        perfil = st.session_state.respostas.get("tipo_usuario", PERFIL_EXTERNO)
        info = PERFIL_INFO[perfil]

        colp, colb = st.columns([5, 1])
        with colp:
            st.markdown(f"#### Perfil selecionado: {perfil}")
        with colb:
            if st.button("Trocar perfil", key="btn_trocar_perfil"):
                reset_wizard()
                st.rerun()

        # Importante: isto é um st.container (não um st.form). Dentro de um st.form,
        # o Streamlit só reroda o script quando o botão de envio é clicado — então
        # perguntas condicionais (ex.: "usabilidade do Balcão Virtual" só aparecer
        # depois de responder "Sim") nunca apareceriam de verdade, só depois de
        # enviar. Com um container comum, cada resposta atualiza a tela na hora.
        with st.container(border=True):
            # marcador invisível: o Streamlit não deixa dar uma classe própria ao
            # container, então usamos este marcador + CSS :has() para aplicar o
            # cartão branco só neste bloco (ver inject_css).
            st.markdown('<div class="questionario-marker"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-title">Dados básicos</div>', unsafe_allow_html=True)
            tem_campo_extra = perfil in (PERFIL_ADVOGADO, PERFIL_SERVIDOR)
            campos = st.columns(4 if tem_campo_extra else 3)

            with campos[0]:
                unidade = st.selectbox("Unidade/Vara", UNIDADES, key="q_unidade")
                if unidade == "Outra (informar)":
                    unidade = st.text_input("Informe a unidade", key="q_unidade_outra").strip() or "Outra"

            idx = 1
            if perfil == PERFIL_ADVOGADO:
                with campos[idx]:
                    area_atuacao = st.selectbox("Área de atuação (opcional)", ["(não informar)"] + AREAS_ATUACAO, key="q_area")
                idx += 1
            elif perfil == PERFIL_SERVIDOR:
                with campos[idx]:
                    area_atuacao = st.selectbox("Lotação (opcional)", ["(não informar)"] + LOTACOES, key="q_lotacao")
                idx += 1
            else:
                area_atuacao = None

            with campos[idx]:
                faixa_idade = st.selectbox("Faixa etária (opcional)", ["(não informar)"] + FAIXAS_IDADE, key="q_faixa")
            with campos[idx + 1]:
                genero = st.selectbox("Gênero (opcional)", ["(não informar)"] + GENEROS, key="q_genero")

            st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
            st.markdown(
                f'<div class="section-title section-title--profile">Perguntas específicas para {perfil}</div>',
                unsafe_allow_html=True,
            )
            st.markdown(f'<div class="likert-legenda">Escala: {LIKERT_LEGENDA}</div>', unsafe_allow_html=True)

            dados = {}

            if perfil == PERFIL_EXTERNO:
                dados["canal_contato"] = st.selectbox("Canal mais usado para contato", CANAIS_EXTERNO, key="q_canal")
                if dados["canal_contato"] == "Balcão Virtual":
                    # já disse que o Balcão Virtual é o canal que mais usa — não faz
                    # sentido perguntar se já usou; vai direto para a avaliação dele.
                    dados["ja_usou_balcao_virtual"] = "Sim"
                    dados["usabilidade_balcao_virtual"] = likert_q("Usabilidade do Balcão Virtual", "q_usab_balcao")
                else:
                    dados["ja_usou_balcao_virtual"] = st.radio(
                        "Mesmo não sendo seu canal principal, já usou o Balcão Virtual alguma vez?",
                        ["Não", "Sim"], key="q_balcao", horizontal=True,
                    )
                    if dados["ja_usou_balcao_virtual"] == "Sim":
                        dados["usabilidade_balcao_virtual"] = likert_q("Usabilidade do Balcão Virtual", "q_usab_balcao")
                dados["ja_participou_audiencia"] = st.radio("Participou de audiência recentemente?", ["Não", "Sim"], key="q_aud_part", horizontal=True)
                if dados["ja_participou_audiencia"] == "Sim":
                    dados["experiencia_audiencia"] = likert_q("Como foi sua experiência na audiência?", "q_exp_aud")

            elif perfil == PERFIL_ADVOGADO:
                dados["canal_contato"] = st.selectbox("Canal mais usado para contato", CANAIS_EXTERNO, key="q_canal")
                dados["usabilidade_pje"] = likert_q("Usabilidade do PJe/peticionamento eletrônico", "q_pje")
                dados["cumprimento_prazos"] = likert_q("Cumprimento de prazos processuais pela vara", "q_prazos")
                dados["comunicacao_secretaria"] = likert_q("Comunicação com a secretaria/servidores", "q_comunic_sec")
                dados["ja_participou_audiencia"] = st.radio("Participou de audiência recentemente?", ["Não", "Sim"], key="q_aud_part", horizontal=True)
                if dados["ja_participou_audiencia"] == "Sim":
                    dados["experiencia_audiencia"] = likert_q("Como foi sua experiência na audiência?", "q_exp_aud")

            elif perfil == PERFIL_SERVIDOR:
                dados["canal_contato"] = st.selectbox("Canal de comunicação interna mais usado", CANAIS_INTERNO, key="q_canal")
                dados["suporte_ti"] = likert_q("Suporte de TI/sistemas internos", "q_suporte_ti")
                dados["comunicacao_interna"] = likert_q("Comunicação interna/gestão", "q_comunic_int")
                dados["condicoes_trabalho"] = likert_q("Condições de trabalho/infraestrutura", "q_condicoes")
                dados["capacitacao"] = likert_q("Capacitação e desenvolvimento profissional", "q_capacitacao")

            else:  # Magistrado
                dados["canal_contato"] = st.selectbox("Canal de comunicação mais usado", CANAIS_INTERNO, key="q_canal")
                dados["suporte_assessoria"] = likert_q("Suporte da assessoria/gabinete", "q_assessoria")
                dados["infraestrutura_tecnologica"] = likert_q("Infraestrutura tecnológica (sistemas, PJe)", "q_infra")
                dados["suporte_administrativo"] = likert_q("Suporte administrativo (secretaria/administração)", "q_suporte_adm")
                dados["comunicacao_gestao"] = likert_q("Comunicação com a gestão do tribunal", "q_comunic_gestao")

            st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-title section-title--rating">Avaliação geral</div>', unsafe_allow_html=True)

            comuns = {}
            for col, label in COMUNS_DIMENSOES:
                comuns[col] = likert_q(label, f"q_{col}")

            st.markdown("De 0 a 10, qual sua avaliação geral da JFCE e a chance de você recomendá-la a outras pessoas?")
            comuns["recomendacao_0_10"] = st.radio(
                "De 0 a 10, qual sua avaliação geral da JFCE e a chance de você recomendá-la a outras pessoas?",
                options=list(range(0, 11)),
                index=8,
                key="q_nps",
                horizontal=True,
                label_visibility="collapsed",
                help="0 = péssima / nada provável  ·  10 = ótima / com certeza recomendaria",
            )
            st.markdown('<div class="likert-legenda">0 = péssima, não recomendaria  ·  10 = ótima, com certeza recomendaria</div>', unsafe_allow_html=True)

            comuns["comentario_aberto"] = st.text_area(
                "Comentário/sugestão (opcional)", placeholder="Escreva aqui (máx. 500 caracteres)...", max_chars=500, key="q_comentario"
            )

            enviar = st.button("Enviar resposta", type="primary", use_container_width=True, key="btn_enviar_resposta")

        if enviar:
            st.session_state.respostas.update(
                {
                    "unidade": unidade,
                    "faixa_idade": "" if faixa_idade == "(não informar)" else faixa_idade,
                    "genero": "" if genero == "(não informar)" else genero,
                    "area_atuacao": area_atuacao if perfil == PERFIL_ADVOGADO and area_atuacao != "(não informar)" else "",
                    "lotacao": area_atuacao if perfil == PERFIL_SERVIDOR and area_atuacao != "(não informar)" else "",
                }
            )
            st.session_state.respostas.update(dados)
            st.session_state.respostas.update(comuns)
            row = {c: None for c in COLUMNS}
            row.update(
                {
                    "timestamp": datetime.now(),
                    "respondent_id": str(uuid.uuid4()),
                    **st.session_state.respostas,
                }
            )
            row["comentario_aberto"] = (row.get("comentario_aberto") or "").strip()
            append_row(row)
            st.session_state.enviado = True
            st.rerun()

# =========================================================
# PÁGINA: PAINEL
# =========================================================
else:
    render_header("Painel de acompanhamento das respostas.")
    st.markdown('<div class="dashboard-title">Painel de resultados</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="dashboard-subtitle">Acompanhe a percepção dos públicos e identifique oportunidades de melhoria.</div>',
        unsafe_allow_html=True,
    )

    df = load_data()

    if df.empty:
        render_sidebar_footer()
        st.info("Ainda não há respostas registradas.")
        st.stop()

    st.sidebar.markdown('<div class="sidebar-filter-title">Filtros</div>', unsafe_allow_html=True)
    unidades = ["(todas)"] + sorted([x for x in df["unidade"].dropna().unique().tolist() if str(x).strip() != ""])
    tipos = ["(todos)"] + sorted([x for x in df["tipo_usuario"].dropna().unique().tolist() if str(x).strip() != ""])
    freq = st.sidebar.selectbox("Periodicidade", ["Diário", "Semanal", "Mensal"])
    freq_map = {"Diário": "D", "Semanal": "W", "Mensal": "M"}

    un_sel = st.sidebar.selectbox("Unidade", unidades)
    tp_sel = st.sidebar.selectbox("Tipo de usuário", tipos)

    dmin = df["timestamp"].min()
    dmax = df["timestamp"].max()
    date_range = st.sidebar.date_input("Período", value=(dmin.date(), dmax.date()))
    render_sidebar_footer()
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
    else:
        start, end = dmin.date(), dmax.date()

    f = df.copy()
    f = f[(f["timestamp"].dt.date >= start) & (f["timestamp"].dt.date <= end)]
    if un_sel != "(todas)":
        f = f[f["unidade"] == un_sel]
    if tp_sel != "(todos)":
        f = f[f["tipo_usuario"] == tp_sel]

    if f.empty:
        st.warning("Sem dados para os filtros selecionados.")
        st.stop()

    f = f.sort_values("timestamp")
    f["periodo"] = period_floor(f["timestamp"], freq_map[freq])

    cutoff = max(f["timestamp"].max() - timedelta(days=30), f["timestamp"].min())
    last30 = f[f["timestamp"] >= cutoff]

    media_last = pd.to_numeric(last30["recomendacao_0_10"], errors="coerce").mean()
    dist_last = distribuicao_satisfacao(last30["recomendacao_0_10"])
    n_last = len(last30)

    # ---------- KPIs ----------
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Respostas (período)", len(f))
    c2.metric(
        "Nota média geral (0–10, últ. 30d)",
        f"{media_last:.1f}" if pd.notna(media_last) else "—",
        delta=f"Meta {TARGET_NOTA_GERAL}",
    )
    c3.metric(
        "Avaliações positivas (últ. 30d)",
        f"{dist_last['positivas']:.0f}%" if dist_last["n"] else "—",
        delta=f"Meta {TARGET_PCT_POSITIVAS:.0f}%",
    )
    c4.metric(
        "Avaliações negativas (últ. 30d)",
        f"{dist_last['negativas']:.0f}%" if dist_last["n"] else "—",
        delta=f"Meta ≤ {TARGET_PCT_NEGATIVAS_MAX:.0f}%",
        delta_color="inverse",
    )

    if pd.notna(media_last) and media_last < TARGET_NOTA_GERAL:
        st.warning(f"⚠️ Nota média geral nos últimos 30 dias abaixo da meta ({media_last:.1f} < {TARGET_NOTA_GERAL}).")
    if dist_last["n"] and dist_last["negativas"] > TARGET_PCT_NEGATIVAS_MAX:
        st.warning(f"⚠️ Percentual de avaliações negativas nos últimos 30 dias acima da meta ({dist_last['negativas']:.0f}% > {TARGET_PCT_NEGATIVAS_MAX:.0f}%).")

    st.caption(f"Janela dos “últimos 30 dias” no filtro: **{cutoff.date()}** até **{f['timestamp'].max().date()}** (n={n_last}). Notas 8–10 = positivas · 5–7 = neutras · 0–4 = negativas.")

    st.markdown("---")
    st.subheader("Avaliação geral (0–10)")

    col_dist, col_pizza = st.columns([2, 1])
    notas = pd.to_numeric(f["recomendacao_0_10"], errors="coerce").dropna()
    with col_dist:
        if px is None or notas.empty:
            st.write("Sem notas suficientes no recorte atual." if notas.empty else "Instale plotly para gráficos interativos.")
        else:
            contagem = notas.value_counts().reindex(range(0, 11), fill_value=0).reset_index()
            contagem.columns = ["Nota", "Respostas"]
            fign = px.bar(contagem, x="Nota", y="Respostas", title="Distribuição das notas (0–10)")
            fign.update_xaxes(dtick=1)
            style_plotly(fign, height=390)
            st.plotly_chart(fign, use_container_width=True)
    with col_pizza:
        dist_geral = distribuicao_satisfacao(f["recomendacao_0_10"])
        if px is None or dist_geral["n"] == 0:
            st.write("Sem dados suficientes.")
        else:
            pizza = pd.DataFrame(
                {
                    "Grupo": ["Positivas (8-10)", "Neutras (5-7)", "Negativas (0-4)"],
                    "%": [dist_geral["positivas"], dist_geral["neutras"], dist_geral["negativas"]],
                }
            )
            figpz = px.pie(
                pizza, names="Grupo", values="%", title="Avaliações positivas × neutras × negativas",
                color="Grupo",
                color_discrete_map={"Positivas (8-10)": "#21855b", "Neutras (5-7)": "#e59a3a", "Negativas (0-4)": "#c33c3c"},
            )
            style_plotly(figpz, height=390)
            figpz.update_layout(legend={"orientation": "v", "y": -0.08})
            st.plotly_chart(figpz, use_container_width=True)

    st.markdown("---")
    st.subheader("Evolução temporal")

    agg = f.groupby("periodo", as_index=False).agg(
        respostas=("respondent_id", "count"),
        nota_media=("recomendacao_0_10", lambda x: pd.to_numeric(x, errors="coerce").mean()),
        pct_positivas=("recomendacao_0_10", lambda x: distribuicao_satisfacao(x)["positivas"]),
    )

    if px is None:
        st.write("Instale plotly para gráficos interativos: `pip install plotly`")
        st.dataframe(agg, use_container_width=True)
    else:
        colA, colB = st.columns(2)
        with colA:
            fig1 = px.line(agg, x="periodo", y="nota_media", markers=True, title="Nota média geral (0–10)")
            fig1.update_yaxes(range=[0, 10])
            style_plotly(fig1, height=360)
            st.plotly_chart(fig1, use_container_width=True)
        with colB:
            fig2 = px.line(agg, x="periodo", y="pct_positivas", markers=True, title="% de avaliações positivas (nota 8–10)")
            fig2.update_yaxes(range=[0, 100])
            style_plotly(fig2, height=360)
            st.plotly_chart(fig2, use_container_width=True)

        fig3 = px.bar(agg, x="periodo", y="respostas", title="Volume de respostas")
        style_plotly(fig3, height=360)
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")
    st.subheader("Dimensões comuns (média no período selecionado, escala 1–5)")

    dim_rows = []
    for col, label in COMUNS_DIMENSOES:
        s = pd.to_numeric(f[col], errors="coerce")
        dim_rows.append({"Dimensão": label, "Média": float(s.mean()) if s.notna().any() else float("nan")})
    dims = pd.DataFrame(dim_rows).sort_values("Média", ascending=False)

    if px is None:
        st.dataframe(dims, use_container_width=True)
    else:
        figd = px.bar(dims, x="Média", y="Dimensão", orientation="h", title="Média por dimensão comum (1–5)")
        figd.update_xaxes(range=[1, 5])
        style_plotly(figd, height=360)
        st.plotly_chart(figd, use_container_width=True)

    st.markdown("---")
    st.subheader("Indicadores específicos por perfil (escala 1–5)")

    abas = st.tabs(list(DIMENSOES_PERFIL.keys()))
    for aba, perfil in zip(abas, DIMENSOES_PERFIL.keys()):
        with aba:
            sub = f[f["tipo_usuario"] == perfil]
            if sub.empty:
                st.info("Sem respostas desse perfil no recorte atual.")
                continue
            rows = []
            for col, label in DIMENSOES_PERFIL[perfil]:
                s = pd.to_numeric(sub[col], errors="coerce")
                if s.notna().any():
                    rows.append({"Indicador": label, "Média": float(s.mean())})
            if not rows:
                st.info("Sem dados suficientes para os indicadores específicos desse perfil.")
                continue
            dsub = pd.DataFrame(rows).sort_values("Média", ascending=False)
            if px is None:
                st.dataframe(dsub, use_container_width=True)
            else:
                figp = px.bar(dsub, x="Média", y="Indicador", orientation="h", title=f"{perfil} (n={len(sub)})")
                figp.update_xaxes(range=[1, 5])
                style_plotly(figp, height=360)
                st.plotly_chart(figp, use_container_width=True)

    st.markdown("---")
    st.subheader("Cortes rápidos")

    colX, colY = st.columns(2)
    with colX:
        by_unit = f.groupby("unidade", as_index=False).agg(
            respostas=("respondent_id", "count"),
            nota_media=("recomendacao_0_10", lambda x: pd.to_numeric(x, errors="coerce").mean()),
            pct_positivas=("recomendacao_0_10", lambda x: distribuicao_satisfacao(x)["positivas"]),
        ).sort_values("respostas", ascending=False)
        by_unit["nota_media"] = by_unit["nota_media"].round(1)
        by_unit["pct_positivas"] = by_unit["pct_positivas"].round(0)
        by_unit = by_unit.rename(columns={"unidade": "Unidade", "respostas": "Respostas", "nota_media": "Nota média (0–10)", "pct_positivas": "% positivas"})
        st.write("**Por unidade**")
        st.dataframe(by_unit, use_container_width=True, hide_index=True)

    with colY:
        by_tipo = f.groupby("tipo_usuario", as_index=False).agg(
            respostas=("respondent_id", "count"),
            nota_media=("recomendacao_0_10", lambda x: pd.to_numeric(x, errors="coerce").mean()),
            pct_positivas=("recomendacao_0_10", lambda x: distribuicao_satisfacao(x)["positivas"]),
        ).sort_values("respostas", ascending=False)
        by_tipo["nota_media"] = by_tipo["nota_media"].round(1)
        by_tipo["pct_positivas"] = by_tipo["pct_positivas"].round(0)
        by_tipo = by_tipo.rename(columns={"tipo_usuario": "Perfil", "respostas": "Respostas", "nota_media": "Nota média (0–10)", "pct_positivas": "% positivas"})
        st.write("**Por perfil**")
        st.dataframe(by_tipo, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Comentários (últimos 50)")
    comm = f.loc[f["comentario_aberto"].fillna("").astype(str).str.strip() != "", ["timestamp", "unidade", "tipo_usuario", "comentario_aberto"]]
    comm = comm.sort_values("timestamp", ascending=False).head(50)
    comm = comm.rename(columns={"timestamp": "Data/hora", "unidade": "Unidade", "tipo_usuario": "Perfil", "comentario_aberto": "Comentário"})
    if comm.empty:
        st.info("Sem comentários abertos no recorte atual.")
    else:
        st.dataframe(comm, use_container_width=True, hide_index=True)
