"""
app.py — Monitor de Admissões RH
Dependências: pip install streamlit pandas openpyxl msal requests
"""

import json
import logging
import re
import unicodedata
import warnings
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

warnings.filterwarnings("ignore")
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────

PLANILHA_LOCAL = Path(__file__).parent / "Controle_Admissoes.xlsx"
ABA            = "BASE NOVA"
DIAS_ALERTA    = 7

# ══════════════════════════════════════════════
# LÓGICA DE PROCESSAMENTO
# ══════════════════════════════════════════════

def safe_json(obj) -> str:
    raw = json.dumps(obj, ensure_ascii=False, default=str)
    return raw.replace("</", "<\\/")


def _strip_acentos(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    ).upper()


_VALIDACAO_PENDENTE_NORM = {
    _strip_acentos("Pendente - Remuneração"),
    _strip_acentos("Pendente - R&S"),
    _strip_acentos("Em Preenchimento - R&S"),
}

# ── Responsabilidade: COLABORADOR ────────────
_STATUS_COLABORADOR_NORM = {
    _strip_acentos(s) for s in {
        "1 - Contato realizado",
        "10 - Contrato de trabalho enviado",
        "11 - Assinatura do contrato pendente",
        "16 - Documentos pendentes",
        "17 - Exame realizado. Documentos pendentes",
        "20 - Aguardando finalização da jornada de benefícios",
    }
}

# ── Responsabilidade: TIME DE ADMISSÃO ───────
_STATUS_TIME_NORM = {
    _strip_acentos(s) for s in {
        "2 - Candidato não retornou aos contatos",
        "4 - Documentação e ASO enviados",
        "5 - Admissão convertida no HCM",
        "7 - Conferência de cadastro no RM",
        "19 - Data admissão alterada|Gestor e recrutador cientes",
    }
}

# ── Responsabilidade: TERCEIRO (clínica/3778) ─
_STATUS_TERCEIRO_NORM = {
    _strip_acentos(s) for s in {
        "3 - Agendamento do exame solicitado à 3778",
        "18 - Documentação ok . Pendente ASO/enquad.PCD",
    }
}

_LABEL_STATUS = {
    "1 - Contato realizado":                                    "Aguardando retorno do candidato",
    "2 - Candidato não retornou aos contatos":                  "Candidato não retornou — time deve acionar",
    "3 - Agendamento do exame solicitado à 3778":               "Aguardando agendamento pela 3778",
    "4 - Documentação e ASO enviados":                          "Documentação e ASO enviados — acompanhar",
    "5 - Admissão convertida no HCM":                           "Admissão convertida no HCM",
    "7 - Conferência de cadastro no RM":                        "Conferência de cadastro no RM",
    "9 - Admissão finalizada - Gestor/Candidato Comun.":        "Admissão finalizada",
    "10 - Contrato de trabalho enviado":                        "Contrato enviado — aguardando assinatura",
    "11 - Assinatura do contrato pendente":                     "Assinatura do contrato pendente",
    "12 - Contrato de trabalho assinado":                       "Contrato assinado",
    "14 - Admissão finalizada":                                 "Admissão finalizada",
    "15 - Admissão cancelada":                                  "Admissão cancelada",
    "16 - Documentos pendentes":                                "Documentos pendentes com colaborador",
    "17 - Exame realizado. Documentos pendentes":               "Exame realizado — documentos pendentes com colaborador",
    "18 - Documentação ok . Pendente ASO/enquad.PCD":           "Aguardando laudo ASO / enquadramento PCD",
    "19 - Data admissão alterada|Gestor e recrutador cientes":  "Data de admissão alterada",
    "20 - Aguardando finalização da jornada de benefícios":     "Candidato deve concluir jornada de benefícios",
}


def _sv(row, col: str, default: str = "") -> str:
    v = row.get(col, default)
    if v is None:
        return default
    try:
        if pd.isna(v):
            return default
    except (TypeError, ValueError):
        pass
    s = str(v).strip().replace("\n", " ").replace("\r", " ")
    return default if s in ("nan", "NaT", "<NA>") else s


def _sd(row, col: str):
    v = row.get(col)
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        result = pd.to_datetime(v, dayfirst=True)
        return None if pd.isna(result) else result.date()
    except Exception:
        return None


def _nome_responsavel(raw: str) -> str:
    m = re.match(r"^\d+\s*-\s*(.+)", raw or "")
    return m.group(1).strip() if m else (raw or "Não atribuído")


def _classificar(s_adm, s_val, s_proc, s_ben, s_jornada=""):
    s_val_n  = _strip_acentos(s_val)
    s_proc_n = _strip_acentos(s_proc)
    jornada_atribuida = _strip_acentos(s_jornada) == "JORNADA ATRIBUIDA"
    ben_pendente      = s_ben.upper() in ("NÃO INICIADO", "NAO INICIADO", "EM ANDAMENTO")

    # Processo não iniciado
    if s_adm in ("", "nan") or (s_proc in ("", "nan") and s_val in ("", "nan")):
        if s_val_n in _VALIDACAO_PENDENTE_NORM:
            if "REMUNER" in s_val_n:
                return "validacao_remuneracao", "remuneracao", "REMUNERAÇÃO"
            return "validacao_rs", "rs", "RECRUTAMENTO"
        # Jornada já atribuída → responsabilidade do colaborador
        if jornada_atribuida and ben_pendente:
            return "colaborador", s_ben, "COLABORADOR"
        return "nao_iniciado", "aguardando_inicio", "ADMISSÃO"

    # Validação pendente — Remuneração
    if "REMUNER" in s_val_n and s_val_n in _VALIDACAO_PENDENTE_NORM:
        return "validacao_remuneracao", "remuneracao", "REMUNERAÇÃO"

    # Validação pendente — R&S
    if s_val_n in _VALIDACAO_PENDENTE_NORM:
        return "validacao_rs", "rs", "RECRUTAMENTO"

    # Responsabilidade do Colaborador
    if s_proc_n in _STATUS_COLABORADOR_NORM:
        return "colaborador", s_proc, "COLABORADOR"

    # Responsabilidade de Terceiros (clínica/3778)
    if s_proc_n in _STATUS_TERCEIRO_NORM:
        return "terceiro", s_proc, "TERCEIRO"

    # Responsabilidade da Admissão
    if s_proc_n in _STATUS_TIME_NORM:
        return "admissao", s_proc, "ADMISSÃO"

    # Processo não iniciado (status vazio/zero)
    if s_proc in ("", "0", "0.0"):
        if jornada_atribuida and ben_pendente:
            return "colaborador", s_ben, "COLABORADOR"
        return "nao_iniciado", "processo_nao_iniciado", "ADMISSÃO"

    # Benefícios — sempre responsabilidade do Colaborador (jornada já atribuída)
    if s_ben.upper() in ("NÃO INICIADO", "NAO INICIADO", "EM ANDAMENTO"):
        return "colaborador", s_ben, "COLABORADOR"

    return "andamento", s_proc, "—"


def _derivar_etapa(s, sv, ben) -> str:
    sl = s.lower(); svl = sv.lower()
    if not svl and not sl:                                   return "Não iniciado"
    if "pendente - remuner" in svl:                          return "Validação · Remuneração"
    if "pendente - r&s" in svl or "preenchimento" in svl:   return "Validação · R&S"
    if "e aso enviados" in sl:                               return "Em andamento"
    if "document" in sl or "doc" in sl:                      return "Documentação"
    if "contrato" in sl or "assinatura" in sl:               return "Contrato"
    if "agendamento" in sl:                                  return "Exame · Agendamento"
    if "aso" in sl or "exame" in sl:                         return "Exame · ASO"
    if "benefício" in sl or "jornada" in sl:
        return "Benefícios"
    if "contato" in sl or "não retornou" in sl:              return "Contato candidato"
    if "hcm" in sl or "cadastro" in sl:                      return "Importação / Cadastro"
    if "finaliz" in sl or "cancelad" in sl:                  return "Finalizado"
    return "Em andamento"


def _derivar_acao(s, sv, ben) -> str:
    sl = s.lower(); svl = sv.lower()
    if not svl and not sl:                   return "Iniciar processo"
    if "pendente - remuner" in svl:          return "Validar dados de remuneração"
    if "pendente - r&s" in svl or "preenchimento" in svl: return "Concluir validação R&S"
    if "documento" in sl:                    return "Solicitar envio de documentos"
    if "assinatura" in sl:                   return "Solicitar assinatura do contrato"
    if "contrato" in sl:                     return "Acompanhar assinatura"
    if "agendamento" in sl:                  return "Aguardar agendamento pela 3778"
    if "e aso enviados" in sl:               return "Monitorar"
    if "aso" in sl or "pcd" in sl:           return "Aguardar laudo ASO / PCD"
    if "exame" in sl:                        return "Aguardar resultado do exame"
    if "contato" in sl:                      return "Aguardar retorno do candidato"
    if "não retornou" in sl:                 return "Novo contato com candidato"
    if "benefício" in sl or "jornada" in sl:
        return "Concluir jornada de benefícios no Sydle"
    if "e aso enviados" in sl:               return "Monitorar"
    if "alterada" in sl:                     return "Confirmar nova data com gestor"
    if "hcm" in sl:                          return "Converter no HCM"
    if "cadastro" in sl:                     return "Conferir cadastro no RM"
    return "Monitorar"


def _prioridade(dias):
    if dias is None: return "Normal", 4
    if dias < 0:     return "Vencida", 0
    if dias <= 3:    return "Crítica", 1
    if dias <= 7:    return "Alta", 2
    if dias <= 14:   return "Média", 3
    return "Normal", 4


def processar_planilha(caminho: Path, aba: str = "BASE NOVA", dias_alerta: int = 7) -> dict:
    hoje = date.today()

    try:
        df = pd.read_excel(caminho, sheet_name=aba, header=0)
    except ValueError:
        df = pd.read_excel(caminho, sheet_name=0, header=0)

    df.columns = df.columns.str.strip().str.replace("\xa0", " ", regex=False)

    if df.empty:
        return {"registros": [], "stats": {
            "total": 0, "urgentes": 0, "sem_responsavel": 0, "nao_iniciados": 0,
            "por_tipo": {}, "por_responsavel": {},
            "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        }}

    def deve_incluir(row) -> bool:
        s_adm = _sv(row, "Status Admissão")
        if s_adm == "Em processo de admissão":
            return True
        if s_adm in ("Admissão cancelada", "Admitido", "Candidato não retornou ao contato"):
            return False
        if s_adm in ("", "nan"):
            data    = _sd(row, "Data de Admissão")
            tem_fut = data is not None and data >= hoje
            val     = _sv(row, "Status de Validação")
            return tem_fut or val not in ("", "nan")
        return False

    df_ativos = df[df.apply(deve_incluir, axis=1)].copy()

    registros  = []
    contadores = {k: 0 for k in ("nao_iniciado", "validacao_remuneracao", "validacao_rs",
                                  "colaborador", "admissao", "terceiro", "andamento")}
    por_resp   = {"COLABORADOR": 0, "ADMISSÃO": 0, "RECRUTAMENTO": 0,
                  "REMUNERAÇÃO": 0, "TERCEIRO": 0}
    urgentes_total = sem_resp_total = 0

    for _, row in df_ativos.iterrows():
        s_adm      = _sv(row, "Status Admissão")
        s_val      = _sv(row, "Status de Validação")
        s_proc_raw = _sv(row, "Status")
        s_proc     = "" if s_proc_raw in ("0", "0.0") else s_proc_raw
        s_ben      = _sv(row, "Jornada Benefícios - Sydle")

        data_adm      = _sd(row, "Data de Admissão")
        data_inclusao = _sd(row, "Data Inclusão Candidato")

        dias_ate    = (data_adm - hoje).days if data_adm else None
        urgente     = dias_ate is not None and 0 <= dias_ate <= dias_alerta
        s_jornada   = _sv(row, "Status Jornada")
        tipo_pend, _, resp_acao = _classificar(s_adm, s_val, s_proc, s_ben, s_jornada)
        resp_adm    = _nome_responsavel(_sv(row, "Responsável pela Admissão"))
        dias_parado = (hoje - data_inclusao).days if data_inclusao else None
        prior, pord = _prioridade(dias_ate)

        if urgente:                     urgentes_total += 1
        if resp_adm == "Não atribuído": sem_resp_total += 1
        contadores[tipo_pend] = contadores.get(tipo_pend, 0) + 1
        if resp_acao in por_resp:
            por_resp[resp_acao] += 1

        registros.append({
            "nome":             _sv(row, "Nome Candidato") or "Sem nome",
            "cargo":            _sv(row, "Nome do Cargo"),
            "tipo":             _sv(row, "Tipo de Educador"),
            "empresa":          _sv(row, "Empresa / Unidade de Negócio"),
            "local":            _sv(row, "Local"),
            "marca":            _sv(row, "Marca"),
            "resp_admissao":    resp_adm,
            "resp_rs":          _sv(row, "Responsável R&S"),
            "resp_acao":        resp_acao,
            "status_processo":  s_proc or "—",
            "status_validacao": s_val or "—",
            "status_beneficios":s_ben or "—",
            "tipo_pendencia":   tipo_pend,
            "data_admissao":    data_adm.strftime("%d/%m/%Y") if data_adm else "—",
            "dias_admissao":    dias_ate,
            "data_inclusao":    data_inclusao.strftime("%d/%m/%Y") if data_inclusao else "—",
            "dias_parado":      dias_parado,
            "urgente":          urgente,
            "etapa":            _derivar_etapa(s_proc, s_val, s_ben),
            "acao":             _derivar_acao(s_proc, s_val, s_ben),
            "prioridade":       prior,
            "prioridade_ord":   pord,
            "label_status":     _LABEL_STATUS.get(s_proc, s_proc or "Processo não iniciado"),
            "obs":              _sv(row, "Observações - Admissão"),
        })

    ORDEM = {"nao_iniciado": 0, "validacao_remuneracao": 1, "validacao_rs": 2,
             "colaborador": 3, "terceiro": 4, "admissao": 5, "andamento": 6}
    registros.sort(key=lambda r: (
        not r["urgente"],
        ORDEM.get(r["tipo_pendencia"], 9),
        r["dias_admissao"] if r["dias_admissao"] is not None else 9999,
    ))

    return {
        "registros": registros,
        "stats": {
            "total":           len(registros),
            "urgentes":        urgentes_total,
            "sem_responsavel": sem_resp_total,
            "nao_iniciados":   contadores.get("nao_iniciado", 0),
            "por_tipo":        contadores,
            "por_responsavel": por_resp,
            "gerado_em":       datetime.now().strftime("%d/%m/%Y %H:%M"),
        }
    }


# ══════════════════════════════════════════════
# INTERFACE STREAMLIT
# ══════════════════════════════════════════════

st.set_page_config(
    page_title="Monitor de Admissões — Ânima Educação",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  #MainMenu, footer { visibility: hidden; }
  .block-container { padding-top: 1rem; }
  [data-testid="metric-container"] {
    background: #ffffff;
    border: 1px solid #e4e7ec;
    border-radius: 10px;
    padding: 12px 16px;
  }
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=300)
def carregar_dados(planilha: Path, aba: str, dias: int) -> dict:
    return processar_planilha(planilha, aba, dias)


def tentar_baixar():
    try:
        from baixar_sharepoint import baixar_planilha
        with st.spinner("⬇️ Baixando planilha do SharePoint..."):
            baixar_planilha()
        st.cache_data.clear()
        st.success("✅ Planilha atualizada!")
        st.rerun()
    except Exception as e:
        st.error(f"❌ Erro ao baixar: {e}")


# ── Sidebar ───────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Monitor de Admissões")
    st.markdown("---")
    if st.button("🔄 Atualizar planilha agora", use_container_width=True):
        tentar_baixar()
    if PLANILHA_LOCAL.exists():
        mtime = datetime.fromtimestamp(PLANILHA_LOCAL.stat().st_mtime)
        st.caption(f"📁 Última atualização:\n{mtime.strftime('%d/%m/%Y às %H:%M')}")
    else:
        st.warning("⚠️ Planilha não encontrada.")
    st.markdown("---")
    st.caption("Ânima Educação · RH Admissões")

aba_input  = ABA
dias_input = DIAS_ALERTA


if not PLANILHA_LOCAL.exists():
    st.title("📋 Monitor de Admissões")
    st.error("Planilha não encontrada. Clique em 'Atualizar planilha agora' na barra lateral.")
    st.stop()

try:
    dados = carregar_dados(PLANILHA_LOCAL, aba_input, dias_input)
except Exception as e:
    st.error(f"Erro ao processar a planilha: {e}")
    st.stop()

registros = dados["registros"]
stats     = dados["stats"]
pt        = stats["por_tipo"]
pr        = stats["por_responsavel"]


# ── Header ────────────────────────────────────
col_t, col_b = st.columns([6, 1])
with col_t:
    st.title("📋 Monitor de Admissões RH")
    st.caption(f"BASE NOVA · Gerado em {stats['gerado_em']}")
with col_b:
    st.markdown("<br>", unsafe_allow_html=True)
    st.success("● Ao vivo")


# ── Dados para os KPIs ────────────────────────
if "filtro_tipo" not in st.session_state:
    st.session_state.filtro_tipo = ""
if "filtro_val" not in st.session_state:
    st.session_state.filtro_val  = ""

_filtro_tipo = st.session_state.filtro_tipo
_filtro_val  = st.session_state.filtro_val

def aplicar_filtro_global(lista):
    if _filtro_tipo == "urgente":
        return [r for r in lista if r["urgente"]]
    if _filtro_tipo == "nao_iniciado":
        return [r for r in lista if r["tipo_pendencia"] == "nao_iniciado"]
    if _filtro_tipo == "resp_acao" and _filtro_val:
        return [r for r in lista if r["resp_acao"] == _filtro_val]
    return lista

registros_filtrados = aplicar_filtro_global(registros)

_RESP_ORDER = ["RECRUTAMENTO", "REMUNERACAO", "COLABORADOR", "TERCEIRO", "ADMISSAO"]
_RESP_META  = {
    "RECRUTAMENTO": {"label": "Recrutamento", "cor": "#7F77DD", "key": "RECRUTAMENTO"},
    "REMUNERACAO":  {"label": "Remuneração",  "cor": "#BA7517", "key": "REMUNERAÇÃO"},
    "COLABORADOR":  {"label": "Candidato",  "cor": "#185FA5", "key": "COLABORADOR"},
    "TERCEIRO":     {"label": "Exame admissional",    "cor": "#A32D2D", "key": "TERCEIRO"},
    "ADMISSAO":     {"label": "Admissão",     "cor": "#0F6E56", "key": "ADMISSÃO"},
}

total    = stats["total"]
urg      = stats["urgentes"]
ni       = pt.get("nao_iniciado", 0)
max_resp = max(pr.values(), default=1)

prio_data = [
    ("Vencida / Crítica", sum(1 for r in registros if r["prioridade"] in ("Vencida","Crítica")), "#E24B4A"),
    ("Alta",              sum(1 for r in registros if r["prioridade"] == "Alta"),                "#BA7517"),
    ("Média",             sum(1 for r in registros if r["prioridade"] == "Média"),               "#EF9F27"),
    ("Normal",            sum(1 for r in registros if r["prioridade"] == "Normal"),              "#B4B2A9"),
]
max_prio = max(v for _, v, _ in prio_data) or 1

# ── CSS: hover nos cards ──────────────────────
st.markdown("""
<style>
.kpi-card {
    background: white;
    border-radius: 10px;
    padding: 1rem 1.15rem;
    position: relative;
    overflow: hidden;
    transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease;
    cursor: default;
}
.kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 20px rgba(0,0,0,.09);
}
.kpi-hero { border: 1.5px solid #e4e7ec; min-height: 128px; }
.kpi-resp { border: 1.5px solid #e4e7ec; min-height: 90px; }
.kpi-hero:hover { border-color: #d0d5dd; }
.kpi-resp:hover { border-color: #d0d5dd; }
.kpi-top-bar {
    position: absolute; top: 0; left: 0; right: 0; height: 3px;
}
.kpi-label {
    font-size: 10px; color: #475467; font-weight: 600;
    text-transform: uppercase; letter-spacing: .05em; margin-bottom: 6px;
}
.kpi-value { font-size: 36px; font-weight: 500; line-height: 1; margin-bottom: 4px; }
.kpi-sub   { font-size: 11px; color: #98a2b3; }
.kpi-badge {
    display: inline-block; margin-top: 8px;
    font-size: 10px; font-weight: 600;
    padding: 2px 9px; border-radius: 20px;
}
.kpi-resp-label { font-size: 10px; color: #475467; font-weight: 600; margin-bottom: 4px; }
.kpi-resp-value { font-size: 22px; font-weight: 500; margin-bottom: 7px; }
.kpi-bar-bg { height: 4px; background: #e4e7ec; border-radius: 2px; overflow: hidden; }
.kpi-bar    { height: 4px; border-radius: 2px; }
</style>
""", unsafe_allow_html=True)

def hero_card(titulo, valor, sub, badge_txt, badge_bg, badge_color, top_cor):
    return (
        f'<div class="kpi-card kpi-hero">' 
        f'<div class="kpi-top-bar" style="background:{top_cor};"></div>'
        f'<div class="kpi-label">{titulo}</div>'
        f'<div class="kpi-value" style="color:{top_cor};">{valor}</div>'
        f'<div class="kpi-sub">{sub}</div>'
        f'<span class="kpi-badge" style="background:{badge_bg};color:{badge_color};">{badge_txt}</span>'
        f'</div>'
    )

def resp_card(label, val, pct, cor):
    return (
        f'<div class="kpi-card kpi-resp">' 
        f'<div class="kpi-top-bar" style="background:{cor};"></div>'
        f'<div class="kpi-resp-label">{label}</div>'
        f'<div class="kpi-resp-value" style="color:{cor};">{val}</div>'
        f'<div class="kpi-bar-bg"><div class="kpi-bar" style="width:{pct}%;background:{cor};"></div></div>'
        f'</div>'
    )

# ── Linha 1: Hero cards ───────────────────────
h1, h2, h3 = st.columns(3)
with h1:
    st.markdown(hero_card("Total monitorado", total, "processos ativos agora",
        "BASE NOVA", "#F1EFE8", "#5F5E5A", "#888780"), unsafe_allow_html=True)
with h2:
    st.markdown(hero_card("Urgentes", urg, "admissão em ≤7 dias",
        "requer atenção imediata", "#FCEBEB", "#A32D2D", "#E24B4A"), unsafe_allow_html=True)
with h3:
    st.markdown(hero_card("Não iniciados", ni, "sem processo em andamento",
        "processo travado", "#FAEEDA", "#854F0B", "#BA7517"), unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ── Linha 2: Cards por responsável ───────────
r_cols = st.columns(5)
for i, rk in enumerate(_RESP_ORDER):
    m   = _RESP_META[rk]
    val = pr.get(m["key"], 0)
    pct = round((val / max_resp) * 100) if max_resp else 0
    with r_cols[i]:
        st.markdown(resp_card(m["label"], val, pct, m["cor"]), unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ── Linha 3: Nível de urgência ────────────────
prio_rows = "".join(
    f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">' +
    f'<div style="width:8px;height:8px;border-radius:50%;background:{c};flex-shrink:0;"></div>' +
    f'<div style="font-size:12px;color:#475467;flex:1;">{lbl}</div>' +
    f'<div style="width:100px;height:6px;background:#e4e7ec;border-radius:3px;overflow:hidden;">' +
    f'<div style="height:6px;width:{round((v/max_prio)*100)}%;background:{c};border-radius:3px;"></div></div>' +
    f'<div style="font-size:12px;font-weight:500;color:#101828;width:24px;text-align:right;">{v}</div></div>'
    for lbl, v, c in prio_data
)

st.markdown(
    f'<div style="background:white;border:1.5px solid #e4e7ec;border-radius:10px;padding:1rem 1.25rem;">' +
    f'<div style="font-size:10px;color:#475467;font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">Nível de urgência</div>' +
    prio_rows + '</div>', unsafe_allow_html=True
)

st.markdown("---")


# ══════════════════════════════════════════════
# ABAS
# ══════════════════════════════════════════════
aba_gargalos, aba_analise, aba_kanban, aba_tabela = st.tabs([
    "📊 Gargalos por Responsável",
    "📈 Análise",
    "🗂 Kanban por Etapa",
    "☰ Tabela detalhada",
])


# ─── GARGALOS ─────────────────────────────────
with aba_gargalos:
    st.markdown("### Onde estão os processos parados?")

    # ── Cards KPI por responsável (ordem do processo) ──
    col_a, col_b, col_c, col_d, col_e = st.columns(5)

    def card_resp_gargalo(col, emoji, label, valor, cor, help_txt):
        col.markdown(f"""
        <div style="background:#fff;border:1.5px solid #e4e7ec;border-top:4px solid {cor};
                    border-radius:10px;padding:16px;text-align:center;
                    transition:transform .18s ease,box-shadow .18s ease;"
             onmouseover="this.style.transform='translateY(-3px)';this.style.boxShadow='0 8px 20px rgba(0,0,0,.09)'"
             onmouseout="this.style.transform='';this.style.boxShadow=''">
            <div style="font-size:28px;font-weight:600;color:{cor};">{valor}</div>
            <div style="font-size:13px;font-weight:600;margin-top:4px;">{emoji} {label}</div>
            <div style="font-size:11px;color:#98a2b3;margin-top:2px;">{help_txt}</div>
        </div>
        """, unsafe_allow_html=True)

    card_resp_gargalo(col_a, "✏️", "Recrutamento", pr.get("RECRUTAMENTO", 0), "#7F77DD", "validação R&S")
    card_resp_gargalo(col_b, "💰", "Remuneração",  pr.get("REMUNERAÇÃO", 0),  "#BA7517", "validação salarial")
    card_resp_gargalo(col_c, "👤", "Candidato",  pr.get("COLABORADOR", 0),  "#185FA5", "ação do candidato")
    card_resp_gargalo(col_d, "🔬", "Exame admissional",    pr.get("TERCEIRO", 0),     "#A32D2D", "clínica / 3778")
    card_resp_gargalo(col_e, "🏢", "Admissão",     pr.get("ADMISSÃO", 0),     "#0F6E56", "ação interna")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # ── Gargalo por responsável (barras) ──────
    _max_r = max(pr.values(), default=1)
    _garg_rows = "".join(
        f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">' +
        f'<div style="width:110px;font-size:12px;color:#475467;flex-shrink:0;">{_RESP_META[rk]["label"]}</div>' +
        f'<div style="flex:1;height:10px;background:#e4e7ec;border-radius:5px;overflow:hidden;">' +
        f'<div style="height:10px;width:{round((pr.get(_RESP_META[rk]["key"],0)/_max_r)*100) if _max_r else 0}%;background:{_RESP_META[rk]["cor"]};border-radius:5px;"></div></div>' +
        f'<div style="width:32px;font-size:13px;font-weight:600;color:#101828;text-align:right;">{pr.get(_RESP_META[rk]["key"],0)}</div></div>'
        for rk in _RESP_ORDER
    )
    st.markdown(
        f'<div style="background:white;border:1.5px solid #e4e7ec;border-radius:10px;padding:1.1rem 1.4rem;margin-bottom:16px;">' +
        f'<div style="font-size:11px;color:#475467;font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-bottom:14px;">Processos parados por responsável</div>' +
        _garg_rows + '</div>', unsafe_allow_html=True
    )

    st.markdown("---")

    # Kanban por responsável
    st.markdown("Visão dos processos agrupados pelo responsável pela próxima ação.")
    dados_json_resp = safe_json({"registros": registros_filtrados})

    kanban_resp_html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<style>
:root {{
  --surface:#fff; --surface2:#f7f8fa; --border:#e4e7ec; --border2:#d0d5dd;
  --text:#101828; --text2:#475467; --text3:#98a2b3;
  --red:#d92d20; --red-bg:#fff1f0; --red-border:#fda29b;
  --orange:#c4320a; --orange-bg:#fff4ed; --orange-border:#f9c6a8;
  --yellow:#b54708; --yellow-bg:#fffaeb; --yellow-border:#fec84b;
  --blue:#1d4ed8; --blue-bg:#eff6ff; --blue-border:#93c5fd;
  --teal:#107569; --teal-bg:#f0fdf9; --teal-border:#5fe3c0;
  --purple:#6941c6; --purple-bg:#f4f3ff; --purple-border:#c3b5fd;
  --gray:#344054; --gray-bg:#f9fafb; --gray-border:#d0d5dd;
}}
* {{ box-sizing:border-box; margin:0; padding:0; font-family:Segoe UI,Arial,sans-serif; }}
body {{ background:#f0f2f5; font-size:13px; padding:8px; }}
.board {{ display:flex; gap:12px; overflow-x:auto; padding-bottom:8px; width:100%; }}
.col {{ flex:1 1 0; min-width:180px; background:#fff; border:1px solid var(--border); border-radius:10px; overflow:hidden; display:flex; flex-direction:column; }}
.col-header {{ padding:12px 14px; border-bottom:1px solid var(--border); display:flex; align-items:center; gap:8px; background:var(--surface2); }}
.col-title {{ font-size:12px; font-weight:700; flex:1; }}
.col-count {{ font-size:10px; font-weight:600; background:#fff; border:1px solid var(--border2); border-radius:10px; padding:1px 8px; color:var(--text2); }}
.col-body {{ padding:8px; overflow-y:visible; flex:1; }}
.card {{ background:var(--surface2); border:1px solid var(--border); border-radius:7px; padding:8px 10px; margin-bottom:6px; }}
.card.urg  {{ border-left:3px solid var(--red); }}
.card.alta {{ border-left:3px solid var(--orange); }}
.card.med  {{ border-left:3px solid var(--yellow); }}
.card-nome  {{ font-size:11px; font-weight:600; margin-bottom:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.card-cargo {{ font-size:10px; color:var(--text3); margin-bottom:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.card-acao  {{ font-size:10px; color:var(--text2); margin-bottom:4px; }}
.card-foot  {{ display:flex; gap:4px; flex-wrap:wrap; }}
.badge {{ display:inline-block; border-radius:5px; padding:2px 6px; font-size:9px; font-weight:600; border:1px solid transparent; }}
.b-vencida {{ background:var(--red-bg);    color:var(--red);    border-color:var(--red-border); }}
.b-critica {{ background:var(--red-bg);    color:var(--red);    border-color:var(--red-border); }}
.b-alta    {{ background:var(--orange-bg); color:var(--orange); border-color:var(--orange-border); }}
.b-media   {{ background:var(--yellow-bg); color:var(--yellow); border-color:var(--yellow-border); }}
.b-normal  {{ background:var(--gray-bg);   color:var(--gray);   border-color:var(--gray-border); }}
.b-etapa   {{ background:var(--purple-bg); color:var(--purple); border-color:var(--purple-border); }}
.col-rec  .col-header {{ border-top:4px solid #7F77DD; }}
.col-rem  .col-header {{ border-top:4px solid #BA7517; }}
.col-colab .col-header {{ border-top:4px solid #185FA5; }}
.col-ter  .col-header {{ border-top:4px solid #A32D2D; }}
.col-adm  .col-header {{ border-top:4px solid #0F6E56; }}
.mais {{ text-align:center; padding:6px; font-size:10px; color:var(--text3); border-top:1px solid var(--border); }}
.urg-badge {{ font-size:9px; background:#fef2f2; color:#dc2626; border-radius:10px; padding:1px 6px; font-weight:600; }}
</style></head><body>
<div class="board" id="board"></div>
<script>
const DADOS = {dados_json_resp};
const registros = DADOS.registros;
const PRIOR_CFG = {{
  'Vencida':{{label:'⚠️ Vencida',cls:'b-vencida',urg:'urg'}},
  'Crítica':{{label:'🔴 Crítica',cls:'b-critica',urg:'urg'}},
  'Alta':   {{label:'🟠 Alta',   cls:'b-alta',   urg:'alta'}},
  'Média':  {{label:'🟡 Média',  cls:'b-media',  urg:'med'}},
  'Normal': {{label:'⚪ Normal', cls:'b-normal', urg:''}},
}};
const COLUNAS = [
  {{resp:'RECRUTAMENTO', cls:'col-rec',   icon:'✏️', label:'Recrutamento', sub:'validação R&S'}},
  {{resp:'REMUNERAÇÃO',  cls:'col-rem',   icon:'💰', label:'Remuneração',  sub:'validação salarial'}},
  {{resp:'COLABORADOR',  cls:'col-colab', icon:'👤', label:'Candidato',  sub:'ação do candidato'}},
  {{resp:'TERCEIRO',     cls:'col-ter',   icon:'🔬', label:'Exame admissional',    sub:'clínica / 3778'}},
  {{resp:'ADMISSÃO',     cls:'col-adm',   icon:'🏢', label:'Admissão',     sub:'ação interna'}},
];
function esc(s){{return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}}
const grupos={{}};
registros.forEach(r=>{{
  if(!grupos[r.resp_acao])grupos[r.resp_acao]=[];
  grupos[r.resp_acao].push(r);
}});
Object.values(grupos).forEach(l=>l.sort((a,b)=>a.prioridade_ord-b.prioridade_ord||(a.dias_admissao||9999)-(b.dias_admissao||9999)));
const MAX=30;
document.getElementById('board').innerHTML=COLUNAS.filter(c=>grupos[c.resp]?.length>0).map(col=>{{
  const lista=grupos[col.resp]||[];
  const urg=lista.filter(r=>['Vencida','Crítica','Alta'].includes(r.prioridade)).length;
  const cards=lista.slice(0,MAX).map(r=>{{
    const pc=PRIOR_CFG[r.prioridade]||PRIOR_CFG['Normal'];
    const adm=r.dias_admissao!==null?(r.dias_admissao<0?`⚠️ ${{Math.abs(r.dias_admissao)}}d atraso`:r.dias_admissao===0?'🔴 HOJE':r.urgente?`🟡 ${{r.dias_admissao}}d`:r.data_admissao):r.data_admissao;
    return `<div class="card ${{pc.urg}}">
      <div class="card-nome">${{esc(r.nome)}}</div>
      <div class="card-cargo">${{esc(r.cargo)}}</div>
      <div class="card-acao">→ ${{esc(r.acao)}}</div>
      <div class="card-foot">
        <span class="badge ${{pc.cls}}">${{pc.label}}</span>
        <span class="badge b-etapa">${{esc(r.etapa)}}</span>
      </div>
      <div style="font-size:10px;color:#98a2b3;margin-top:4px;">${{esc(adm)}} · ${{esc(r.resp_admissao)}}</div>
    </div>`;
  }}).join('');
  const mais=lista.length>MAX?`<div class="mais">+ ${{lista.length-MAX}} candidatos</div>`:'';
  const urgB=urg>0?`<span class="urg-badge">${{urg}} urg.</span>`:'';
  return `<div class="col ${{col.cls}}">
    <div class="col-header">
      <span style="font-size:16px;">${{col.icon}}</span>
      <div><div class="col-title">${{col.label}}</div><div style="font-size:10px;color:#98a2b3;">${{col.sub}}</div></div>
      ${{urgB}}<span class="col-count">${{lista.length}}</span>
    </div>
    <div class="col-body">${{cards}}${{mais}}</div>
  </div>`;
}}).join('');
// Auto-resize
function autoResize() {{
  const h = document.body.scrollHeight;
  window.parent.postMessage({{type:'streamlit:setFrameHeight', height: h}}, '*');
}}
window.addEventListener('load', autoResize);
setTimeout(autoResize, 300);
</script></body></html>"""

    components.html(kanban_resp_html, height=700, scrolling=True)


# ─── ANÁLISE ──────────────────────────────────
with aba_analise:

    # ── Dados auxiliares ─────────────────────
    from datetime import timedelta
    from calendar import monthrange

    hoje_dt = date.today()

    # Admissões por dia (próximos 2 meses)
    admissoes_por_dia = {}
    for r in registros:
        if r["data_admissao"] == "—":
            continue
        try:
            dt = datetime.strptime(r["data_admissao"], "%d/%m/%Y").date()
        except Exception:
            continue
        if dt >= hoje_dt - timedelta(days=7):
            admissoes_por_dia[dt] = admissoes_por_dia.get(dt, 0) + 1

    # Tempo médio total do processo
    dias_validos = [r["dias_parado"] for r in registros if r["dias_parado"] is not None]
    media_total  = round(sum(dias_validos) / len(dias_validos)) if dias_validos else 0

    # Distribuição por marca
    marca_count = {}
    for r in registros:
        m = r["marca"] or "Sem marca"
        marca_count[m] = marca_count.get(m, 0) + 1
    marca_ord  = sorted(marca_count.items(), key=lambda x: -x[1])[:12]
    max_marca  = max((v for _, v in marca_ord), default=1)

    # ══════════════════════════════════════════
    # Card: Tempo médio total
    # ══════════════════════════════════════════
    st.markdown("### ⏱ Tempo médio do processo")

    cor_media = "#A32D2D" if media_total > 30 else "#BA7517" if media_total > 14 else "#0F6E56"
    bg_media  = "#FCEBEB" if media_total > 30 else "#FAEEDA" if media_total > 14 else "#E1F5EE"
    txt_media = "acima do ideal" if media_total > 30 else "atenção recomendada" if media_total > 14 else "dentro do esperado"

    st.markdown(
        '<div style="background:white;border:1.5px solid #e4e7ec;border-top:3px solid ' + cor_media + ';'
        'border-radius:10px;padding:1.1rem 1.4rem;display:flex;align-items:center;gap:20px;">'
        '<div style="font-size:56px;font-weight:500;line-height:1;color:' + cor_media + ';">' + str(media_total) + '</div>'
        '<div>'
        '<div style="font-size:11px;color:#475467;font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;">dias em média</div>'
        '<div style="font-size:12px;color:#98a2b3;">do início do processo até hoje</div>'
        '<span style="display:inline-block;margin-top:6px;font-size:10px;font-weight:600;padding:2px 9px;'
        'border-radius:20px;background:' + bg_media + ';color:' + cor_media + ';">' + txt_media + '</span>'
        '</div>'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ══════════════════════════════════════════
    # Calendário — grid de dias
    # ══════════════════════════════════════════
    st.markdown("### 📅 Admissões previstas")

    max_dia = max(admissoes_por_dia.values(), default=1)

    def cor_intensidade(n):
        if n == 0:   return "var(--color-border-tertiary)", "transparent"
        if n <= 2:   return "#B5D4F4", "#0C447C"
        if n <= 6:   return "#85B7EB", "#0C447C"
        return "#378ADD", "white"

    # Gera meses necessários
    meses_necessarios = set()
    for dt in admissoes_por_dia:
        meses_necessarios.add((dt.year, dt.month))
    # Adiciona mês atual e próximo mesmo que vazios
    meses_necessarios.add((hoje_dt.year, hoje_dt.month))
    prox = hoje_dt.replace(day=1) + timedelta(days=32)
    meses_necessarios.add((prox.year, prox.month))
    meses_ord = sorted(meses_necessarios)[:3]

    nomes_mes = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    dias_sem  = ["Seg","Ter","Qua","Qui","Sex","Sáb","Dom"]

    cal_html_parts = [
        '<!DOCTYPE html><html><head><meta charset="UTF-8">',
        '<style>',
        '* { box-sizing:border-box; margin:0; padding:0; font-family:Segoe UI,Arial,sans-serif; }',
        'body { background:transparent; padding:0; }',
        '.meses { display:flex; gap:24px; flex-wrap:wrap; }',
        '.mes { flex:1; min-width:200px; }',
        '.mes-titulo { font-size:12px; font-weight:600; color:#344054; margin-bottom:8px; }',
        '.grid { display:grid; grid-template-columns:repeat(7,1fr); gap:3px; }',
        '.hdr { font-size:9px; font-weight:600; color:#98a2b3; text-align:center; padding:2px 0; }',
        '.dia { aspect-ratio:1; border-radius:5px; display:flex; flex-direction:column;',
        '       align-items:center; justify-content:center; font-size:10px; font-weight:500;',
        '       border:0.5px solid transparent; cursor:default; }',
        '.dia.vazio { background:transparent; border-color:transparent; }',
        '.dia .dn { line-height:1; }',
        '.dia .cnt { font-size:8px; font-weight:700; line-height:1; margin-top:1px; }',
        '.dia.hoje { outline:2px solid #378ADD; outline-offset:1px; border-radius:5px; }',
        '.legenda { display:flex; gap:10px; margin-top:10px; align-items:center; font-size:10px; color:#98a2b3; }',
        '.leg-sq { width:10px; height:10px; border-radius:2px; flex-shrink:0; }',
        '</style></head><body>',
        '<div class="meses">',
    ]

    for ano, mes in meses_ord:
        _, dias_no_mes = monthrange(ano, mes)
        primeiro_dia   = date(ano, mes, 1)
        offset         = primeiro_dia.weekday()  # 0=Seg

        cal_html_parts.append('<div class="mes">')
        cal_html_parts.append(f'<div class="mes-titulo">{nomes_mes[mes-1]} {ano}</div>')
        cal_html_parts.append('<div class="grid">')

        for d in dias_sem:
            cal_html_parts.append(f'<div class="hdr">{d}</div>')

        for _ in range(offset):
            cal_html_parts.append('<div class="dia vazio"></div>')

        for dia in range(1, dias_no_mes + 1):
            dt      = date(ano, mes, dia)
            n       = admissoes_por_dia.get(dt, 0)
            bg, fc  = cor_intensidade(n)
            is_hoje = dt == hoje_dt
            extra   = ' hoje' if is_hoje else ''
            style   = f'background:{bg};border-color:{bg};color:{"#475467" if n == 0 else fc};'
            cal_html_parts.append(
                f'<div class="dia{extra}" style="{style}">'
                f'<div class="dn">{dia}</div>'
                + (f'<div class="cnt">{n}</div>' if n > 0 else '')
                + '</div>'
            )

        cal_html_parts.append('</div></div>')

    cal_html_parts += [
        '</div>',
        '<div class="legenda">',
        '<span>Volume:</span>',
        '<span class="leg-sq" style="background:var(--color-border-tertiary);border:0.5px solid #e4e7ec;"></span><span>0</span>',
        '<span class="leg-sq" style="background:#B5D4F4;"></span><span>1–2</span>',
        '<span class="leg-sq" style="background:#85B7EB;"></span><span>3–6</span>',
        '<span class="leg-sq" style="background:#378ADD;"></span><span>7+</span>',
        '</div>',
        '</body></html>',
    ]

    cal_html_final = "".join(cal_html_parts)
    n_meses = len(meses_ord)
    components.html(cal_html_final, height=n_meses * 240 + 60, scrolling=False)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ══════════════════════════════════════════
    # Distribuição por marca/unidade
    # ══════════════════════════════════════════
    st.markdown("### 🏢 Processos por marca / unidade")

    if not marca_ord:
        st.info("Sem dados de marca.")
    else:
        marca_rows = ""
        for marca, count in marca_ord:
            pct = round((count / max_marca) * 100)
            marca_rows += (
                '<div style="margin-bottom:10px;">'
                '<div style="display:flex;justify-content:space-between;margin-bottom:3px;">'
                '<span style="font-size:11px;color:#475467;font-weight:500;">' + marca + '</span>'
                '<span style="font-size:11px;font-weight:700;color:#185FA5;">' + str(count) + '</span>'
                '</div>'
                '<div style="height:8px;background:#e4e7ec;border-radius:4px;overflow:hidden;">'
                '<div style="height:8px;width:' + str(pct) + '%;background:#185FA5;border-radius:4px;"></div>'
                '</div></div>'
            )
        st.markdown(
            '<div style="background:white;border:1.5px solid #e4e7ec;border-radius:10px;padding:1.1rem 1.4rem;">'
            + marca_rows + '</div>',
            unsafe_allow_html=True
        )


# ─── KANBAN POR ETAPA ─────────────────────────
with aba_kanban:
    st.markdown("Visão do processo por etapa — ordenado por urgência.")
    dados_kanban = dict(dados); dados_kanban["registros"] = registros_filtrados
    dados_json = safe_json(dados_kanban)

    kanban_html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<style>
:root {{
  --surface:#fff; --surface2:#f7f8fa; --border:#e4e7ec; --border2:#d0d5dd;
  --text:#101828; --text2:#475467; --text3:#98a2b3;
  --red:#d92d20; --red-bg:#fff1f0; --red-border:#fda29b;
  --orange:#c4320a; --orange-bg:#fff4ed; --orange-border:#f9c6a8;
  --yellow:#b54708; --yellow-bg:#fffaeb; --yellow-border:#fec84b;
  --blue:#1d4ed8; --blue-bg:#eff6ff; --blue-border:#93c5fd;
  --teal:#107569; --teal-bg:#f0fdf9; --teal-border:#5fe3c0;
  --purple:#6941c6; --purple-bg:#f4f3ff; --purple-border:#c3b5fd;
  --gray:#344054; --gray-bg:#f9fafb; --gray-border:#d0d5dd;
}}
* {{ box-sizing:border-box; margin:0; padding:0; font-family:Segoe UI,Arial,sans-serif; }}
body {{ background:#f0f2f5; font-size:13px; padding:8px; }}
.board {{ display:flex; gap:10px; overflow-x:auto; padding-bottom:8px; width:100%; }}
.col {{ flex:1 1 0; min-width:160px; background:#fff; border:1px solid var(--border); border-radius:10px; overflow:hidden; display:flex; flex-direction:column; }}
.col-header {{ padding:10px 12px; border-bottom:1px solid var(--border); display:flex; align-items:center; gap:8px; background:var(--surface2); }}
.col-title {{ font-size:12px; font-weight:600; flex:1; }}
.col-count {{ font-size:10px; font-weight:600; background:#fff; border:1px solid var(--border2); border-radius:10px; padding:1px 7px; color:var(--text2); }}
.col-body {{ padding:8px; overflow-y:visible; flex:1; }}
.card {{ background:var(--surface2); border:1px solid var(--border); border-radius:7px; padding:8px 10px; margin-bottom:6px; }}
.card.urg  {{ border-left:3px solid var(--red); }}
.card.alta {{ border-left:3px solid var(--orange); }}
.card.med  {{ border-left:3px solid var(--yellow); }}
.card-nome  {{ font-size:11px; font-weight:600; margin-bottom:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.card-cargo {{ font-size:10px; color:var(--text3); margin-bottom:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.card-foot  {{ display:flex; gap:4px; flex-wrap:wrap; }}
.badge {{ display:inline-block; border-radius:5px; padding:2px 6px; font-size:9px; font-weight:600; border:1px solid transparent; }}
.b-vencida {{ background:var(--red-bg);    color:var(--red);    border-color:var(--red-border); }}
.b-critica {{ background:var(--red-bg);    color:var(--red);    border-color:var(--red-border); }}
.b-alta    {{ background:var(--orange-bg); color:var(--orange); border-color:var(--orange-border); }}
.b-media   {{ background:var(--yellow-bg); color:var(--yellow); border-color:var(--yellow-border); }}
.b-normal  {{ background:var(--gray-bg);   color:var(--gray);   border-color:var(--gray-border); }}
.b-colab   {{ background:var(--blue-bg);   color:var(--blue);   border-color:var(--blue-border); }}
.b-time    {{ background:var(--teal-bg);   color:var(--teal);   border-color:var(--teal-border); }}
.b-rec     {{ background:var(--yellow-bg); color:var(--yellow); border-color:var(--yellow-border); }}
.b-ter     {{ background:var(--red-bg);    color:var(--red);    border-color:var(--red-border); }}
.col-ni    .col-header {{ border-top:3px solid #dc2626; }}
.col-val   .col-header {{ border-top:3px solid #d97706; }}
.col-doc   .col-header {{ border-top:3px solid #1d4ed8; }}
.col-exame .col-header {{ border-top:3px solid #0f766e; }}
.col-cont  .col-header {{ border-top:3px solid #7c3aed; }}
.col-ben   .col-header {{ border-top:3px solid #059669; }}
.col-and   .col-header {{ border-top:3px solid #6b7280; }}
.mais {{ text-align:center; padding:6px; font-size:10px; color:var(--text3); border-top:1px solid var(--border); }}
.urg-badge {{ font-size:9px; background:#fef2f2; color:#dc2626; border-radius:10px; padding:1px 5px; font-weight:600; }}
</style></head><body>
<div class="board" id="board"></div>
<script>
const DADOS = {dados_json};
const registros = DADOS.registros;
const PRIOR_CFG = {{
  'Vencida':{{label:'⚠️ Vencida',cls:'b-vencida',urg:'urg'}},
  'Crítica':{{label:'🔴 Crítica',cls:'b-critica',urg:'urg'}},
  'Alta':   {{label:'🟠 Alta',   cls:'b-alta',   urg:'alta'}},
  'Média':  {{label:'🟡 Média',  cls:'b-media',  urg:'med'}},
  'Normal': {{label:'⚪ Normal', cls:'b-normal', urg:''}},
}};
const RESP_CFG = {{
  'COLABORADOR': {{label:'👤 Candidato', cls:'b-colab'}},
  'ADMISSÃO':    {{label:'🏢 Admissão',    cls:'b-time'}},
  'RECRUTAMENTO':{{label:'✏️ Recrutamento',cls:'b-rec'}},
  'REMUNERAÇÃO': {{label:'💰 Remuneração', cls:'b-rec'}},
  'TERCEIRO':    {{label:'🔬 Terceiro',    cls:'b-ter'}},
  '—':           {{label:'⚠️ Sem atrib.', cls:'b-vencida'}},
}};
const COLUNAS = [
  {{etapa:'Não iniciado',           cls:'col-ni',   icon:'⚠️'}},
  {{etapa:'Validação · R&S',        cls:'col-val',  icon:'✏️'}},
  {{etapa:'Validação · Remuneração',cls:'col-val',  icon:'💰'}},
  {{etapa:'Contato candidato',      cls:'col-cont', icon:'📞'}},
  {{etapa:'Documentação',           cls:'col-doc',  icon:'📄'}},
  {{etapa:'Exame · Agendamento',    cls:'col-exame',icon:'📅'}},
  {{etapa:'Exame · ASO',            cls:'col-exame',icon:'🩺'}},
  {{etapa:'Contrato',               cls:'col-doc',  icon:'📝'}},
  {{etapa:'Benefícios',             cls:'col-ben',  icon:'🎁'}},
  {{etapa:'Importação / Cadastro',  cls:'col-and',  icon:'💻'}},
  {{etapa:'Em andamento',           cls:'col-and',  icon:'✅'}},
  {{etapa:'Finalizado',             cls:'col-and',  icon:'✔️'}},
];
function esc(s){{return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}}
const grupos={{}};
registros.forEach(r=>{{if(!grupos[r.etapa])grupos[r.etapa]=[];grupos[r.etapa].push(r);}});
Object.values(grupos).forEach(l=>l.sort((a,b)=>a.prioridade_ord-b.prioridade_ord||(a.dias_admissao||9999)-(b.dias_admissao||9999)));
const colunas=COLUNAS.filter(c=>grupos[c.etapa]?.length>0);
Object.keys(grupos).forEach(e=>{{if(!COLUNAS.find(c=>c.etapa===e))colunas.push({{etapa:e,cls:'col-and',icon:'📌'}});}});
document.getElementById('board').innerHTML=colunas.map(col=>{{
  const lista=grupos[col.etapa]||[];
  const urg=lista.filter(r=>['Vencida','Crítica','Alta'].includes(r.prioridade)).length;
  const MAX=25;
  const cards=lista.slice(0,MAX).map(r=>{{
    const pc=PRIOR_CFG[r.prioridade]||PRIOR_CFG['Normal'];
    const rc=RESP_CFG[r.resp_acao]||RESP_CFG['—'];
    const adm=r.dias_admissao!==null?(r.dias_admissao<0?`⚠️ ${{Math.abs(r.dias_admissao)}}d atraso`:r.dias_admissao===0?'🔴 HOJE':r.urgente?`🟡 ${{r.dias_admissao}}d`:r.data_admissao):r.data_admissao;
    return `<div class="card ${{pc.urg}}" title="${{esc(r.obs||'')}}"><div class="card-nome">${{esc(r.nome)}}</div><div class="card-cargo">${{esc(r.cargo)}}</div><div class="card-foot"><span class="badge ${{pc.cls}}">${{pc.label}}</span><span class="badge ${{rc.cls}}">${{rc.label}}</span></div><div style="font-size:10px;color:#98a2b3;margin-top:4px;">${{esc(adm)}} · ${{esc(r.resp_admissao)}}</div></div>`;
  }}).join('');
  const mais=lista.length>MAX?`<div class="mais">+ ${{lista.length-MAX}} candidatos</div>`:'';
  const urgB=urg>0?`<span class="urg-badge">${{urg}} urg.</span>`:'';
  return `<div class="col ${{col.cls}}"><div class="col-header"><span style="font-size:14px;">${{col.icon}}</span><span class="col-title">${{esc(col.etapa)}}</span>${{urgB}}<span class="col-count">${{lista.length}}</span></div><div class="col-body">${{cards}}${{mais}}</div></div>`;
}}).join('');
function autoResize() {{
  const h = document.body.scrollHeight;
  window.parent.postMessage({{type:'streamlit:setFrameHeight', height: h}}, '*');
}}
window.addEventListener('load', autoResize);
window.addEventListener('resize', autoResize);
setTimeout(autoResize, 300);
</script></body></html>"""

    components.html(kanban_html, height=700, scrolling=True)


# ─── TABELA ───────────────────────────────────
with aba_tabela:
    st.markdown("Tabela detalhada com filtros.")

    col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
    with col_f1:
        busca = st.text_input("🔍 Buscar", "")
    with col_f2:
        resps  = ["Todos"] + sorted({r["resp_admissao"] for r in registros if r["resp_admissao"]})
        f_resp = st.selectbox("Resp. admissão", resps)
    with col_f3:
        f_prior = st.selectbox("Prioridade", ["Todas", "Vencida", "Crítica", "Alta", "Média", "Normal"])
    with col_f4:
        resp_acoes = {
            "Todos": "",
            "✏️ Recrutamento":  "RECRUTAMENTO",
            "💰 Remuneração":   "REMUNERAÇÃO",
            "👤 Candidato":   "COLABORADOR",
            "🔬 Exame admissional":     "TERCEIRO",
            "🏢 Admissão":      "ADMISSÃO",
        }
        f_resp_acao = resp_acoes[st.selectbox("Responsável pela ação", list(resp_acoes.keys()))]
    with col_f5:
        etapas  = ["Todas"] + sorted({r["etapa"] for r in registros})
        f_etapa = st.selectbox("Etapa atual", etapas)

    filtrados = registros_filtrados
    if busca:
        bl = busca.lower()
        filtrados = [r for r in filtrados if bl in (r["nome"] + r["cargo"] + r["empresa"]).lower()]
    if f_resp      != "Todos":  filtrados = [r for r in filtrados if r["resp_admissao"] == f_resp]
    if f_prior     != "Todas":  filtrados = [r for r in filtrados if r["prioridade"]    == f_prior]
    if f_resp_acao:             filtrados = [r for r in filtrados if r["resp_acao"]      == f_resp_acao]
    if f_etapa     != "Todas":  filtrados = [r for r in filtrados if r["etapa"]          == f_etapa]

    st.caption(f"Exibindo {len(filtrados)} de {len(registros)} registros")

    if not filtrados:
        st.info("Nenhum registro encontrado com os filtros selecionados.")
    else:
        df_show = pd.DataFrame([{
            "Prioridade":      r["prioridade"],
            "Candidato":       r["nome"],
            "Cargo":           r["cargo"],
            "Etapa":           r["etapa"],
            "Ação necessária": r["acao"],
            "Resp. ação":      r["resp_acao"],
            "Resp. admissão":  r["resp_admissao"],
            "Admissão":        r["data_admissao"],
            "Dias p/ adm.":    r["dias_admissao"],
            "Tempo parado":    f"{r['dias_parado']}d" if r["dias_parado"] is not None else "—",
            "Validação":       r["status_validacao"],
            "Benefícios":      r["status_beneficios"],
            "Observações":     r["obs"],
        } for r in filtrados])

        st.dataframe(
            df_show,
            use_container_width=True,
            hide_index=True,
            height=500,
            column_config={
                "Dias p/ adm.": st.column_config.NumberColumn(format="%d dias"),
                "Prioridade":   st.column_config.TextColumn(width="small"),
                "Resp. ação":   st.column_config.TextColumn(width="small"),
            }
        )