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
DIAS_ALERTA    = 2

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


# ── Responsabilidade: CANDIDATO ──────────────
_STATUS_CANDIDATO_NORM = {
    _strip_acentos(s) for s in {
        "10 - Contrato de trabalho enviado",
        "11 - Assinatura do contrato pendente",
        "16 - Documentos pendentes",
        "17 - Exame realizado. Documentos pendentes",
    }
}

# ── Responsabilidade: BENEFÍCIOS ─────────────
_STATUS_BENEFICIOS_NORM = {
    _strip_acentos(s) for s in {
        "20 - Aguardando finalização da jornada de benefícios",
    }
}

# ── Responsabilidade: EXAME ADMISSIONAL ──────
_STATUS_EXAME_NORM = {
    _strip_acentos(s) for s in {
        "3 - Agendamento do exame solicitado à 3778",
        "18 - Documentação ok . Pendente ASO/enquad.PCD",
    }
}

# ── Responsabilidade: ADMISSÃO ───────────────
_STATUS_ADMISSAO_NORM = {
    _strip_acentos(s) for s in {
        "1 - Contato realizado",
        "2 - Candidato não retornou aos contatos",
        "4 - Documentação e ASO enviados",
        "5 - Admissão convertida no HCM",
        "7 - Conferência de cadastro no RM",
        "9 - Admissão finalizada - Gestor/Candidato Comun.",
        "12 - Contrato de trabalho assinado",
        "14 - Admissão finalizada",
        "19 - Data admissão alterada|Gestor e recrutador cientes",
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


def _classificar(s_adm, s_val, s_proc, s_ben, s_importacao=""):
    """Fluxo sequencial de classificação por responsável.
    Ordem: Validacao -> Importacao -> Status do processo -> Beneficios (filtro final)
    Após passar pela importação, qualquer jornada pendente = BENEFÍCIOS.
    """
    s_val_n  = _strip_acentos(s_val)
    s_proc_n = _strip_acentos(s_proc)
    s_imp_n  = _strip_acentos(s_importacao)
    # DOCENTE = tipo de jornada especial, não tem obrigação de concluir benefícios
    ben_pendente = (s_ben.upper() in ("NÃO INICIADO", "NAO INICIADO", "EM ANDAMENTO")
                    and s_ben.upper() != "DOCENTE")

    # ── PASSO 2: Validação ────────────────────
    if not s_val or s_val in ("", "nan"):
        return "nao_iniciado", "sem_validacao", "ADMISSÃO"

    if "PENDENTE" in s_val_n and "REMUNER" in s_val_n:
        return "validacao_remuneracao", s_val, "REMUNERAÇÃO"

    if "PENDENTE" in s_val_n or "PREENCHIMENTO" in s_val_n:
        return "validacao_rs", s_val, "RECRUTAMENTO"

    # ── PASSO 3: Importação ───────────────────
    if s_importacao == "0" or s_imp_n == "0":
        return "importacao", s_importacao, "IMPORTAÇÃO"

    if "ERRO" in s_imp_n:
        # FIX CANCELADO: validação cancelada com erro de importação
        # → não deve ir para Remuneração, apenas sinalizar no Kanban como pendente de encerramento
        if "CANCELAD" in s_val_n:
            return "cancelado_pendente", s_val, "ADMISSÃO"
        # FIX: se remuneração já foi validada, o erro é operacional → responsabilidade IMPORTAÇÃO
        if "VALIDADO" in s_val_n:
            return "importacao", s_importacao, "IMPORTAÇÃO"
        return "erro_importacao", s_importacao, "REMUNERAÇÃO"

    # Importado ou Duplo Vínculo → candidato recebeu o link, prossegue

    # ── PASSO 4: Status do processo ───────────
    if "15 - ADMISSAO CANCELADA" in s_proc_n:
        return "excluido", s_proc, "—"

    # Exame admissional: pendência operacional com clínica/3778
    if s_proc_n in _STATUS_EXAME_NORM:
        return "exame", s_proc, "EXAME ADMISSIONAL"

    # Candidato: docs ou assinatura pendentes — pendência operacional
    if s_proc_n in _STATUS_CANDIDATO_NORM:
        return "candidato", s_proc, "CANDIDATO"

    # Status 20: jornada explícita no processo
    if s_proc_n in _STATUS_BENEFICIOS_NORM:
        return "beneficios", s_proc, "BENEFÍCIOS"

    # ── PASSO 5: Filtro final de Benefícios ───
    # Após importação confirmada, jornada pendente = responsabilidade do candidato

    # Statuses de Admissão → verifica jornada
    if s_proc_n in _STATUS_ADMISSAO_NORM:
        if ben_pendente:
            return "beneficios", s_proc, "BENEFÍCIOS"
        return "admissao", s_proc, "ADMISSÃO"

    # Status vazio/zero após importação → verifica jornada diretamente
    if s_proc in ("", "0", "0.0"):
        if ben_pendente:
            return "beneficios", s_proc, "BENEFÍCIOS"
        return "admissao", s_proc, "ADMISSÃO"

    # Qualquer outro status com jornada pendente
    if ben_pendente:
        return "beneficios", s_proc, "BENEFÍCIOS"

    return "andamento", s_proc, "ADMISSÃO"


def _ben_pendente(ben: str) -> bool:
    """Retorna True se a jornada de benefícios ainda está pendente."""
    return (
        ben.upper() in ("NÃO INICIADO", "NAO INICIADO", "EM ANDAMENTO")
        and ben.upper() != "DOCENTE"
    )


def _derivar_etapa(s, sv, ben, s_imp="") -> str:
    """Logica deterministica de etapa - nenhum status fica sem categoria.
    Prioridade: Validacao > Importacao > Status > Jornada > Default
    """
    sl   = s.lower()
    svl  = sv.lower()
    impl = s_imp.lower()

    # Passo 1: Validacao pendente
    if "pendente - remuner" in svl:
        return "Validação · Remun."
    if "pendente - r&s" in svl or "preenchimento" in svl:
        return "Validação · R&S"

    # Passo 2: Importacao bloqueada
    if s_imp == "0" or impl == "0" or "erro" in impl:
        # FIX CANCELADO: validação cancelada com erro de importação
        # → exibe coluna própria no Kanban, não polui "Importação"
        if "cancelad" in svl:
            return "Cancelado"
        return "Importação"

    # Passo 3: Status do processo
    # Contato
    if "1 - contato" in sl or "2 - candidato" in sl:
        return "Contato"

    # Documentos (status 16 e 17 — 17 prioriza docs sobre exame)
    if "16 - documentos pendentes" in sl or "17 - exame realizado" in sl:
        return "Documentos"

    # Exame (agendamento = status 3, pendente ASO = status 18)
    if "3 - agendamento" in sl or "18 - documentação ok" in sl:
        return "Exame"

    # Contrato
    if "10 - contrato de trabalho enviado" in sl or "11 - assinatura" in sl:
        return "Contrato"

    # Jor. Sydle — status 20 explícito
    if "20 - aguardando" in sl:
        return "Jor. Sydle"

    # Encerramento — status administrativos que indicam processo concluído
    if any(x in sl for x in [
        "4 - documentação e aso",
        "5 - admissão convertida",
        "9 - admissão finalizada",
        "12 - contrato de trabalho assinado",
        "14 - admissão finalizada",
        "19 - data admissão"
    ]):
        if _ben_pendente(ben):
            return "Jor. Sydle"
        return "Encerramento"

    # Passo 4: Status vazio - usa jornada
    if not sl:
        if _ben_pendente(ben):
            return "Jor. Sydle"
        return "Não iniciado"

    # Passo 5: Default
    if _ben_pendente(ben):
        return "Jor. Sydle"

    return "Não iniciado"


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
        s_adm         = _sv(row, "Status Admissão")
        data_inclusao = _sd(row, "Data Inclusão Candidato")
        if data_inclusao is None:
            return False
        nome = _sv(row, "Nome Candidato")
        id_  = _sv(row, "ID")
        tem_identificador = bool(nome and nome not in ("", "nan")) or bool(id_ and id_ not in ("", "nan"))
        if not tem_identificador:
            return False
        s_proc_raw = _sv(row, "Status")
        if s_proc_raw.startswith("2 -"):
            return True
        return s_adm in ("", "nan") or s_adm == "Em processo de admissão"

    df_ativos = df[df.apply(deve_incluir, axis=1)].copy()

    registros  = []
    contadores = {k: 0 for k in ("nao_iniciado", "validacao_remuneracao", "validacao_rs",
                                  "candidato", "admissao", "exame", "importacao",
                                  "erro_importacao", "beneficios", "excluido", "andamento",
                                  "cancelado_pendente")}
    por_resp   = {"CANDIDATO": 0, "ADMISSÃO": 0, "RECRUTAMENTO": 0,
                  "REMUNERAÇÃO": 0, "EXAME ADMISSIONAL": 0, "IMPORTAÇÃO": 0, "BENEFÍCIOS": 0}
    urgentes_total = sem_resp_total = 0

    for _, row in df_ativos.iterrows():
        s_adm      = _sv(row, "Status Admissão")
        s_val      = _sv(row, "Status de Validação")
        s_proc_raw = _sv(row, "Status")
        s_proc     = "" if s_proc_raw in ("0", "0.0") else s_proc_raw
        s_ben      = _sv(row, "Jornada Benefícios - Sydle")

        data_adm      = _sd(row, "Data de Admissão Prevista - Time de Admissão") or _sd(row, "Data de Admissão")
        data_inclusao = _sd(row, "Data Inclusão Candidato")

        dias_ate    = (data_adm - hoje).days if data_adm else None
        urgente     = dias_ate is not None and 0 <= dias_ate <= dias_alerta
        s_jornada    = _sv(row, "Status Jornada")
        s_importacao = _sv(row, "Status Importação")
        s_cod_posicao = _sv(row, "Código da Posição")
        s_obs_imp     = _sv(row, "OBS Importação")
        tipo_pend, _, resp_acao = _classificar(s_adm, s_val, s_proc, s_ben, s_importacao)
        resp_adm    = _nome_responsavel(_sv(row, "Responsável pela Admissão"))
        dias_parado = (hoje - data_inclusao).days if data_inclusao else None
        prior, pord = _prioridade(dias_ate)

        if urgente:                     urgentes_total += 1
        if resp_adm == "Não atribuído": sem_resp_total += 1
        contadores[tipo_pend] = contadores.get(tipo_pend, 0) + 1
        if resp_acao in por_resp:
            por_resp[resp_acao] += 1

        # Label especial para cancelados pendentes de encerramento
        if tipo_pend == "cancelado_pendente":
            label_st = "⚠️ Cancelado na validação — pendente de encerramento nos demais sistemas"
        else:
            label_st = _LABEL_STATUS.get(s_proc, s_proc or "Processo não iniciado")

        registros.append({
            "nome":             _sv(row, "Nome Candidato") or f"⚠️ Sem nome ({_sv(row, 'ID') or 'ID desconhecido'})",
            "cargo":            _sv(row, "Nome do Cargo"),
            "tipo":             _sv(row, "Tipo de Educador"),
            "empresa":          _sv(row, "Empresa / Unidade de Negócio"),
            "local":            _sv(row, "Local"),
            "marca":            _sv(row, "Marca"),
            "filial":           _sv(row, "Filial"),
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
            "etapa":            _derivar_etapa(s_proc, s_val, s_ben, s_importacao),
            "acao":             _derivar_acao(s_proc, s_val, s_ben),
            "prioridade":       prior,
            "prioridade_ord":   pord,
            "label_status":     label_st,
            "obs":              _sv(row, "Observações - Admissão"),
            "motivo_remuneracao": (
                s_obs_imp if "ERRO" in _strip_acentos(s_importacao)
                else "Pendência de posição"
            ),
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

_RESP_ORDER = ["RECRUTAMENTO", "REMUNERACAO", "IMPORTACAO", "CANDIDATO", "EXAME", "BENEFICIOS", "ADMISSAO"]
_RESP_META  = {
    "RECRUTAMENTO": {"label": "Recrutamento",     "cor": "#7F77DD", "key": "RECRUTAMENTO"},
    "REMUNERACAO":  {"label": "Remuneração",       "cor": "#BA7517", "key": "REMUNERAÇÃO"},
    "IMPORTACAO":   {"label": "Importação",        "cor": "#C4320A", "key": "IMPORTAÇÃO"},
    "CANDIDATO":    {"label": "Candidato",         "cor": "#185FA5", "key": "CANDIDATO"},
    "EXAME":        {"label": "Exame Adm.", "cor": "#A32D2D", "key": "EXAME ADMISSIONAL"},
    "BENEFICIOS":   {"label": "Jor. Sydle",        "cor": "#059669", "key": "BENEFÍCIOS"},
    "ADMISSAO":     {"label": "Admissão",          "cor": "#0F6E56", "key": "ADMISSÃO"},
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
    st.markdown(hero_card("Urgentes", urg, "admissão em ≤2 dias",
        "requer atenção imediata", "#FCEBEB", "#A32D2D", "#E24B4A"), unsafe_allow_html=True)
with h3:
    st.markdown(hero_card("Não iniciados", ni, "sem processo em andamento",
        "processo travado", "#FAEEDA", "#854F0B", "#BA7517"), unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ── Linha 2: Cards por responsável ───────────
r_cols = st.columns(7)
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
aba_gargalos, aba_analise, aba_kanban, aba_tempo, aba_tabela = st.tabs([
    "📊 Gargalos por Responsável",
    "📈 Análise",
    "🗂 Kanban por Etapa",
    "⏳ Maior Tempo em Aberto",
    "☰ Tabela detalhada",
])


# ─── GARGALOS ─────────────────────────────────
with aba_gargalos:
    st.markdown("### Onde estão os processos parados?")

    col_a, col_b, col_c, col_d, col_e, col_f, col_g = st.columns(7)

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

    card_resp_gargalo(col_a, "✏️", "Recrutamento",      pr.get("RECRUTAMENTO", 0),      "#7F77DD", "validação R&S")
    card_resp_gargalo(col_b, "💰", "Remuneração",       pr.get("REMUNERAÇÃO", 0),       "#BA7517", "validação salarial")
    card_resp_gargalo(col_c, "📥", "Importação",        pr.get("IMPORTAÇÃO", 0),        "#C4320A", "cadastro no sistema")
    card_resp_gargalo(col_d, "👤", "Candidato",         pr.get("CANDIDATO", 0),         "#185FA5", "ação do candidato")
    card_resp_gargalo(col_e, "🔬", "Exame Adm.", pr.get("EXAME ADMISSIONAL", 0), "#A32D2D", "clínica / 3778")
    card_resp_gargalo(col_f, "🎁", "Jor. Sydle",        pr.get("BENEFÍCIOS", 0),        "#059669", "jornada Sydle")
    card_resp_gargalo(col_g, "🏢", "Admissão",          pr.get("ADMISSÃO", 0),          "#0F6E56", "ação interna")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    st.markdown("---")

    # Calendário de admissões previstas
    st.markdown("### 📅 Admissões previstas")

    from datetime import timedelta
    from calendar import monthrange

    hoje_dt = date.today()
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

    meses_necessarios = {(hoje_dt.year, hoje_dt.month)}
    for dt in admissoes_por_dia:
        meses_necessarios.add((dt.year, dt.month))
    meses_ord = sorted(meses_necessarios)[:3]

    nomes_mes = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    dias_sem  = ["S","T","Q","Q","S","S","D"]

    def cor_cal(n):
        if n == 0:  return "#f0f2f5", "#98a2b3"
        if n <= 2:  return "#dbeafe", "#1d4ed8"
        if n <= 6:  return "#93c5fd", "#1e3a8a"
        return "#2563eb", "white"

    cal_parts = [
        '<!DOCTYPE html><html><head><meta charset="UTF-8">',
        '<style>',
        '*{box-sizing:border-box;margin:0;padding:0;font-family:Segoe UI,Arial,sans-serif;}',
        'body{background:white;padding:10px 14px;border:1.5px solid #e4e7ec;border-radius:10px;}',
        '.wrap{display:flex;gap:20px;flex-wrap:wrap;}',
        '.mes{flex:1;min-width:160px;}',
        '.titulo{font-size:11px;font-weight:600;color:#344054;margin-bottom:6px;}',
        '.grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;}',
        '.hdr{font-size:8px;font-weight:600;color:#98a2b3;text-align:center;padding:1px 0;}',
        '.d{aspect-ratio:1;border-radius:4px;display:flex;flex-direction:column;',
        '   align-items:center;justify-content:center;font-size:9px;font-weight:500;}',
        '.d.vz{background:transparent;}',
        '.d .n{line-height:1;}',
        '.d .c{font-size:7px;font-weight:700;line-height:1;}',
        '.d.hj{outline:2px solid #2563eb;outline-offset:1px;border-radius:4px;}',
        '.leg{display:flex;gap:8px;margin-top:8px;align-items:center;font-size:9px;color:#98a2b3;flex-wrap:wrap;}',
        '.sq{width:8px;height:8px;border-radius:2px;flex-shrink:0;}',
        '</style></head><body>',
        '<div class="wrap">',
    ]

    for ano, mes in meses_ord:
        _, dias_no_mes = monthrange(ano, mes)
        primeiro_dia   = date(ano, mes, 1)
        offset         = primeiro_dia.weekday()
        cal_parts.append(f'<div class="mes"><div class="titulo">{nomes_mes[mes-1]} {ano}</div><div class="grid">')
        for d in dias_sem:
            cal_parts.append(f'<div class="hdr">{d}</div>')
        for _ in range(offset):
            cal_parts.append('<div class="d vz"></div>')
        for dia in range(1, dias_no_mes + 1):
            dt     = date(ano, mes, dia)
            n      = admissoes_por_dia.get(dt, 0)
            bg, fc = cor_cal(n)
            extra  = ' hj' if dt == hoje_dt else ''
            cal_parts.append(
                f'<div class="d{extra}" style="background:{bg};color:{fc};">'
                f'<div class="n">{dia}</div>'
                + (f'<div class="c">{n}</div>' if n > 0 else '')
                + '</div>'
            )
        cal_parts.append('</div></div>')

    cal_parts += [
        '</div>',
        '<div class="leg">',
        '<span>Volume:</span>',
        '<span class="sq" style="background:#dbeafe;border:0.5px solid #93c5fd;"></span><span>1–2</span>',
        '<span class="sq" style="background:#93c5fd;"></span><span>3–6</span>',
        '<span class="sq" style="background:#2563eb;"></span><span>7+</span>',
        '</div></body></html>',
    ]

    n_meses = len(meses_ord)
    components.html("".join(cal_parts), height=n_meses * 185 + 50, scrolling=False)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # ── Kanban por responsável ────────────────
    st.markdown("Visão dos processos agrupados pelo responsável pela próxima ação.")
    dados_json_resp = safe_json({"registros": registros_filtrados})

    kanban_resp_html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<style>
:root {{
  --surface:#fff;--surface2:#f7f8fa;--border:#e4e7ec;--border2:#d0d5dd;
  --text:#101828;--text2:#475467;--text3:#98a2b3;
  --red:#d92d20;--red-bg:#fff1f0;--red-border:#fda29b;
  --orange:#c4320a;--orange-bg:#fff4ed;--orange-border:#f9c6a8;
  --yellow:#b54708;--yellow-bg:#fffaeb;--yellow-border:#fec84b;
  --purple:#6941c6;--purple-bg:#f4f3ff;--purple-border:#c3b5fd;
  --gray:#344054;--gray-bg:#f9fafb;--gray-border:#d0d5dd;
}}
*{{box-sizing:border-box;margin:0;padding:0;font-family:Segoe UI,Arial,sans-serif;}}
body{{background:#f0f2f5;font-size:13px;padding:8px;}}
.board{{display:flex;gap:12px;overflow-x:auto;padding-bottom:8px;width:100%;}}
.col{{flex:1 1 0;min-width:180px;background:#fff;border:1px solid var(--border);border-radius:10px;overflow:hidden;display:flex;flex-direction:column;}}
.col-header{{padding:12px 14px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px;background:var(--surface2);}}
.col-title{{font-size:12px;font-weight:700;flex:1;}}
.col-count{{font-size:10px;font-weight:600;background:#fff;border:1px solid var(--border2);border-radius:10px;padding:1px 8px;color:var(--text2);}}
.col-body{{padding:8px;overflow-y:visible;flex:1;}}
.card{{background:var(--surface2);border:1px solid var(--border);border-radius:7px;padding:8px 10px;margin-bottom:6px;}}
.card.urg{{border-left:3px solid var(--red);}}
.card.alta{{border-left:3px solid var(--orange);}}
.card.med{{border-left:3px solid var(--yellow);}}
.card-nome{{font-size:11px;font-weight:600;margin-bottom:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.card-cargo{{font-size:10px;color:var(--text3);margin-bottom:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.card-acao{{font-size:10px;color:var(--text2);margin-bottom:4px;}}
.card-foot{{display:flex;gap:4px;flex-wrap:wrap;}}
.badge{{display:inline-block;border-radius:5px;padding:2px 6px;font-size:9px;font-weight:600;border:1px solid transparent;}}
.b-etapa{{background:var(--purple-bg);color:var(--purple);border-color:var(--purple-border);}}
.col-rec .col-header{{border-top:4px solid #7F77DD;}}
.col-rem .col-header{{border-top:4px solid #BA7517;}}
.col-imp .col-header{{border-top:4px solid #C4320A;}}
.col-can .col-header{{border-top:4px solid #185FA5;}}
.col-exa .col-header{{border-top:4px solid #A32D2D;}}
.col-ben2 .col-header{{border-top:4px solid #059669;}}
.col-adm .col-header{{border-top:4px solid #0F6E56;}}
.mais{{text-align:center;padding:6px;font-size:10px;color:var(--text3);border-top:1px solid var(--border);}}
.urg-badge{{font-size:9px;background:#fef2f2;color:#dc2626;border-radius:10px;padding:1px 6px;font-weight:600;}}
</style></head><body>
<div class="board" id="board"></div>
<script>
const DADOS = {dados_json_resp};
const registros = DADOS.registros;
const PRIOR_CFG = {{
  'Vencida':{{label:'⚠️ Vencida',cls:'b-vencida',urg:'urg'}},
  'Crítica':{{label:'🔴 Crítica',cls:'b-critica',urg:'urg'}},
  'Alta':{{label:'🟠 Alta',cls:'b-alta',urg:'alta'}},
  'Média':{{label:'🟡 Média',cls:'b-media',urg:'med'}},
  'Normal':{{label:'⚪ Normal',cls:'b-normal',urg:''}},
}};
const COLUNAS = [
  {{resp:'RECRUTAMENTO',cls:'col-rec',icon:'✏️',label:'Recrutamento',sub:'validação R&S'}},
  {{resp:'REMUNERAÇÃO',cls:'col-rem',icon:'💰',label:'Remuneração',sub:'validação salarial'}},
  {{resp:'IMPORTAÇÃO',cls:'col-imp',icon:'📥',label:'Importação',sub:'cadastro no sistema'}},
  {{resp:'CANDIDATO',cls:'col-can',icon:'👤',label:'Candidato',sub:'ação do candidato'}},
  {{resp:'EXAME ADMISSIONAL',cls:'col-exa',icon:'🔬',label:'Exame Adm.',sub:'clínica / 3778'}},
  {{resp:'BENEFÍCIOS',cls:'col-ben2',icon:'🎁',label:'Jor. Sydle',sub:'jornada Sydle'}},
  {{resp:'ADMISSÃO',cls:'col-adm',icon:'🏢',label:'Admissão',sub:'ação interna'}},
];
function esc(s){{return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}}
const grupos={{}};
registros.forEach(r=>{{if(!grupos[r.resp_acao])grupos[r.resp_acao]=[];grupos[r.resp_acao].push(r);}});
Object.values(grupos).forEach(l=>l.sort((a,b)=>a.prioridade_ord-b.prioridade_ord||(a.dias_admissao||9999)-(b.dias_admissao||9999)));
const MAX=30;
document.getElementById('board').innerHTML=COLUNAS.filter(c=>grupos[c.resp]?.length>0).map(col=>{{
  const lista=grupos[col.resp]||[];
  const urg=lista.filter(r=>['Vencida','Crítica','Alta'].includes(r.prioridade)).length;
  const cards=lista.slice(0,MAX).map(r=>{{
    const pc=PRIOR_CFG[r.prioridade]||PRIOR_CFG['Normal'];
    const adm=r.dias_admissao!==null?(r.dias_admissao<0?`⚠️ ${{Math.abs(r.dias_admissao)}}d atraso`:r.dias_admissao===0?'🔴 HOJE':r.urgente?`🟡 ${{r.dias_admissao}}d`:r.data_admissao):r.data_admissao;
    const diasAberto=r.dias_parado!==null&&r.dias_parado!==undefined?`📅 ${{r.dias_parado}} dias em aberto`:'';
    const motivoRem=col.resp==='REMUNERAÇÃO'&&r.motivo_remuneracao?`⚠️ ${{esc(r.motivo_remuneracao)}}`:'';
    return `<div class="card ${{pc.urg}}">
      <div class="card-nome">${{esc(r.nome)}}</div>
      <div class="card-cargo">${{esc(r.cargo)}}</div>
      ${{motivoRem?`<div style="font-size:10px;color:#BA7517;font-weight:600;margin-bottom:4px;">${{motivoRem}}</div>`:''}}
      <div style="font-size:10px;color:#98a2b3;margin-top:4px;">${{diasAberto}}</div>
      <div style="font-size:10px;color:#98a2b3;">${{esc(adm)}} · ${{esc(r.resp_admissao)}}</div>
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
function autoResize(){{
  const h=document.body.scrollHeight;
  window.parent.postMessage({{type:'streamlit:setFrameHeight',height:h}},'*');
}}
window.addEventListener('load',autoResize);
setTimeout(autoResize,300);
</script></body></html>"""

    components.html(kanban_resp_html, height=700, scrolling=True)


# ─── ANÁLISE ──────────────────────────────────
with aba_analise:

    # Tempo médio total
    dias_validos = [r["dias_parado"] for r in registros if r["dias_parado"] is not None]
    media_total  = round(sum(dias_validos) / len(dias_validos)) if dias_validos else 0
    cor_media = "#A32D2D" if media_total > 30 else "#BA7517" if media_total > 14 else "#0F6E56"
    bg_media  = "#FCEBEB" if media_total > 30 else "#FAEEDA" if media_total > 14 else "#E1F5EE"
    txt_media = "acima do ideal" if media_total > 30 else "atencao recomendada" if media_total > 14 else "dentro do esperado"

    st.markdown("### ⏱ Tempo médio do processo")
    st.markdown(
        '<div style="background:white;border:1.5px solid #e4e7ec;border-top:3px solid ' + cor_media + ';'
        'border-radius:10px;padding:1.1rem 1.4rem;display:flex;align-items:center;gap:20px;">'
        '<div style="font-size:56px;font-weight:500;line-height:1;color:' + cor_media + ';">' + str(media_total) + '</div>'
        '<div><div style="font-size:11px;color:#475467;font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;">dias em média</div>'
        '<div style="font-size:12px;color:#98a2b3;">do início do processo até hoje</div>'
        '<span style="display:inline-block;margin-top:6px;font-size:10px;font-weight:600;padding:2px 9px;'
        'border-radius:20px;background:' + bg_media + ';color:' + cor_media + ';">' + txt_media + '</span>'
        '</div></div>', unsafe_allow_html=True
    )

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # Gargalo por responsável (barras %)
    st.markdown("### 📊 Gargalo por responsável")
    _max_r2 = max(pr.values(), default=1)
    _total_pr = max(sum(pr.values()), 1)
    _garg_rows2 = "".join(
        f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">' +
        f'<div style="width:110px;font-size:12px;color:#475467;flex-shrink:0;">{_RESP_META[rk]["label"]}</div>' +
        f'<div style="flex:1;height:10px;background:#e4e7ec;border-radius:5px;overflow:hidden;">' +
        f'<div style="height:10px;width:{round((pr.get(_RESP_META[rk]["key"],0)/_max_r2)*100) if _max_r2 else 0}%;background:{_RESP_META[rk]["cor"]};border-radius:5px;"></div></div>' +
        f'<div style="width:36px;font-size:12px;font-weight:600;color:#101828;text-align:right;">{round((pr.get(_RESP_META[rk]["key"],0)/_total_pr)*100)}%</div></div>'
        for rk in _RESP_ORDER
    )
    st.markdown(
        f'<div style="background:white;border:1.5px solid #e4e7ec;border-radius:10px;padding:1.1rem 1.4rem;">' +
        _garg_rows2 + '</div>', unsafe_allow_html=True
    )

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # ── Mapa dinâmico por filial ──
    st.markdown("### 🏢 Processos por filial")
    filial_count = {}
    for r in registros:
        m = r["filial"] or "Sem filial"
        filial_count[m] = filial_count.get(m, 0) + 1
    filial_ord = sorted(filial_count.items(), key=lambda x: -x[1])

    import json as _json
    filial_list = [{"filial": m, "count": c} for m, c in filial_ord]
    filial_json = _json.dumps(filial_list, ensure_ascii=False)

    mapa_html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:Segoe UI,Arial,sans-serif;}}
body{{background:white;padding:12px 16px;border:1.5px solid #e4e7ec;border-radius:10px;}}
.controls{{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center;}}
.ctrl-label{{font-size:11px;color:#475467;font-weight:600;}}
input[type=text]{{border:1px solid #d0d5dd;border-radius:6px;padding:4px 10px;font-size:11px;width:200px;outline:none;}}
input[type=text]:focus{{border-color:#185FA5;box-shadow:0 0 0 2px #dbeafe;}}
select{{border:1px solid #d0d5dd;border-radius:6px;padding:4px 8px;font-size:11px;background:white;outline:none;cursor:pointer;}}
select:focus{{border-color:#185FA5;}}
.badge-total{{font-size:10px;color:#475467;background:#f0f2f5;border-radius:10px;padding:2px 8px;}}
.chart-wrap{{position:relative;}}
.no-results{{text-align:center;padding:40px;color:#98a2b3;font-size:12px;}}
</style></head><body>
<div class="controls">
  <span class="ctrl-label">🔍</span>
  <input type="text" id="busca" placeholder="Filtrar filial..." oninput="filtrar()">
  <span class="ctrl-label">Ordenar:</span>
  <select id="ordem" onchange="filtrar()">
    <option value="desc">Maior → Menor</option>
    <option value="asc">Menor → Maior</option>
    <option value="az">A → Z</option>
    <option value="za">Z → A</option>
  </select>
  <span class="ctrl-label">Exibir:</span>
  <select id="limite" onchange="filtrar()">
    <option value="15">Top 15</option>
    <option value="25">Top 25</option>
    <option value="999">Todas</option>
  </select>
  <span class="badge-total" id="badge">0 filiais</span>
</div>
<div class="chart-wrap" id="wrap"><canvas id="chart"></canvas><div class="no-results" id="nores" style="display:none">Nenhuma filial encontrada</div></div>
<script>
const DADOS = {filial_json};
let chartInst = null;
const PALETA = [
  '#185FA5','#2563eb','#3b82f6','#60a5fa','#0F6E56','#059669','#10b981','#34d399',
  '#7F77DD','#6941c6','#8b5cf6','#BA7517','#d97706','#f59e0b','#C4320A','#dc2626',
  '#0891b2','#0e7490','#4338ca','#7c3aed'
];
function filtrar() {{
  const busca = document.getElementById('busca').value.toLowerCase();
  const ordem = document.getElementById('ordem').value;
  const limite = parseInt(document.getElementById('limite').value);
  let dados = DADOS.filter(d => d.filial && d.filial !== 'Sem filial' ? d.filial.toLowerCase().includes(busca) : busca === '');
  dados.sort((a,b) => {{
    if(ordem==='desc') return b.count-a.count;
    if(ordem==='asc')  return a.count-b.count;
    if(ordem==='az')   return a.filial.localeCompare(b.filial,'pt');
    return b.filial.localeCompare(a.filial,'pt');
  }});
  const total = dados.length;
  dados = dados.slice(0,limite);
  document.getElementById('badge').textContent = total + (total===1?' filial':' filiais');
  const nr = document.getElementById('nores');
  if(dados.length===0){{nr.style.display='flex';nr.style.alignItems='center';nr.style.justifyContent='center';nr.style.minHeight='120px';if(chartInst){{chartInst.destroy();chartInst=null;document.getElementById('chart').style.display='none';}}return;}}
  nr.style.display='none';
  document.getElementById('chart').style.display='block';
  const labels = dados.map(d=>d.filial);
  const values = dados.map(d=>d.count);
  const colors = dados.map((_,i)=>PALETA[i%PALETA.length]);
  if(chartInst) chartInst.destroy();
  const ctx = document.getElementById('chart').getContext('2d');
  chartInst = new Chart(ctx, {{
    type: 'bar',
    data: {{labels, datasets:[{{
      data:values,
      backgroundColor: colors.map(c=>c+'cc'),
      borderColor: colors,
      borderWidth: 1,
      borderRadius: 5,
      borderSkipped: false
    }}]}},
    options: {{
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend:{{display:false}},
        tooltip:{{
          callbacks:{{
            label: ctx=>`  ${{ctx.raw}} processo${{ctx.raw!==1?'s':''}}`,
            title: ctx=>ctx[0].label
          }},
          backgroundColor:'#1e293b',
          titleFont:{{size:11}},
          bodyFont:{{size:11}},
          padding:8,
          cornerRadius:6,
        }}
      }},
      scales: {{
        x: {{
          grid:{{color:'#f0f2f5'}},
          ticks:{{font:{{size:10}},color:'#98a2b3',stepSize:1}},
          border:{{display:false}}
        }},
        y: {{
          grid:{{display:false}},
          ticks:{{font:{{size:10}},color:'#344054',maxTicksLimit:dados.length}},
          border:{{display:false}}
        }}
      }}
    }}
  }});
  const h = Math.max(200, dados.length * 24 + 50);
  document.getElementById('wrap').style.height = h + 'px';
  setTimeout(()=>{{
    const total_h = document.body.scrollHeight;
    window.parent.postMessage({{type:'streamlit:setFrameHeight',height:total_h+20}},'*');
  }},100);
}}
filtrar();
</script></body></html>"""

    altura_mapa = max(420, min(len(filial_ord), 15) * 26 + 130)
    components.html(mapa_html, height=altura_mapa, scrolling=False)


# ─── KANBAN POR ETAPA ─────────────────────────
with aba_kanban:
    st.markdown("Visão do processo por etapa — ordenado por urgência.")
    dados_kanban = dict(dados); dados_kanban["registros"] = registros_filtrados
    dados_json = safe_json(dados_kanban)

    kanban_html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<style>
:root {{
  --surface:#fff;--surface2:#f7f8fa;--border:#e4e7ec;--border2:#d0d5dd;
  --red:#d92d20;--red-bg:#fff1f0;--red-border:#fda29b;
  --orange:#c4320a;--orange-bg:#fff4ed;--orange-border:#f9c6a8;
  --yellow:#b54708;--yellow-bg:#fffaeb;--yellow-border:#fec84b;
  --blue:#1d4ed8;--blue-bg:#eff6ff;--blue-border:#93c5fd;
  --teal:#107569;--teal-bg:#f0fdf9;--teal-border:#5fe3c0;
  --purple:#6941c6;--purple-bg:#f4f3ff;--purple-border:#c3b5fd;
  --gray:#344054;--gray-bg:#f9fafb;--gray-border:#d0d5dd;
  --violet:#6b21a8;
}}
*{{box-sizing:border-box;margin:0;padding:0;font-family:Segoe UI,Arial,sans-serif;}}
body{{background:#f0f2f5;font-size:13px;padding:8px;}}
.board{{display:flex;gap:10px;overflow-x:auto;padding-bottom:8px;width:100%;}}
.col{{flex:1 1 0;min-width:160px;background:#fff;border:1px solid var(--border);border-radius:10px;overflow:hidden;display:flex;flex-direction:column;}}
.col-header{{padding:10px 12px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px;background:var(--surface2);}}
.col-title{{font-size:12px;font-weight:600;flex:1;}}
.col-count{{font-size:10px;font-weight:600;background:#fff;border:1px solid var(--border2);border-radius:10px;padding:1px 7px;color:#475467;}}
.col-body{{padding:8px;overflow-y:visible;flex:1;}}
.card{{background:var(--surface2);border:1px solid var(--border);border-radius:7px;padding:8px 10px;margin-bottom:6px;}}
.card.urg{{border-left:3px solid var(--red);}}
.card.alta{{border-left:3px solid var(--orange);}}
.card.med{{border-left:3px solid var(--yellow);}}
.card-nome{{font-size:11px;font-weight:600;margin-bottom:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.card-cargo{{font-size:10px;color:#98a2b3;margin-bottom:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.card-foot{{display:flex;gap:4px;flex-wrap:wrap;}}
.badge{{display:inline-block;border-radius:5px;padding:2px 6px;font-size:9px;font-weight:600;border:1px solid transparent;}}
.b-colab{{background:var(--blue-bg);color:var(--blue);border-color:var(--blue-border);}}
.b-time{{background:var(--teal-bg);color:var(--teal);border-color:var(--teal-border);}}
.b-rec{{background:var(--yellow-bg);color:var(--yellow);border-color:var(--yellow-border);}}
.b-ter{{background:var(--red-bg);color:var(--red);border-color:var(--red-border);}}
.b-imp{{background:var(--orange-bg);color:var(--orange);border-color:var(--orange-border);}}
.b-ben{{background:var(--teal-bg);color:var(--teal);border-color:var(--teal-border);}}
.col-ni    .col-header{{border-top:3px solid #dc2626;}}
.col-val   .col-header{{border-top:3px solid #d97706;}}
.col-doc   .col-header{{border-top:3px solid #1d4ed8;}}
.col-exame .col-header{{border-top:3px solid #0f766e;}}
.col-cont  .col-header{{border-top:3px solid #7c3aed;}}
.col-ben   .col-header{{border-top:3px solid #059669;}}
.col-and   .col-header{{border-top:3px solid #6b7280;}}
.col-imp2  .col-header{{border-top:3px solid #C4320A;}}
.col-enc   .col-header{{border-top:3px solid #0F6E56;}}
.col-can2  .col-header{{border-top:3px solid #6b7280;}}
.col-cancel .col-header{{border-top:3px solid #6b21a8;}}
.mais{{text-align:center;padding:6px;font-size:10px;color:#98a2b3;border-top:1px solid var(--border);}}
.urg-badge{{font-size:9px;background:#fef2f2;color:#dc2626;border-radius:10px;padding:1px 5px;font-weight:600;}}
.card-cancel-alert{{font-size:10px;font-weight:600;color:#6b21a8;background:#f5f3ff;
  border:1px solid #ddd6fe;border-radius:4px;padding:3px 7px;margin-bottom:4px;}}
</style></head><body>
<div class="board" id="board"></div>
<script>
const DADOS={dados_json};
const registros=DADOS.registros;
const PRIOR_CFG={{
  'Vencida':{{label:'⚠️ Vencida',cls:'b-vencida',urg:'urg'}},
  'Crítica':{{label:'🔴 Crítica',cls:'b-critica',urg:'urg'}},
  'Alta':{{label:'🟠 Alta',cls:'b-alta',urg:'alta'}},
  'Média':{{label:'🟡 Média',cls:'b-media',urg:'med'}},
  'Normal':{{label:'⚪ Normal',cls:'b-normal',urg:''}},
}};
const RESP_CFG={{
  'RECRUTAMENTO':{{label:'✏️ Recrutamento',cls:'b-rec'}},
  'REMUNERAÇÃO':{{label:'💰 Remun.',cls:'b-rec'}},
  'IMPORTAÇÃO':{{label:'📥 Importação',cls:'b-imp'}},
  'CANDIDATO':{{label:'👤 Candidato',cls:'b-colab'}},
  'EXAME ADMISSIONAL':{{label:'🔬 Exame Adm.',cls:'b-ter'}},
  'BENEFÍCIOS':{{label:'🎁 Jor. Sydle',cls:'b-ben'}},
  'ADMISSÃO':{{label:'🏢 Admissão',cls:'b-time'}},
  '—':{{label:'⚠️ Sem atrib.',cls:'b-vencida'}},
}};
const COLUNAS=[
  {{etapa:'Não iniciado',       cls:'col-ni',     icon:'⚠️'}},
  {{etapa:'Validação · R&S',    cls:'col-val',    icon:'✏️'}},
  {{etapa:'Validação · Remun.', cls:'col-val',    icon:'💰'}},
  {{etapa:'Importação',         cls:'col-imp2',   icon:'📥'}},
  {{etapa:'Contato',            cls:'col-cont',   icon:'📞'}},
  {{etapa:'Documentos',         cls:'col-doc',    icon:'📄'}},
  {{etapa:'Exame',              cls:'col-exame',  icon:'🩺'}},
  {{etapa:'Contrato',           cls:'col-doc',    icon:'📝'}},
  {{etapa:'Jor. Sydle',         cls:'col-ben',    icon:'🎁'}},
  {{etapa:'Encerramento',       cls:'col-enc',    icon:'✅'}},
  {{etapa:'Cancelado',          cls:'col-cancel', icon:'🚫'}},
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
  const isCancelado=col.etapa==='Cancelado';
  const cards=lista.slice(0,MAX).map(r=>{{
    const pc=PRIOR_CFG[r.prioridade]||PRIOR_CFG['Normal'];
    const rc=RESP_CFG[r.resp_acao]||RESP_CFG['—'];
    const adm=r.dias_admissao!==null?(r.dias_admissao<0?`⚠️ ${{Math.abs(r.dias_admissao)}}d atraso`:r.dias_admissao===0?'🔴 HOJE':r.urgente?`🟡 ${{r.dias_admissao}}d`:r.data_admissao):r.data_admissao;
    const diasAberto2=r.dias_parado!==null&&r.dias_parado!==undefined?`📅 ${{r.dias_parado}} dias em aberto`:'';
    const alertaCancelado=isCancelado
      ?`<div class="card-cancel-alert">🚫 Encerrar nos demais sistemas</div>`
      :'';
    return `<div class="card ${{pc.urg}}" title="${{esc(r.obs||'')}}">
      <div class="card-nome">${{esc(r.nome)}}</div>
      <div class="card-cargo">${{esc(r.cargo)}}</div>
      ${{alertaCancelado}}
      <div class="card-foot"><span class="badge ${{rc.cls}}">${{rc.label}}</span></div>
      <div style="font-size:10px;color:#98a2b3;margin-top:4px;">${{diasAberto2}}</div>
      <div style="font-size:10px;color:#98a2b3;">${{esc(adm)}} · ${{esc(r.resp_admissao)}}</div>
    </div>`;
  }}).join('');
  const mais=lista.length>MAX?`<div class="mais">+ ${{lista.length-MAX}} candidatos</div>`:'';
  const urgB=urg>0?`<span class="urg-badge">${{urg}} urg.</span>`:'';
  return `<div class="col ${{col.cls}}"><div class="col-header"><span style="font-size:14px;">${{col.icon}}</span><span class="col-title">${{esc(col.etapa)}}</span>${{urgB}}<span class="col-count">${{lista.length}}</span></div><div class="col-body">${{cards}}${{mais}}</div></div>`;
}}).join('');
function autoResize(){{
  const h=document.body.scrollHeight;
  window.parent.postMessage({{type:'streamlit:setFrameHeight',height:h}},'*');
}}
window.addEventListener('load',autoResize);
window.addEventListener('resize',autoResize);
setTimeout(autoResize,300);
</script></body></html>"""

    components.html(kanban_html, height=700, scrolling=True)

# ─── MAIOR TEMPO EM ABERTO ────────────────────
with aba_tempo:
    st.markdown("Candidatos com maior tempo em aberto — ordenados do mais antigo ao mais recente.")

    _CARGOS_DOCENTES = {"tutor", "professor", "mediador", "preceptor"}

    def _eh_docente(cargo: str) -> bool:
        cargo_lower = cargo.lower()
        return any(palavra in cargo_lower for palavra in _CARGOS_DOCENTES)

    col_t1, col_t2, col_t3 = st.columns([3, 2, 1])
    with col_t1:
        busca_tempo = st.text_input("🔍 Buscar candidato ou cargo", "", key="busca_tempo")
    with col_t2:
        tipo_cargo = st.radio(
            "Tipo de cargo",
            ["Todos", "Docente", "Administrativo"],
            horizontal=True,
            key="tipo_cargo_tempo",
        )
    with col_t3:
        top_n = st.selectbox("Exibir", [10, 25, 50, 100, "Todos"], key="top_tempo")

    lista_tempo = sorted(
        [r for r in registros_filtrados if r["dias_parado"] is not None],
        key=lambda r: r["dias_parado"],
        reverse=True,
    )

    if busca_tempo:
        bl = busca_tempo.lower()
        lista_tempo = [r for r in lista_tempo if bl in (r["nome"] + r["cargo"]).lower()]

    if tipo_cargo == "Docente":
        lista_tempo = [r for r in lista_tempo if _eh_docente(r["cargo"])]
    elif tipo_cargo == "Administrativo":
        lista_tempo = [r for r in lista_tempo if not _eh_docente(r["cargo"])]

    if top_n != "Todos":
        lista_tempo = lista_tempo[:int(top_n)]

    if not lista_tempo:
        st.info("Nenhum candidato encontrado com os filtros selecionados.")
    else:
        max_dias = lista_tempo[0]["dias_parado"] or 1

        def _cor_dias(d):
            if d is None: return "#B4B2A9", "#5F5E5A"
            if d > 30:    return "#FCEBEB", "#A32D2D"
            if d > 14:    return "#FAEEDA", "#854F0B"
            if d > 7:     return "#FFF9E6", "#BA7517"
            return "#E1F5EE", "#0F6E56"

        cards_html = []
        for i, r in enumerate(lista_tempo, 1):
            dias = r["dias_parado"]
            bg_dias, txt_dias = _cor_dias(dias)
            barra_pct = round((dias / max_dias) * 100) if max_dias else 0
            adm = r["data_admissao"] if r["data_admissao"] != "—" else "Data não informada"
            obs = r["obs"] or "—"
            resp_label = r["resp_acao"] or "—"

            cards_html.append(f"""
<div style="background:white;border:1px solid #e4e7ec;border-radius:10px;
            padding:14px 16px;margin-bottom:8px;display:flex;gap:14px;align-items:flex-start;">
  <div style="font-size:13px;font-weight:500;color:#98a2b3;min-width:28px;
              padding-top:2px;text-align:right;">#{i}</div>
  <div style="flex:1;min-width:0;">
    <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:3px;">
      <span style="font-size:14px;font-weight:500;color:#101828;
                   white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:260px;">
        {r['nome']}
      </span>
      <span style="font-size:11px;color:#475467;">{r['cargo'] or '—'}</span>
    </div>
    <div style="display:flex;gap:12px;flex-wrap:wrap;font-size:11px;color:#98a2b3;margin-bottom:8px;">
      <span>📅 Admissão prevista: <strong style="color:#344054;">{adm}</strong></span>
      <span>👤 Resp.: <strong style="color:#344054;">{r['resp_admissao']}</strong></span>
      <span>🔖 Etapa: <strong style="color:#344054;">{r['etapa']}</strong></span>
      <span>➡️ Ação: <strong style="color:#344054;">{resp_label}</strong></span>
    </div>
    <div style="height:4px;background:#e4e7ec;border-radius:2px;overflow:hidden;margin-bottom:8px;">
      <div style="height:4px;width:{barra_pct}%;background:{txt_dias};border-radius:2px;"></div>
    </div>
    <div style="font-size:11px;color:#475467;border-left:3px solid #e4e7ec;
                padding-left:8px;line-height:1.5;">
      <strong style="color:#344054;">Obs:</strong> {obs}
    </div>
  </div>
  <div style="flex-shrink:0;text-align:center;background:{bg_dias};border-radius:8px;
              padding:8px 12px;min-width:60px;">
    <div style="font-size:22px;font-weight:500;line-height:1;color:{txt_dias};">{dias}</div>
    <div style="font-size:10px;font-weight:600;color:{txt_dias};margin-top:2px;">dias</div>
  </div>
</div>
""")

        st.markdown("".join(cards_html), unsafe_allow_html=True)

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
            "✏️ Recrutamento":      "RECRUTAMENTO",
            "💰 Remuneração":       "REMUNERAÇÃO",
            "📥 Importação":        "IMPORTAÇÃO",
            "👤 Candidato":         "CANDIDATO",
            "🔬 Exame Adm.": "EXAME ADMISSIONAL",
            "🎁 Jor. Sydle":        "BENEFÍCIOS",
            "🏢 Admissão":          "ADMISSÃO",
        }
        f_resp_acao = resp_acoes[st.selectbox("Responsável pela ação", list(resp_acoes.keys()))]
    with col_f5:
        etapas  = ["Todas"] + sorted({r["etapa"] for r in registros})
        f_etapa = st.selectbox("Etapa atual", etapas)

    filtrados = list(registros)
    if busca:
        bl = busca.lower()
        filtrados = [r for r in filtrados if bl in (r["nome"] + r["cargo"] + r["empresa"]).lower()]
    if f_resp      != "Todos":  filtrados = [r for r in filtrados if r["resp_admissao"] == f_resp]
    if f_prior     != "Todas":  filtrados = [r for r in filtrados if r["prioridade"]    == f_prior]
    if f_resp_acao:             filtrados = [r for r in filtrados if r["resp_acao"]      == f_resp_acao]
    if f_etapa     != "Todas":  filtrados = [r for r in filtrados if r["etapa"]          == f_etapa]

    # ── Totalizador dinâmico ──────────────────
    total_geral = len(registros)
    total_filtrado = len(filtrados)
    has_filter = any([busca, f_resp != "Todos", f_prior != "Todas", f_resp_acao, f_etapa != "Todas"])

    tot_col1, tot_col2, tot_col3, tot_col4 = st.columns(4)
    tot_col1.metric("📋 Exibindo", total_filtrado, help="Registros no filtro atual")
    tot_col2.metric("📊 Total em aberto", total_geral)
    tot_col3.metric("🔴 Urgentes (filtro)", sum(1 for r in filtrados if r["urgente"]))
    tot_col4.metric("⏱ Média dias (filtro)",
        f"{round(sum(r['dias_parado'] for r in filtrados if r['dias_parado']) / max(sum(1 for r in filtrados if r['dias_parado']),1))}d"
        if filtrados else "—")

    if has_filter:
        st.caption(f"Filtro ativo — exibindo {total_filtrado} de {total_geral} registros ({round(total_filtrado/total_geral*100) if total_geral else 0}%)")
    else:
        st.caption(f"Sem filtro — todos os {total_geral} registros em aberto")

    if not filtrados:
        st.info("Nenhum registro encontrado com os filtros selecionados.")
    else:
        df_show = pd.DataFrame([{
            "Prioridade":      r["prioridade"],
            "Candidato":       r["nome"],
            "Cargo":           r["cargo"],
            "Filial":          r["filial"],
            "Etapa":           r["etapa"],
            "Ação necessária": r["acao"],
            "Resp. ação":      r["resp_acao"],
            "Resp. admissão":  r["resp_admissao"],
            "Admissão":        r["data_admissao"],
            "Dias p/ adm.":    r["dias_admissao"],
            "Dias em aberto":  f"{r['dias_parado']}d" if r["dias_parado"] is not None else "—",
            "Validação":       r["status_validacao"],
            "Jor. Sydle":      r["status_beneficios"],
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
                "Filial":       st.column_config.TextColumn(width="medium"),
            }
        )