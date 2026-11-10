"""
Tema escuro das ferramentas (mesmo visual da Central de ferramentas).

Uso (app.py): para as ferramentas listadas em TEMA_ESCURO, em vez do logo + botão de voltar:

    cabecalho_ferramenta(analise, voltar=voltar_ao_inicio)

Isso injeta o CSS escuro (fundo da página inteira, abas, campos, botões, uploader, alertas,
tabelas) e desenha a faixa do topo: título da ferramenta à esquerda, a marca «informa» centralizada
(ela é um botão: clicar volta à Central de ferramentas) e o botão «← Central de ferramentas» à direita.
Para temar outra ferramenta, basta incluí-la em TEMA_ESCURO.
"""
from __future__ import annotations

import html
from typing import Callable, Dict

import streamlit as st

# ferramenta (nome no MENU) -> (subtítulo, ícone SVG interno 24x24, cor)
TEMA_ESCURO: Dict[str, tuple] = {
    "Relatório Analítico": (
        "Confere o relatório e prepara a apresentação",
        '<path d="M6 3h9l4 4v14H6z"/><path d="M14 3v5h5"/><path d="M9 13h7M9 17h7M9 9h2"/>',
        "#34d399",
    ),
    "Análise das Perguntas": (
        "Cada pergunta do Total cruzada pelos segmentos",
        '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16"/>',
        "#34d399",
    ),
    "Padronização de Títulos": (
        "O título padrão de cada tipo de pergunta",
        '<path d="M4 7V5h16v2M12 5v14M9 19h6"/>',
        "#34d399",
    ),
    "Tratamento do Banco": (
        "Roda a Syntax no banco e confere a recodificação",
        '<path d="M4 6h16M4 12h10M4 18h7"/><path d="m15 16 2 2 4-4"/>',
        "#60a5fa",
    ),
    "Total Automático": (
        "Relatório pelo total: do SPSS ao Excel formatado",
        '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16"/>',
        "#a78bfa",
    ),
    "Perfil": (
        "Sexo, idade, renda, escolaridade e religião do banco",
        '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7"/>',
        "#34d399",
    ),
}

_CSS = """
<style>
:root{--tf-bg:#121923;--tf-bg2:#17202c;--tf-nav:#1a2331;--tf-linha:#273244;--tf-linha2:#33415a;
      --tf-txt:#eef2f8;--tf-muted:#9aa8bd;--tf-fraco:#6f7f96;--tf-lima:#b4e26b;--tf-lima-bg:rgba(164,214,86,.14)}

/* ---------- página ---------- */
html, body, .stApp, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"], section.main
  { background: var(--tf-bg) !important; color: var(--tf-txt) !important; }
.stApp::before{content:"";position:fixed;inset:0;pointer-events:none;z-index:0;
  background:radial-gradient(520px 260px at 0% 0%,rgba(164,214,86,.10),transparent 70%),
             radial-gradient(620px 320px at 100% 100%,rgba(96,165,250,.10),transparent 70%)}
[data-testid="stHeader"]{background:transparent !important}
[data-testid="stHeader"] *, [data-testid="stToolbar"] *{color:#cbd5e1 !important}
[data-testid="stDecoration"], footer{display:none !important}
[data-testid="stMainBlockContainer"], .block-container{max-width:1240px !important;padding:1.2rem 2.2rem 4rem !important}

/* ---------- textos ---------- */
.stApp p, .stApp li, .stApp label, .stApp span, .stApp summary,
[data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"] *{color:var(--tf-txt)}
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] *, .stApp small{color:var(--tf-muted) !important}
.stApp h1, .stApp h2{color:#fff !important;font-weight:800 !important;letter-spacing:-.4px}
.stApp h3{color:#fff !important;font-weight:800 !important;letter-spacing:-.2px;
  padding-left:14px;border-left:4px solid var(--tf-lima);margin-top:1.6rem !important;line-height:1.2}
.stApp strong, .stApp b{color:#fff}
.stApp code{background:rgba(255,255,255,.07) !important;color:var(--tf-lima) !important;border-radius:6px}
.stApp hr{border-color:var(--tf-linha) !important}
.stApp a{color:var(--tf-lima) !important}

/* ---------- faixa do topo ---------- */
[data-testid="stHorizontalBlock"]:has(.tf-ancora){
  align-items:center;gap:1rem;margin:0 0 1.2rem;padding:16px 26px;border:1px solid var(--tf-linha);border-radius:20px;
  background:radial-gradient(circle,rgba(255,255,255,.06) 1px,transparent 1.6px) 0 0/14px 14px,linear-gradient(180deg,#18212e,#141c27);
  box-shadow:0 12px 32px rgba(5,9,16,.35)}
.tf-topo{display:flex;align-items:center;gap:22px}
.tf-tit{display:flex;align-items:center;gap:14px}
.tf-chip{flex:none;width:46px;height:46px;border-radius:14px;display:flex;align-items:center;justify-content:center;color:#fff;
  background:linear-gradient(135deg,var(--c),color-mix(in srgb,var(--c) 55%,#000));box-shadow:0 8px 20px color-mix(in srgb,var(--c) 30%,transparent)}
.tf-chip svg{width:24px;height:24px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.tf-tit h1{margin:0 !important;padding:0 !important;font-size:27px !important;line-height:1.1 !important}
.tf-tit p{margin:3px 0 0 !important;color:var(--tf-muted) !important;font-size:15px}

/* ---------- abas ---------- */
[data-baseweb="tab-list"]{gap:8px;border-bottom:1px solid var(--tf-linha)}
button[data-baseweb="tab"]{border:0 !important;background:transparent !important;border-radius:12px 12px 0 0;padding:12px 18px;height:auto}
button[data-baseweb="tab"] *{color:var(--tf-muted) !important;font-size:16.5px !important;font-weight:700 !important}
button[data-baseweb="tab"]:hover{background:rgba(255,255,255,.04) !important}
button[data-baseweb="tab"][aria-selected="true"] *{color:#fff !important}
[data-baseweb="tab-highlight"]{background:var(--tf-lima) !important;height:3px !important;border-radius:3px}
[data-baseweb="tab-border"]{background:transparent !important}

/* ---------- campos ---------- */
[data-baseweb="input"], [data-baseweb="base-input"], [data-baseweb="textarea"], [data-baseweb="select"] > div{
  background:rgba(255,255,255,.03) !important;border:1px solid var(--tf-linha2) !important;border-radius:12px !important;
  transition:border-color .15s,box-shadow .15s}
[data-baseweb="input"] > div, [data-baseweb="base-input"] > div{background:transparent !important;border:0 !important}
[data-baseweb="input"]:focus-within, [data-baseweb="textarea"]:focus-within, [data-baseweb="select"]:focus-within > div{
  border-color:var(--tf-lima) !important;box-shadow:0 0 0 4px var(--tf-lima-bg) !important}
input, textarea{color:var(--tf-txt) !important;-webkit-text-fill-color:var(--tf-txt) !important;background:transparent !important;caret-color:var(--tf-lima)}
input::placeholder, textarea::placeholder{color:#8190a6 !important;-webkit-text-fill-color:#8190a6 !important}
[data-baseweb="select"] *{color:var(--tf-txt) !important}
[data-baseweb="select"] svg{fill:var(--tf-muted) !important}
[data-baseweb="popover"] [data-baseweb="menu"], [data-baseweb="popover"] ul{background:var(--tf-nav) !important;border:1px solid var(--tf-linha2);border-radius:12px}
[data-baseweb="popover"] li{background:transparent !important;color:var(--tf-txt) !important}
[data-baseweb="popover"] li:hover, [data-baseweb="popover"] li[aria-selected="true"]{background:rgba(164,214,86,.14) !important}
[data-baseweb="checkbox"] > span, [data-baseweb="radio"] > div:first-child{border-color:var(--tf-linha2) !important}
[data-baseweb="checkbox"] input:checked + div, [data-baseweb="checkbox"] [aria-checked="true"] > span{background:var(--tf-lima) !important;border-color:var(--tf-lima) !important}

/* ---------- botões ---------- */
.stButton > button, .stDownloadButton > button, [data-testid^="stBaseButton"]{
  background:rgba(255,255,255,.04) !important;color:var(--tf-txt) !important;border:1px solid var(--tf-linha2) !important;
  border-radius:12px !important;font-weight:700 !important;padding:.55rem 1.1rem !important;transition:border-color .15s,color .15s,transform .15s,background .15s}
.stButton > button:hover, .stDownloadButton > button:hover, [data-testid^="stBaseButton"]:hover{
  border-color:var(--tf-lima) !important;color:var(--tf-lima) !important;transform:translateY(-1px)}
.stButton > button *, .stDownloadButton > button *{color:inherit !important}
.stButton > button[kind="primary"], .stButton > button[data-testid="stBaseButton-primary"],
.stDownloadButton > button[kind="primary"], button[data-testid="stBaseButton-primary"]{
  background:var(--tf-lima) !important;color:#0e1520 !important;border-color:var(--tf-lima) !important}
.stButton > button[kind="primary"]:hover, .stButton > button[data-testid="stBaseButton-primary"]:hover,
button[data-testid="stBaseButton-primary"]:hover{background:#c6ee86 !important;color:#0e1520 !important;border-color:#c6ee86 !important}

/* ---------- marca centralizada = botão «voltar à Central» ---------- */
.st-key-tf_logo_home, .st-key-tf_logo_home [data-testid="stButton"]{display:flex;justify-content:center}
.st-key-tf_logo_home button{background:transparent !important;border:0 !important;box-shadow:none !important;min-height:0 !important;
  padding:8px 18px !important;border-radius:14px !important;transform:none !important;cursor:pointer}
.st-key-tf_logo_home button:hover{background:var(--tf-lima-bg) !important;transform:none !important}
.st-key-tf_logo_home button p{font-size:0 !important;line-height:1 !important;margin:0 !important;text-align:center}
.st-key-tf_logo_home button p::before{content:"informa";display:block;font-size:32px;font-weight:800;letter-spacing:-1.5px;line-height:1;color:#fff;transition:color .15s}
.st-key-tf_logo_home button p::after{content:"PESQUISA E DADOS";display:block;margin-top:6px;font-size:11px;font-weight:700;letter-spacing:2.6px;line-height:1;color:var(--tf-lima)}
.st-key-tf_logo_home button:hover p::before{color:var(--tf-lima)}
.st-key-tf_logo_home button:focus-visible{outline:2px solid var(--tf-lima) !important;outline-offset:2px}

/* ---------- uploader ---------- */
[data-testid="stFileUploaderDropzone"]{background:rgba(255,255,255,.025) !important;border:1.5px dashed var(--tf-linha2) !important;border-radius:16px !important;padding:1.2rem}
[data-testid="stFileUploaderDropzone"]:hover{border-color:var(--tf-lima) !important;background:var(--tf-lima-bg) !important}
[data-testid="stFileUploaderDropzone"] *{color:var(--tf-muted) !important}
[data-testid="stFileUploaderFile"] *{color:var(--tf-txt) !important}

/* ---------- cartões (expander) ---------- */
[data-testid="stExpander"]{border:1px solid var(--tf-linha) !important;border-radius:16px !important;
  background:linear-gradient(180deg,rgba(255,255,255,.035),rgba(255,255,255,.015)) !important}
[data-testid="stExpander"] summary{padding:.8rem 1rem}
[data-testid="stExpander"] summary:hover *{color:var(--tf-lima) !important}
[data-testid="stExpander"] details{border:0 !important}

/* ---------- avisos ---------- */
[data-testid="stAlert"]{border-radius:14px !important;border:1px solid var(--tf-linha2) !important;background:rgba(255,255,255,.04) !important}
[data-testid="stAlert"] *{color:var(--tf-txt) !important}
[data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]){border-color:rgba(96,165,250,.5) !important;background:rgba(96,165,250,.10) !important}
[data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]){border-color:rgba(52,211,153,.5) !important;background:rgba(52,211,153,.10) !important}
[data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]){border-color:rgba(251,191,36,.5) !important;background:rgba(251,191,36,.10) !important}
[data-testid="stAlert"]:has([data-testid="stAlertContentError"]){border-color:rgba(248,113,113,.55) !important;background:rgba(248,113,113,.10) !important}

/* ---------- seleção múltipla (tags) e métricas ---------- */
[data-baseweb="tag"]{background:var(--tf-lima-bg) !important;border:1px solid rgba(164,214,86,.35);border-radius:8px !important}
[data-baseweb="tag"] *{color:var(--tf-lima) !important}
[data-testid="stMetric"]{padding:14px 18px;border:1px solid var(--tf-linha);border-radius:16px;
  background:linear-gradient(180deg,rgba(255,255,255,.035),rgba(255,255,255,.015))}
[data-testid="stMetricLabel"] *{color:var(--tf-muted) !important;font-weight:700}
[data-testid="stMetricValue"] *{color:#fff !important;font-weight:800}

/* ---------- tabelas ---------- */
[data-testid="stDataFrame"]{border:1px solid var(--tf-linha);border-radius:14px;overflow:hidden;
  filter:invert(.92) hue-rotate(180deg) saturate(.8)}
[data-testid="stTable"] table, [data-testid="stTable"] th, [data-testid="stTable"] td{color:var(--tf-txt) !important;border-color:var(--tf-linha) !important}
[data-testid="stSpinner"] *{color:var(--tf-muted) !important}
[data-testid="stProgress"] > div > div{background:var(--tf-lima) !important}

::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-thumb{background:var(--tf-linha2);border-radius:8px}::-webkit-scrollbar-track{background:transparent}
</style>
"""


def aplicar_tema_escuro() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)


def html_cabecalho(titulo: str, sub: str, icone: str, cor: str) -> str:
    """HTML do título da ferramenta (chip + nome + subtítulo). Separado para poder ser testado.
    A classe `tf-ancora` marca a faixa do topo para o CSS."""
    return (
        '<div class="tf-topo tf-ancora"><div class="tf-tit">'
        f'<span class="tf-chip" style="--c:{html.escape(cor)}">'
        f'<svg viewBox="0 0 24 24" aria-hidden="true">{icone}</svg></span>'
        f'<div><h1>{html.escape(titulo)}</h1><p>{html.escape(sub)}</p></div></div></div>'
    )


def cabecalho_ferramenta(nome: str, voltar: Callable[[], None]) -> None:
    """Tema escuro + faixa do topo da ferramenta `nome` (que deve estar em TEMA_ESCURO).

    A marca no centro é um st.button estilizado (não um link): assim a sessão continua viva e os
    arquivos já enviados não se perdem ao voltar."""
    sub, icone, cor = TEMA_ESCURO[nome]
    aplicar_tema_escuro()
    col_titulo, col_marca, col_voltar = st.columns([3, 2, 3], vertical_alignment="center")
    with col_titulo:
        st.markdown(html_cabecalho(nome, sub, icone, cor), unsafe_allow_html=True)
    with col_marca:
        st.button("Início", key="tf_logo_home", on_click=voltar, help="Voltar à Central de ferramentas")
    with col_voltar:
        _, direita = st.columns([1, 1.6])
        with direita:
            st.button("← Central de ferramentas", key="voltar_inicio", on_click=voltar, use_container_width=True)
