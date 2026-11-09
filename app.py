"""
Ponto de entrada do app. Responsável por: configuração da página,
CSS/tema visual, logo, e o roteamento entre os módulos de análise
(tela inicial com cards). A lógica de cada análise mora em core/ e analises/.
"""
import os

import streamlit as st

from analises.base_reduzida import processador_base_reduzida
from analises.exclusao_colunas import modulo_exclusao_colunas
from analises.base_multiplas import modulo_base_multiplas
from analises.legendas import modulo_legendas
from analises.relatorio_automatizado import modulo_relatorio_automatizado
from analises.total_automatico import modulo_total_automatico
from analises.codigos_individuais import modulo_codigos_individuais
from analises.exclusoes import modulo_exclusoes
from analises.gerador_amostra import modulo_gerador_amostra
from analises.banco_amostras_ui import modulo_banco_amostras
from analises.ponderacao import modulo_ponderacao
from analises.divisor_tabelas import modulo_divisor_tabelas
from analises.padronizacao_titulos import modulo_padronizacao_titulos
from analises.revisor_syntax import modulo_revisor_syntax
from analises.cruzamentos import modulo_cruzamentos
from analises.mapa_calor import modulo_mapa_calor
from analises.correlacao import modulo_correlacao
from analises.analise_residual import modulo_analise_residual
from analises.relatorio_analitico import modulo_relatorio_analitico
from analises.controle import modulo_controle, pre_carregar_base_brasil
from analises._tema_ferramenta import TEMA_ESCURO, cabecalho_ferramenta
from analises.tela_inicial import CHAVE_MENU, INICIO, aplicar_clique_do_inicio, guardar_ferramenta, restaurar_ferramenta, tela_inicial, voltar_ao_inicio
from analises.correspondencia_multipla import analise_correspondencia_multipla
from analises.correspondencia_simples import analise_correspondencia
from core.dados import carregar_dados


st.set_page_config(page_title="Informa Core", layout="wide")

st.markdown(
    """
    <style>
    :root {
        --accent: #2563eb;
        --accent-hover: #1d4ed8;
        --text-main: #111827;
        --text-muted: #4b5563;
        --border: #94a3b8;
        --bg-card: #f8fafc;
        --bg-app: #ffffff;
        --sidebar-bg: #1f2937;
    }

    html, body, .stApp {
        background: var(--bg-app);
        color: var(--text-main);
        font-size: 16px;
    }

    /* texto geral */
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div,
    .stMarkdown, .stText, h1, h2, h3, h4, h5, h6 {
        color: var(--text-main);
    }

    h1, h2, h3 {
        font-weight: 700;
        border-bottom: 3px solid var(--accent);
        padding-bottom: 6px;
        display: inline-block;
    }

    /* sidebar escura */
    [data-testid="stSidebar"] {
        background-color: var(--sidebar-bg);
        color: #f9fafb;
    }

    /* Só força cor clara em elementos de texto "folha" (p, label, span,
       small). NÃO inclui 'div' de propósito: divs também envolvem
       componentes como selectbox/multiselect, que têm fundo BRANCO — se
       forçássemos texto claro neles, o texto ficaria branco sobre
       branco (invisível). Esses componentes recebem sua própria regra
       de cor mais abaixo, com especificidade maior. */
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] small {
        color: #f9fafb !important;
    }

    /* selects/multiselects dentro da sidebar têm fundo branco (definido
       mais abaixo), então o texto deles precisa continuar escuro —
       especificidade maior que a regra acima para garantir a prioridade */
    [data-testid="stSidebar"] [data-baseweb="select"],
    [data-testid="stSidebar"] [data-baseweb="select"] * {
        color: var(--text-main) !important;
    }

    /* file uploader */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #111827 !important;
        border: 2px dashed #6b7280 !important;
        border-radius: 10px;
    }

    [data-testid="stFileUploaderDropzone"] * {
        color: #f9fafb !important;
    }

    [data-testid="stFileUploaderDropzone"] button {
        background-color: var(--accent) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px;
    }

    /* inputs de texto e número */
    textarea, input {
        background-color: #ffffff !important;
        color: var(--text-main) !important;
        border: 1.5px solid var(--border) !important;
        border-radius: 6px;
    }

    textarea:focus, input:focus {
        border-color: var(--accent) !important;
        outline: none !important;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.25) !important;
    }

    /* selects e multiselects (componentes BaseWeb, não são <input> puro) */
    [data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border: 1.5px solid var(--border) !important;
        border-radius: 6px !important;
    }

    [data-baseweb="select"] * {
        color: var(--text-main) !important;
    }

    [data-baseweb="tag"] {
        background-color: var(--accent) !important;
        color: #ffffff !important;
        border-radius: 4px !important;
    }

    [data-baseweb="tag"] svg {
        fill: #ffffff !important;
    }

    [data-baseweb="popover"] [data-baseweb="menu"] {
        background-color: #ffffff !important;
        border: 1px solid var(--border) !important;
    }

    [data-baseweb="menu"] li {
        color: var(--text-main) !important;
    }

    [data-baseweb="menu"] li:hover {
        background-color: #eff6ff !important;
    }

    /* header do Streamlit (barra fina no topo da página) — elemento
       separado da sidebar, mantém o tema escuro padrão do Streamlit por
       fora do nosso CSS. Força fundo e texto/ícones claros nele. */
    [data-testid="stHeader"] {
        background-color: var(--sidebar-bg) !important;
    }

    [data-testid="stHeader"] * {
        color: #f9fafb !important;
    }

    /* Ícones de recolher/expandir a sidebar ("»"/"«"), tanto no header
       (sidebar fechada) quanto dentro da própria sidebar (aberta).
       Só ajusta a cor do preenchimento (fill) — NÃO mexe em "stroke",
       pois forçar as duas juntas apaga o desenho interno do ícone
       (contorno e preenchimento ficam iguais, virando um bloco sólido
       sem forma reconhecível). Cor levemente suave (não branco puro)
       pra não pesar visualmente. */
    [data-testid="stHeader"] svg,
    [data-testid="stSidebar"] button svg {
        fill: #e5e7eb !important;
    }

    [data-testid="stHeader"] button:hover svg,
    [data-testid="stSidebar"] button:hover svg {
        fill: var(--accent) !important;
    }

    /* checkboxes e radios dentro da sidebar (fundo escuro) */
    [data-testid="stSidebar"] [data-testid="stCheckbox"] label,
    [data-testid="stSidebar"] [data-testid="stRadio"] label {
        color: #f9fafb !important;
    }

    /* tabelas */
    .dataframe {
        background-color: #ffffff;
        border: 1px solid var(--border);
        color: var(--text-main);
        border-radius: 8px;
    }

    .dataframe tbody tr:nth-child(odd) {
        background-color: #f1f5f9;
    }

    .dataframe tbody tr:nth-child(even) {
        background-color: #ffffff;
    }

    .dataframe thead {
        background-color: var(--accent);
        color: white;
    }

    /* botões (comuns, de formulário e de download) */
    .stButton button,
    .stFormSubmitButton button,
    .stDownloadButton button {
        background-color: var(--accent);
        border: none;
        border-radius: 8px;
        color: #ffffff;
        padding: 8px 18px;
        font-weight: 600;
        cursor: pointer;
        transition: background-color 0.15s ease-in-out;
    }

    .stButton button:hover,
    .stFormSubmitButton button:hover,
    .stDownloadButton button:hover {
        background-color: var(--accent-hover);
        color: #ffffff;
    }

    /* expanders viram "cards", ajuda a distinguir vários mapas na tela */
    [data-testid="stExpander"] {
        border: 1px solid var(--border) !important;
        border-radius: 10px !important;
        background-color: var(--bg-card) !important;
        margin-bottom: 14px;
    }

    [data-testid="stExpander"] summary {
        font-weight: 600;
        font-size: 1.05rem;
    }

    hr {
        border-top: 1.5px solid var(--border);
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================
# Logo (só nas ferramentas; a Central de ferramentas já traz a marca)
# =========================
LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.jpg")


def mostrar_logo():
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=300)
    else:
        st.warning("Não foi possível carregar a logo do sistema (arquivo não encontrado).")


# =========================
# Ferramentas (navegação pela tela inicial)
# =========================
# Os cards da tela inicial (analises/tela_inicial.py) apontam para estes nomes.
# Não há mais seletor na barra lateral: a ferramenta atual fica em st.session_state[CHAVE_MENU].
MENU = (
    INICIO,
    "[Antigo] Mapas de Correspondência - V1",
    "[Atual] Mapas de Correspondência - V2",
    "Processador de Base Reduzida",
    "Exclusão de Colunas por Base",
    "Exclusões",
    "Base nas Múltiplas",
    "Legendas",
    "Relatório Automatizado",
    "Total Automático",
    "Códigos Individuais",
    "Gerador de Amostra",
    "Banco de Amostras",
    "Ponderação",
    "Divisor de Tabelas de Regioes",
    "Padronização de Títulos",
    "Revisor de Syntax SPSS",
    "Cruzamentos (Tabelas)",
    "Mapa de Calor",
    "Matriz de Correlação",
    "Análise Residual (A.R)",
    "Relatório Analítico",
    "Controle Geral",
)


# Ferramentas que ainda usam a barra lateral (pilha de exportação / parâmetros). Nas demais ela some.
USAM_BARRA_LATERAL = {
    "Cruzamentos (Tabelas)",
    "Mapa de Calor",
    "Análise Residual (A.R)",
    "Gerador de Amostra",
}

_CSS_SEM_BARRA = """
<style>
[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"],
[data-testid="collapsedControl"] { display: none !important; }
</style>
"""

# Tela inicial em tela cheia: fundo escuro na página inteira e o componente (iframe) cobrindo a janela.
# O cabeçalho do Streamlit (menu ⋮) continua acessível, por cima, transparente.
_CSS_INICIO = """
<style>
html, body, .stApp, [data-testid="stApp"], [data-testid="stAppViewContainer"],
[data-testid="stMain"], section.main { background: #121923 !important; overflow: hidden !important; }
[data-testid="stHeader"] { background: transparent !important; z-index: 1000 !important; }
[data-testid="stHeader"] *, [data-testid="stToolbar"] * { color: #cbd5e1 !important; }
[data-testid="stDecoration"], footer, [data-testid="stStatusWidget"] { display: none !important; }
[data-testid="stMainBlockContainer"], .block-container, [data-testid="stAppViewBlockContainer"] {
  padding: 0 !important; max-width: none !important; }
[data-testid="stMain"] iframe, section.main iframe {
  position: fixed !important; inset: 0 !important; width: 100vw !important; height: 100vh !important;
  border: 0 !important; z-index: 1 !important; }
</style>
"""


# =========================
# Função principal
# =========================
def main():
    pre_carregar_base_brasil()  # lê a Base Brasil em segundo plano (só na 1ª execução)
    # card clicado na tela inicial -> a ferramenta atual passa a ser a do card
    aplicar_clique_do_inicio(MENU)
    restaurar_ferramenta(MENU)  # estado perdido (rerun/F5) -> volta à ferramenta guardada na URL
    if st.session_state.get(CHAVE_MENU) not in MENU:
        st.session_state[CHAVE_MENU] = INICIO
    analise = st.session_state[CHAVE_MENU]
    guardar_ferramenta(analise)

    if analise not in USAM_BARRA_LATERAL:
        st.markdown(_CSS_SEM_BARRA, unsafe_allow_html=True)

    if analise == INICIO:
        st.markdown(_CSS_INICIO, unsafe_allow_html=True)
        tela_inicial()
        return

    if analise in TEMA_ESCURO:   # ferramentas já repaginadas no visual da Central
        cabecalho_ferramenta(analise, voltar=voltar_ao_inicio)
    else:
        mostrar_logo()
        st.button("← Central de ferramentas", key="voltar_inicio", on_click=voltar_ao_inicio)

    # estes módulos têm upload próprio e não dependem do arquivo carregado abaixo
    if analise == "Processador de Base Reduzida":
        processador_base_reduzida()
        return

    if analise == "Exclusão de Colunas por Base":
        modulo_exclusao_colunas()
        return

    if analise == "Exclusões":
        modulo_exclusoes()
        return

    if analise == "Base nas Múltiplas":
        modulo_base_multiplas()
        return

    if analise == "Legendas":
        modulo_legendas()
        return

    if analise == "Relatório Automatizado":
        modulo_relatorio_automatizado()
        return

    if analise == "Total Automático":
        modulo_total_automatico()
        return

    if analise == "Códigos Individuais":
        modulo_codigos_individuais()
        return

    if analise == "Gerador de Amostra":
        modulo_gerador_amostra()
        return

    if analise == "Banco de Amostras":
        modulo_banco_amostras()
        return

    if analise == "Ponderação":
        modulo_ponderacao()
        return

    if analise == "Divisor de Tabelas de Regioes":
        modulo_divisor_tabelas()
        return

    if analise == "Padronização de Títulos":
        modulo_padronizacao_titulos()
        return

    if analise == "Revisor de Syntax SPSS":
        modulo_revisor_syntax()
        return

    # ferramentas portadas do app desktop "Relatória Analítica" (liberadas uma por vez)
    if analise == "Cruzamentos (Tabelas)":
        modulo_cruzamentos()
        return

    if analise == "Mapa de Calor":
        modulo_mapa_calor()
        return

    if analise == "Matriz de Correlação":
        modulo_correlacao()
        return

    if analise == "Análise Residual (A.R)":
        modulo_analise_residual()
        return

    if analise == "Relatório Analítico":
        modulo_relatorio_analitico()
        return

    if analise == "Controle Geral":
        modulo_controle()
        return

    uploaded_file = st.file_uploader("Carregar Arquivo", type=["csv", "xlsx", "sav"])

    if uploaded_file is not None:
        dados, dados_exibicao = carregar_dados(uploaded_file)

        if dados is not None:
            if analise == "[Antigo] Mapas de Correspondência - V1":
                analise_correspondencia(dados)

            elif analise == "[Atual] Mapas de Correspondência - V2":
                analise_correspondencia_multipla(dados)

    else:
        st.info("Por favor, carregue um arquivo para começar.")


if __name__ == "__main__":
    main()
