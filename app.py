"""
Ponto de entrada do app. Responsável por: configuração da página,
CSS/tema visual, logo, e o roteamento entre os módulos de análise
(menu lateral). A lógica de cada análise mora em core/ e analises/.
"""
from io import BytesIO

import requests
import streamlit as st
from PIL import Image

from analises.base_reduzida import processador_base_reduzida
from analises.exclusoes import modulo_exclusoes
from analises.correspondencia_multipla import analise_correspondencia_multipla
from analises.correspondencia_simples import analise_correspondencia
from core.dados import carregar_dados


st.set_page_config(page_title="Análise de Dados Automática", layout="wide")

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
# Logo
# =========================
url = "https://institutoinforma.com.br/wp-content/uploads/2025/01/logo_informa.webp"
headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/110.0.0.0 Safari/537.36"
    )
}

try:
    response = requests.get(url, headers=headers, timeout=20)
    response.raise_for_status()
    image_bytes = BytesIO(response.content)
    logo = Image.open(image_bytes)
    st.image(logo, width=300)
except Exception:
    st.warning("Não foi possível carregar a logo do sistema.")


# =========================
# Função principal
# =========================
def main():
    analise = st.sidebar.selectbox(
        "Escolha a análise",
        (
            "Tratamento de Dados",
            "Mapas de Correspondência",
            "Mapas de Correspondência (Múltiplas Variáveis)",
            "Processador de Base Reduzida",
            "Exclusões",
        ),
    )

    # estes módulos têm upload próprio e não dependem do arquivo carregado abaixo
    if analise == "Processador de Base Reduzida":
        processador_base_reduzida()
        return

    if analise == "Exclusões":
        modulo_exclusoes()
        return

    uploaded_file = st.file_uploader("Carregar Arquivo", type=["csv", "xlsx", "sav"])

    if uploaded_file is not None:
        dados, dados_exibicao = carregar_dados(uploaded_file)

        if dados is not None:
            if analise == "Mapas de Correspondência":
                analise_correspondencia(dados)

            elif analise == "Mapas de Correspondência (Múltiplas Variáveis)":
                analise_correspondencia_multipla(dados)

    else:
        st.info("Por favor, carregue um arquivo para começar.")


if __name__ == "__main__":
    main()
