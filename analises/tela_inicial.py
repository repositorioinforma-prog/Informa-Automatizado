"""
Central de ferramentas (tela inicial): categorias à esquerda, busca e lista de ferramentas à direita.
É a única porta de entrada do app (não há mais seletor na barra lateral).

Para mudar a tela, edite só a lista SECOES abaixo:
- uma categoria = {"titulo", "sub", "cor", "icone", "cards"}
    cor: "azul" | "roxo" | "verde" | "laranja" | "violeta" | "ciano"
- uma ferramenta (card) = {"titulo", "sub", "desc", "icone", "alvo", opcionais: "entrada", "tags", "legado"}
    alvo:    o nome EXATO da opção do MENU (app.py) que o card abre; sem alvo (None) aparece como "Em breve".
    sub:     frase curta que aparece na lista.
    desc:    finalidade (aparece no painel de baixo ao passar o mouse / focar a ferramenta).
    entrada: lista curta do que a ferramenta precisa (ex.: ["Descritivo .xlsx"]).
    tags:    palavras extras para a busca (sinônimos, siglas).
    legado:  True = vai para «Versões anteriores» da categoria.
Ícones disponíveis: ver o objeto I em componentes/tela_inicial/index.html.
"""
from __future__ import annotations

from typing import Any, Dict, List

import streamlit as st

from analises._ponte_relatoria import _componente

INICIO = "Início"
CHAVE_MENU = "analise_menu"
_CHAVE_CLIQUE = "inicio_cards"
_CHAVE_ULTIMO = "inicio_ultimo_clique"

SECOES: List[Dict[str, Any]] = [
    {
        "titulo": "Programações", "sub": "Prepare as programações das tabelas.", "cor": "azul", "icone": "calendario",
        "cards": [
            {"titulo": "Gerador de Programações SPSS", "sub": "Syntax, Total, Completo e Perfil em um só lugar",
             "desc": "Syntax, Total, Completo e Perfil em um só lugar. Entra quando o app externo for integrado.",
             "icone": "codigo", "alvo": None, "tags": ["spss", "syntax", "sps"]},
            {"titulo": "Revisor de Syntax SPSS", "sub": "Confere a syntax antes de rodar",
             "desc": "Confere a syntax SPSS antes de rodar.",
             "icone": "lupa", "alvo": "Revisor de Syntax SPSS", "tags": ["spss", "sps", "programação"]},
        ],
    },
    {
        "titulo": "Formatação", "sub": "Configure a estrutura e o layout das tabelas.", "cor": "roxo", "icone": "documento",
        "cards": [
            {"titulo": "Total", "sub": "Formatos de tabela",
             "desc": "Formatos de tabela do Total automático.",
             "icone": "tabela", "alvo": "Total Automático", "tags": ["total automático"]},
            {"titulo": "Descritivo", "sub": "Textos e cabeçalhos",
             "desc": "Textos e cabeçalhos do Relatório Descritivo.",
             "icone": "documento", "alvo": "Relatório Automatizado", "tags": ["relatório automatizado"]},
            {"titulo": "Padronização de Títulos", "sub": "O título padrão de cada tipo de pergunta",
             "desc": "Consulta o título padrão (e a pergunta) de cada tipo de pergunta: voto, rejeição, IPV, avaliação, mídia e outros.",
             "icone": "aa", "alvo": "Padronização de Títulos",
             "tags": ["títulos", "padrão", "pergunta", "ipv", "rejeição", "avaliação"]},
            {"titulo": "Divisor de Tabelas de Regiões", "sub": "Separa as tabelas por região",
             "desc": "Separa as tabelas por região.",
             "icone": "dividir", "alvo": "Divisor de Tabelas de Regioes", "tags": ["regioes", "divisor"]},
        ],
    },
    {
        "titulo": "Analíticos", "sub": "Explore os dados e prepare suas análises.", "cor": "verde", "icone": "barras",
        "cards": [
            {"titulo": "Mapas de Calor", "sub": "Visualização geográfica",
             "desc": "Visualização geográfica por região (mapas de calor).",
             "icone": "mapa", "alvo": "Mapa de Calor", "tags": ["regiões", "geográfico"]},
            {"titulo": "Cruzamentos", "sub": "Tabelas cruzadas e exportação",
             "desc": "Tabelas cruzadas e exportação do resultado.",
             "icone": "tabela", "alvo": "Cruzamentos (Tabelas)", "tags": ["tabelas", "cruzamento"]},
            {"titulo": "Mapa de Correspondência", "sub": "Associações entre variáveis",
             "desc": "Mapa de correspondência (versão atual, V2): associações entre variáveis.",
             "entrada": ["Base .csv, .xlsx ou .sav"],
             "icone": "dispersao", "alvo": "[Atual] Mapas de Correspondência - V2", "tags": ["correspondência", "v2"]},
            {"titulo": "Correlação", "sub": "Análise de variáveis",
             "desc": "Matriz de correlação entre variáveis.",
             "icone": "linha", "alvo": "Matriz de Correlação", "tags": ["matriz"]},
            {"titulo": "Análise Residual", "sub": "AR · Análise residual",
             "desc": "Análise residual (A.R) dos cruzamentos.",
             "icone": "residuo", "alvo": "Análise Residual (A.R)", "tags": ["ar", "resíduos"]},
            {"titulo": "Análise das Perguntas", "sub": "Cada pergunta do Total cruzada pelos segmentos",
             "desc": "Lista as perguntas do Total, confere os labels no banco (ou usa o título da tabela) e cruza cada uma pelos segmentos, em Excel.",
             "entrada": ["Programação .sps (Total)", "Banco .sav"],
             "icone": "tabela", "alvo": "Análise das Perguntas", "tags": ["perguntas", "total", "cruzamento", "resíduo"]},
            {"titulo": "Perfil", "sub": "Sexo, idade, renda, escolaridade e religião",
             "desc": "Gera as tabelas do Perfil (P1, P2, P3, P4 e religião) direto do banco, com o peso ligado, em Excel. Presencial leva a média de idade.",
             "entrada": ["Banco .sav"],
             "icone": "pessoas", "alvo": "Perfil", "tags": ["perfil", "sexo", "idade", "renda", "escolaridade", "religião"]},
            {"titulo": "Relatório Analítico", "sub": "Conferência do relatório e preparação da apresentação",
             "desc": "Confere o Relatório Descritivo (divergências entre totais e segmentos) e prepara o esqueleto da apresentação.",
             "entrada": ["Relatório Descritivo .xlsx"],
             "icone": "documento", "alvo": "Relatório Analítico", "tags": ["apresentação", "pptx", "descritivo"]},
            {"titulo": "Mapa de Correspondência V1", "sub": "Versão antiga, mantida para consulta",
             "desc": "Versão antiga do mapa de correspondência. Prefira a versão atual.",
             "entrada": ["Base .csv, .xlsx ou .sav"],
             "icone": "dispersao", "alvo": "[Antigo] Mapas de Correspondência - V1", "legado": True, "tags": ["correspondência", "v1"]},
        ],
    },
    {
        "titulo": "Controle", "sub": "Valide e monitore os resultados.", "cor": "laranja", "icone": "escudo",
        "cards": [
            {"titulo": "Controle Geral", "sub": "Recodificação geográfica, duplicidades e validação eleitoral",
             "desc": "Recodificação geográfica, duplicidades e validação eleitoral.",
             "icone": "validacao", "alvo": "Controle Geral", "tags": ["duplicidade", "eleitoral", "geografia"]},
            {"titulo": "Gerador de Amostra", "sub": "Distribuição da amostra",
             "desc": "Distribuição da amostra.",
             "icone": "amostra", "alvo": "Gerador de Amostra", "tags": ["cotas"]},
            {"titulo": "Banco de Amostras", "sub": "Amostras já geradas",
             "desc": "Amostras já geradas.",
             "icone": "banco", "alvo": "Banco de Amostras"},
        ],
    },
    {
        "titulo": "Codificação e ponderação", "sub": "Organize as respostas e aplique os pesos da amostra.",
        "cor": "violeta", "icone": "balanca",
        "cards": [
            {"titulo": "Codificação", "sub": "Categorias e rótulos",
             "desc": "Categorias e rótulos das respostas.",
             "icone": "codificacao", "alvo": None},
            {"titulo": "Ponderação", "sub": "Cálculo e aplicação dos pesos da amostra",
             "desc": "Cálculo e aplicação dos pesos da amostra.",
             "icone": "balanca", "alvo": "Ponderação", "tags": ["pesos"]},
        ],
    },
    {
        "titulo": "Validações", "sub": "Confira bases e tabelas.", "cor": "ciano", "icone": "checklist",
        "cards": [
            {"titulo": "Processador de Base Reduzida", "sub": "Base reduzida a partir do banco",
             "desc": "Base reduzida a partir do banco.",
             "icone": "reduzir", "alvo": "Processador de Base Reduzida"},
            {"titulo": "Exclusão de Colunas", "sub": "Remoção de colunas por base ou base reduzida",
             "desc": "Apaga colunas por Base / Base reduzida.",
             "icone": "tabela", "alvo": "Exclusão de Colunas por Base"},
            {"titulo": "Exclusões", "sub": "Casos a excluir do banco",
             "desc": "Casos a excluir do banco.",
             "icone": "lixeira", "alvo": "Exclusões"},
            {"titulo": "Base nas Múltiplas", "sub": "Bases das perguntas de múltipla resposta",
             "desc": "Bases das perguntas de múltipla resposta.",
             "icone": "tabela", "alvo": "Base nas Múltiplas"},
            {"titulo": "Legendas", "sub": "Legendas das tabelas",
             "desc": "Legendas das tabelas.",
             "icone": "etiqueta", "alvo": "Legendas"},
            {"titulo": "Códigos Individuais", "sub": "Conferência código a código",
             "desc": "Conferência código a código.",
             "icone": "codificacao", "alvo": "Códigos Individuais"},
        ],
    },
]


def aplicar_clique_do_inicio(opcoes_menu) -> None:
    """Chame no começo de cada execução: se um card foi clicado, a ferramenta
    atual (CHAVE_MENU) passa a ser a do card."""
    clique = st.session_state.get(_CHAVE_CLIQUE)
    if not isinstance(clique, dict) or not clique.get("nonce"):
        return
    if clique["nonce"] == st.session_state.get(_CHAVE_ULTIMO):
        return  # clique já atendido (o valor do componente fica guardado)
    st.session_state[_CHAVE_ULTIMO] = clique["nonce"]
    alvo = clique.get("alvo")
    if alvo in opcoes_menu:
        st.session_state[CHAVE_MENU] = alvo


_PARAM_URL = "ferramenta"


def restaurar_ferramenta(opcoes_menu) -> None:
    """Se a ferramenta atual se perdeu do session_state (reinício da sessão, rerun que limpou
    o estado, F5), volta para a que está guardada na URL (?ferramenta=...)."""
    if st.session_state.get(CHAVE_MENU) in opcoes_menu:
        return
    try:
        guardada = st.query_params.get(_PARAM_URL)
    except Exception:
        guardada = None
    if isinstance(guardada, (list, tuple)):
        guardada = guardada[0] if guardada else None
    if guardada in opcoes_menu:
        st.session_state[CHAVE_MENU] = guardada


def guardar_ferramenta(analise: str) -> None:
    """Grava a ferramenta atual na URL (só escreve quando muda). Início = sem parâmetro."""
    try:
        atual = st.query_params.get(_PARAM_URL)
        if isinstance(atual, (list, tuple)):
            atual = atual[0] if atual else None
        if analise == INICIO:
            if atual is not None:
                del st.query_params[_PARAM_URL]
        elif atual != analise:
            st.query_params[_PARAM_URL] = analise
    except Exception:
        pass  # versões antigas do Streamlit: sem persistência na URL, o resto funciona igual


def voltar_ao_inicio() -> None:
    """Callback do botão «Voltar ao Início» exibido no topo de cada ferramenta."""
    st.session_state[CHAVE_MENU] = INICIO


def tela_inicial() -> None:
    _componente("tela_inicial")(secoes=SECOES, key=_CHAVE_CLIQUE, default=None)
