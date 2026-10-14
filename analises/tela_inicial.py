"""
Tela inicial: seções com cards que abrem as ferramentas do menu lateral.

Para mudar a tela, edite só a lista SECOES abaixo:
- uma seção = {"titulo", "sub", "cor", "icone", "largura", "layout", "cards"}
    cor: "azul" | "roxo" | "verde" | "laranja" | "violeta" | "ciano"
    largura: colunas de 12 (7 + 5 numa linha, 12 = linha inteira)
    layout: "vertical" (ícone em cima) | "horizontal" (ícone ao lado)
    colunas: (opcional) cards por linha; sem isso, todos numa linha só
- um card = {"titulo", "sub", "icone", "alvo"}
    alvo: o nome EXATO da opção do menu lateral (MENU em app.py) que o card
    abre. Sem alvo (None) o card aparece como "Em breve".
Ícones disponíveis: ver o objeto I em componentes/tela_inicial/index.html
("calendario", "documento", "barras", "engrenagem", "banco", "codigo",
"tabela", "pessoas", "mapa", "dispersao", "linha", "validacao",
"codificacao", "balanca", "residuo", "lixeira", "etiqueta", "dividir", "lupa",
"amostra", "reduzir") e "aa" (texto "Aa").
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
        "titulo": "Programações", "sub": "Gere e gerencie as programações de tabelas.",
        "cor": "azul", "icone": "calendario", "largura": 7,
        "layout": "horizontal",
        "cards": [  # app desenvolvido por fora (um botão só com as 3 funções) — entra quando for integrado
            {"titulo": "Gerador de Programações SPSS", "sub": "Syntax, Total, Completo e Perfil em um só lugar",
             "icone": "codigo", "alvo": None},
        ],
    },
    {
        "titulo": "Formatação", "sub": "Configure a estrutura e o layout das tabelas.",
        "cor": "roxo", "icone": "documento", "largura": 5, "layout": "vertical",
        "cards": [
            {"titulo": "Total", "sub": "Formatos de tabela", "icone": "tabela", "alvo": "Total Automático"},
            {"titulo": "Descritivo", "sub": "Textos e cabeçalhos", "icone": "documento", "alvo": "Relatório Automatizado"},
        ],
    },
    {
        "titulo": "Analíticos", "sub": "Gere análises e visualizações dos dados.",
        "cor": "verde", "icone": "barras", "largura": 7, "layout": "vertical",
        "cards": [
            {"titulo": "Mapas de Calor", "sub": "Visualização geográfica", "icone": "mapa", "alvo": "Mapa de Calor"},
            {"titulo": "Mapa de Correspondência", "sub": "Visualização de associações entre variáveis",
             "icone": "dispersao", "alvo": "[Atual] Mapas de Correspondência - V2"},
            {"titulo": "Correlação", "sub": "Análise de variáveis", "icone": "linha", "alvo": "Matriz de Correlação"},
            {"titulo": "AR", "sub": "Análise Residual", "icone": "residuo", "alvo": "Análise Residual (A.R)"},
        ],
    },
    {
        "titulo": "Controle", "sub": "Valide e monitore os resultados.",
        "cor": "laranja", "icone": "engrenagem", "largura": 5, "layout": "horizontal",
        "cards": [
            {"titulo": "Controle Geral", "sub": "Recodificação geográfica, duplicidades e validação eleitoral",
             "icone": "validacao", "alvo": "Controle Geral"},
        ],
    },
    {
        "titulo": "Codificação e Ponderação", "sub": "Codifique as respostas e aplique a ponderação da amostra.",
        "cor": "violeta", "icone": "banco", "largura": 12, "layout": "horizontal",
        "cards": [
            {"titulo": "Codificação", "sub": "Gerencie categorias e rótulos", "icone": "codificacao", "alvo": None},
            {"titulo": "Ponderação", "sub": "Calcule e aplique os pesos da amostra", "icone": "balanca",
             "alvo": "Ponderação"},
        ],
    },
    {
        "titulo": "Validações", "sub": "Ferramentas de apoio para conferir bases, tabelas, amostras e syntax.",
        "cor": "ciano", "icone": "validacao", "largura": 12, "layout": "vertical", "colunas": 5,
        "cards": [
            {"titulo": "Processador de Base Reduzida", "sub": "Base reduzida a partir do banco",
             "icone": "reduzir", "alvo": "Processador de Base Reduzida"},
            {"titulo": "Exclusões", "sub": "Casos a excluir do banco", "icone": "lixeira", "alvo": "Exclusões"},
            {"titulo": "Base nas Múltiplas", "sub": "Bases das perguntas de múltipla resposta", "icone": "tabela",
             "alvo": "Base nas Múltiplas"},
            {"titulo": "Legendas", "sub": "Legendas das tabelas", "icone": "etiqueta", "alvo": "Legendas"},
            {"titulo": "Códigos Individuais", "sub": "Conferência código a código", "icone": "codificacao",
             "alvo": "Códigos Individuais"},
            {"titulo": "Gerador de Amostra", "sub": "Distribuição da amostra", "icone": "amostra",
             "alvo": "Gerador de Amostra"},
            {"titulo": "Banco de Amostras", "sub": "Amostras já geradas", "icone": "banco", "alvo": "Banco de Amostras"},
            {"titulo": "Divisor de Tabelas de Regiões", "sub": "Separa as tabelas por região", "icone": "dividir",
             "alvo": "Divisor de Tabelas de Regioes"},
            {"titulo": "Padronização de Títulos", "sub": "Títulos das tabelas no padrão", "icone": "aa",
             "alvo": "Padronização de Títulos"},
            {"titulo": "Revisor de Syntax SPSS", "sub": "Confere a syntax antes de rodar", "icone": "lupa",
             "alvo": "Revisor de Syntax SPSS"},
        ],
    },
]


def aplicar_clique_do_inicio(opcoes_menu) -> None:
    """Chame ANTES de desenhar o menu lateral: se um card foi clicado, o menu
    passa a apontar para a ferramenta do card."""
    clique = st.session_state.get(_CHAVE_CLIQUE)
    if not isinstance(clique, dict) or not clique.get("nonce"):
        return
    if clique["nonce"] == st.session_state.get(_CHAVE_ULTIMO):
        return  # clique já atendido (o valor do componente fica guardado)
    st.session_state[_CHAVE_ULTIMO] = clique["nonce"]
    alvo = clique.get("alvo")
    if alvo in opcoes_menu:
        st.session_state[CHAVE_MENU] = alvo


def tela_inicial() -> None:
    _componente("tela_inicial")(secoes=SECOES, key=_CHAVE_CLIQUE, default=None)
