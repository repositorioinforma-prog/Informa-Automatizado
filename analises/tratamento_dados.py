"""
Tratamento de Dados — tela do módulo no estilo do SPSS (a lógica mora em core/tratamento_dados_math.py).

Organização (como no Editor de Dados do SPSS):
  * título da janela e faixas de comandos por tarefa (Variáveis · Casos · Rótulos e ausentes · Juntar bancos ·
    Arquivo · Ajuda), sempre visíveis, com grupos; cada comando abre uma janela;
  * janelas com os botões OK / Ver syntax / Reiniciar / Cancelar;
  * grades limpas em HTML (cabeçalho cinza, nº do caso) e abas de baixo: dados · variáveis · Syntax;
  * barra de status com arquivo, casos × variáveis e passos aplicados.

As partes puras (menus, ajudas, conversão das grades) ficam em core/tratamento_dados_dialogos.py.
Chaves de estado: prefixo "td_".
"""
from __future__ import annotations

import html
import os
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

from core import tratamento_dados_dialogos as dlg
from core import tratamento_dados_math as td
from core.tratamento_dados_math import Banco, ErroTratamento, Passo

K_ORIG, K_ATUAL, K_PASSOS = "td_original", "td_atual", "td_passos"
K_ASSIN, K_MSG, K_VER = "td_assinatura", "td_mensagem", "td_versao"
K_OUTRO, K_ASSIN2, K_SAIDAS = "td_outro", "td_assinatura2", "td_saidas"
K_FOCO = "td_variavel_foco"      # variável escolhida na Visualização de variáveis (pré-seleciona as janelas)
K_RESET = "td_reinicio"          # "Reiniciar" da janela: muda as chaves dos widgets e recria tudo em branco

CSS = """
<style>
/* Visual limpo, no estilo do SPSS. TODAS as cores são explícitas (fundo e texto juntos): o tema do Streamlit
   pode ser escuro e o app.py força texto escuro, o que deixava texto escuro sobre fundo escuro. */

/* título da janela */
.td-titulo { text-align: center; font-size: .85rem; color: #111827; background: #ffffff;
             border: 1px solid #d1d5db; padding: 5px 0; margin-top: 4px; }

/* faixas de comandos (abas por tarefa) */
.st-key-td_faixa { background: #ffffff; border: 1px solid #d1d5db; border-top: none; padding: 0 12px 10px 12px; }
.st-key-td_faixa [data-baseweb="tab"] p { color: #111827 !important; font-size: .95rem; }
.st-key-td_faixa [data-baseweb="tab"][aria-selected="true"] p { color: #1d4ed8 !important; font-weight: 700; }
.st-key-td_faixa [data-baseweb="tab-highlight"] { background-color: #2563eb !important; }
.st-key-td_faixa [data-baseweb="tab-border"] { background-color: #d1d5db !important; }
.td-grupo { color: #4b5563; font-size: .72rem; font-weight: 700; letter-spacing: .05em; text-transform: uppercase;
            border-bottom: 1px solid #d1d5db; padding-bottom: 3px; margin: 6px 0 8px 0; }
[class*="st-key-td_c_"] button { background: #ffffff !important; color: #111827 !important; border: 1px solid #cbd5e1 !important;
    border-radius: 6px !important; box-shadow: none !important; justify-content: flex-start !important; text-align: left !important;
    font-weight: 600 !important; font-size: .9rem !important; padding: 6px 12px !important; min-height: 2.3rem !important; }
[class*="st-key-td_c_"] button * { color: #111827 !important; fill: #111827 !important; }
[class*="st-key-td_c_"] button:hover { background: #eff6ff !important; border-color: #2563eb !important; }
.st-key-td_faixa [data-testid="stCaptionContainer"], .st-key-td_faixa [data-testid="stCaptionContainer"] * {
    color: #4b5563 !important; font-size: .76rem !important; line-height: 1.25; }
[class*="st-key-td_q_"] button { background: #ffffff !important; color: #111827 !important; border: 1px solid #cbd5e1 !important;
    border-radius: 6px !important; box-shadow: none !important; font-weight: 500 !important; font-size: .85rem !important; }
[class*="st-key-td_q_"] button * { color: #111827 !important; fill: #111827 !important; }
[class*="st-key-td_q_"] button:hover { background: #eff6ff !important; border-color: #2563eb !important; }
[class*="st-key-td_q_"] button:disabled, [class*="st-key-td_q_"] button:disabled * { color: #9ca3af !important; background: #f3f4f6 !important; }

/* grade (Data View / Variable View) */
.td-grid { max-height: 520px; overflow: auto; background: #ffffff; border: 1px solid #c8c8c8; }
.td-grid table { border-collapse: separate; border-spacing: 0; font-size: .82rem; color: #111827; background: #ffffff; width: max-content; min-width: 100%; }
.td-grid th { position: sticky; top: 0; z-index: 2; background: #f0f0f0; color: #111827; font-weight: 400;
              border-right: 1px solid #d0d0d0; border-bottom: 1px solid #d0d0d0; padding: 3px 10px; text-align: center; white-space: nowrap; }
.td-grid th.td-rn { left: 0; z-index: 3; min-width: 2.4rem; color: #4b5563; }
.td-grid tbody th.td-rn { z-index: 1; position: sticky; }
.td-grid td { background: #ffffff; color: #111827; border-right: 1px solid #e5e7eb; border-bottom: 1px solid #e5e7eb;
              padding: 2px 10px; white-space: nowrap; max-width: 28rem; overflow: hidden; text-overflow: ellipsis; }
.td-grid td.td-n { text-align: right; }
.td-grid tbody tr:hover td { background: #f1f8ff; }

/* barra de estado */
.td-status { display: flex; flex-wrap: wrap; gap: 4px 0; background: #f0f0f0; border: 1px solid #d1d5db; color: #111827;
             font-size: .8rem; padding: 3px 0; margin-top: 2px; }
.td-status span { color: #111827; padding: 0 16px; border-right: 1px solid #c8c8c8; }
.td-status span:last-child { border-right: none; }
.td-status .td-ok { color: #16a34a; padding: 0; border: none; }

/* abas de baixo (Visualização de dados / variáveis / Syntax) */
.stApp [data-testid="stButtonGroup"] button, .stApp button[kind^="segmented_control"],
.stApp button[data-testid^="stBaseButton-segmented_control"] {
    background: #e5e7eb !important; border: 1px solid #c8c8c8 !important; border-radius: 0 !important; color: #111827 !important;
    font-weight: 400 !important; box-shadow: none !important; }
.stApp [data-testid="stButtonGroup"] button *, .stApp button[kind^="segmented_control"] * { color: #111827 !important; }
.stApp [data-testid="stButtonGroup"] button:hover { background: #f3f4f6 !important; }
.stApp button[kind="segmented_controlActive"], .stApp button[data-testid="stBaseButton-segmented_controlActive"],
.stApp [data-testid="stButtonGroup"] button[data-testid$="Active"], .stApp [data-testid="stButtonGroup"] button[kind$="Active"],
.stApp [data-testid="stButtonGroup"] button[aria-checked="true"], .stApp [data-testid="stButtonGroup"] button[aria-pressed="true"] {
    background: #ffffff !important; border-bottom: 3px solid #2563eb !important; font-weight: 700 !important; }
.stApp button[kind="segmented_controlActive"] *, .stApp button[data-testid="stBaseButton-segmented_controlActive"] *,
.stApp [data-testid="stButtonGroup"] button[data-testid$="Active"] *, .stApp [data-testid="stButtonGroup"] button[kind$="Active"] *,
.stApp [data-testid="stButtonGroup"] button[aria-checked="true"] *, .stApp [data-testid="stButtonGroup"] button[aria-pressed="true"] * {
    color: #111827 !important; }

/* expanders (abrir arquivo, "Se…") */
[data-testid="stExpander"] details, [data-testid="stExpander"] summary { background: #f1f5f9 !important; color: #111827 !important; }
[data-testid="stExpander"] summary * { color: #111827 !important; }
[data-testid="stExpander"] summary svg { fill: #111827 !important; }
[data-testid="stExpander"] [data-testid="stExpanderDetails"] { background: #f8fafc !important; }

/* janelas de diálogo: fundo claro e texto escuro. Usa só [role="dialog"] porque o data-testid muda entre versões. */
[role="dialog"], [role="dialog"] > div { background: #ffffff !important; }
[role="dialog"] :is(p, label, li, h1, h2, h3, h4, span, small, summary, strong, em, div):not(button *):not(pre *):not(code):not(svg *):not([data-baseweb="select"] *):not([data-baseweb="tag"] *):not([data-testid="stFileUploaderDropzone"] *) { color: #111827; }
[role="dialog"] h1, [role="dialog"] h2, [role="dialog"] h3, [role="dialog"] [data-testid="stHeading"] * { color: #111827 !important; }
[role="dialog"] [data-testid="stWidgetLabel"] *, [role="dialog"] [data-testid="stRadio"] label *,
[role="dialog"] [data-testid="stCheckbox"] label *, [role="dialog"] [data-testid="stExpander"] summary * { color: #111827 !important; }
[role="dialog"] [data-testid="stCaptionContainer"], [role="dialog"] [data-testid="stCaptionContainer"] *:not(code) { color: #4b5563 !important; }
[role="dialog"] [data-testid="stRadio"] [role="radiogroup"] label > div:first-child { border-color: #2563eb !important; }
[role="dialog"] button { background: #ffffff !important; color: #111827 !important; border: 1px solid #94a3b8 !important;
    border-radius: 6px !important; box-shadow: none !important; font-weight: 500 !important; white-space: nowrap; }
[role="dialog"] button * { color: #111827 !important; }
[role="dialog"] button:hover { background: #eff6ff !important; border-color: #2563eb !important; }
[role="dialog"] button[kind="primary"], [role="dialog"] button[data-testid="stBaseButton-primary"] { background: #2563eb !important; border-color: #2563eb !important; }
[role="dialog"] button[kind="primary"] *, [role="dialog"] button[data-testid="stBaseButton-primary"] * { color: #ffffff !important; }
[role="dialog"] button:disabled, [role="dialog"] button:disabled * { color: #9ca3af !important; }
[role="dialog"] [class*="_kp_"] button { background: #f3f4f6 !important; padding: 2px 0 !important; min-height: 2rem !important; }
/* o X de fechar a janela */
[role="dialog"] button[aria-label="Close"], [role="dialog"] button[aria-label="Close"] * { background: #ffffff !important; color: #111827 !important; }

/* trechos de código, caixas de syntax e upload de arquivo: claros */
code:not(pre code) { background: #eef2f7 !important; color: #111827 !important; border-radius: 4px; padding: 0 4px; }
[data-testid="stCode"], [data-testid="stCode"] pre { background: #f8fafc !important; border: 1px solid #d1d5db; border-radius: 6px; }
[data-testid="stCode"] pre, [data-testid="stCode"] pre *, [data-testid="stCode"] code { color: #111827 !important; background: transparent !important; }
[data-testid="stFileUploaderDropzone"] { background: #f8fafc !important; border: 2px dashed #94a3b8 !important; }
[data-testid="stFileUploaderDropzone"] *:not(button):not(button *) { color: #111827 !important; }
[data-testid="stFileUploaderFile"], [data-testid="stFileUploaderFile"] * { background: #eef2f7 !important; color: #111827 !important; }
[data-testid="stFileUploaderFile"] svg { fill: #111827 !important; }

/* botões da Visualização de variáveis: o texto não pode quebrar nem cortar */
[class*="st-key-td_vv_"] button { padding: 4px 8px !important; white-space: nowrap !important; }
[class*="st-key-td_vv_"] button p, [class*="st-key-td_vv_"] button div { white-space: nowrap !important; word-break: keep-all !important;
    overflow: visible !important; text-overflow: clip !important; }
</style>
"""


# ----------------------------------------------------------------------------- estado
def _estado_vazio() -> None:
    for k in (K_ORIG, K_ATUAL, K_PASSOS, K_MSG, K_OUTRO, K_ASSIN2, K_SAIDAS):
        st.session_state.pop(k, None)


def _carregar(up) -> None:
    assinatura = (up.name, up.size)
    if st.session_state.get(K_ASSIN) == assinatura and K_ATUAL in st.session_state:
        return
    _estado_vazio()
    banco = td.carregar_banco(up.getvalue(), up.name)
    st.session_state[K_ORIG] = banco
    st.session_state[K_ATUAL] = banco
    st.session_state[K_PASSOS] = []
    st.session_state[K_ASSIN] = assinatura
    st.session_state[K_VER] = 0


def _nova_versao() -> None:
    st.session_state[K_VER] = st.session_state.get(K_VER, 0) + 1
    st.session_state.pop(K_SAIDAS, None)
    st.session_state.pop(K_FOCO, None)


def _k(nome: str) -> str:
    """Chave de widget das janelas; muda ao aplicar um passo ou ao clicar em Reiniciar."""
    return f"td_d_{nome}_{st.session_state.get(K_VER, 0)}_{st.session_state.get(K_RESET, 0)}"


def _aplicar(operacao: str, **params: Any) -> None:
    """Executa a operação; em sucesso guarda o resultado, fecha a janela e recarrega a tela."""
    try:
        novo, passo = td.aplicar(st.session_state[K_ATUAL], operacao, **params)
    except ErroTratamento as exc:
        st.error(str(exc))
        return
    except Exception as exc:  # erro inesperado: mostra sem derrubar o app
        st.error(f"Não foi possível aplicar a operação: {exc}")
        return
    st.session_state[K_ATUAL] = novo
    st.session_state[K_PASSOS] = st.session_state[K_PASSOS] + [passo]
    st.session_state[K_MSG] = passo
    _nova_versao()
    st.rerun()


def _desfazer(quantos: Optional[int] = 1) -> None:
    passos: List[Passo] = st.session_state[K_PASSOS]
    restam = [] if quantos is None else passos[:-quantos]
    st.session_state[K_ATUAL] = td.reaplicar(st.session_state[K_ORIG], restam)
    st.session_state[K_PASSOS] = restam
    st.session_state.pop(K_MSG, None)
    _nova_versao()
    st.rerun()


def _rerun_da_janela() -> None:
    """Recarrega só a janela aberta (mantém o diálogo); versões antigas do Streamlit recarregam tudo."""
    try:
        st.rerun(scope="fragment")
    except TypeError:
        st.rerun()


# ----------------------------------------------------------------------------- utilidades de tela
def _rot(b: Banco, c: str) -> str:
    r = b.rotulos_var.get(c, "")
    return f"{c} — {r[:45]}" if r else c


def _opcoes(b: Banco, so: Optional[str] = None) -> List[str]:
    cols = b.colunas
    if so == "num":
        cols = [c for c in cols if not td.eh_texto(b.df[c])]
    elif so == "txt":
        cols = [c for c in cols if td.eh_texto(b.df[c])]
    return cols


def _multi(b: Banco, titulo: str, chave: str, so: Optional[str] = None, **kw) -> List[str]:
    return st.multiselect(titulo, _opcoes(b, so), format_func=lambda c: _rot(b, c), key=chave, **kw)


def _foco(b: Banco) -> Optional[str]:
    """Variável escolhida na Visualização de variáveis (para as janelas já abrirem nela)."""
    v = st.session_state.get(K_FOCO)
    return v if v in b.colunas else None


def _um(b: Banco, titulo: str, chave: str, so: Optional[str] = None, vazio: bool = False,
        padrao: Optional[str] = None) -> Optional[str]:
    ops = ([""] if vazio else []) + _opcoes(b, so)
    indice = ops.index(padrao) if padrao in ops else 0
    return st.selectbox(titulo, ops, index=indice, format_func=lambda c: _rot(b, c) if c else "—", key=chave) or None


def _lista_nomes(texto: str) -> List[str]:
    return [p for p in texto.replace(",", " ").split() if p]


def _fmt_valor(k: Any) -> str:
    return k if isinstance(k, str) else td._num(k)


def _nome_saida(orig: Banco) -> str:
    return f"{os.path.splitext(orig.nome)[0]}_tratado"


def _versao_streamlit() -> Tuple[int, int]:
    try:
        a, b = st.__version__.split(".")[:2]
        return int(a), int(b)
    except Exception:
        return (0, 0)


def _baixar(rotulo: str, dados: Any, nome: str, mime: str, **kw) -> None:
    """Botão de download que, nas versões novas, não recarrega a tela (e não fecha a janela) ao clicar."""
    if _versao_streamlit() >= (1, 43):
        st.download_button(rotulo, dados, file_name=nome, mime=mime, on_click="ignore", **kw)
    else:
        st.download_button(rotulo, dados, file_name=nome, mime=mime, **kw)


def _caixa(chave: str):
    try:
        return st.container(key=chave)
    except TypeError:               # Streamlit sem container(key=...): só perde o estilo da barra
        return st.container()


def _carregar_segundo_banco(chave: str) -> Optional[Banco]:
    up = st.file_uploader("Segundo banco (.sav)", type=["sav"], key=chave)
    if up is None:
        return None
    assin = (up.name, up.size)
    if st.session_state.get(K_ASSIN2) != assin or K_OUTRO not in st.session_state:
        try:
            st.session_state[K_OUTRO] = td.carregar_banco(up.getvalue(), up.name)
            st.session_state[K_ASSIN2] = assin
        except Exception as exc:
            st.error(f"Não foi possível ler o segundo banco: {exc}")
            return None
    outro: Banco = st.session_state[K_OUTRO]
    st.caption(f"{outro.nome}: {len(outro.df)} casos e {len(outro.df.columns)} variáveis.")
    return outro


# ----------------------------------------------------------------------------- peças das janelas
def _previa(operacao: str, params: Dict[str, Any]) -> None:
    """'Ver syntax': mostra o comando equivalente sem alterar o banco."""
    try:
        _, passo = td.aplicar(st.session_state[K_ATUAL], operacao, **params)
    except ErroTratamento as exc:
        st.error(str(exc))
        return
    except Exception as exc:
        st.error(f"Não foi possível montar a syntax: {exc}")
        return
    st.caption("Syntax equivalente (nada foi alterado no banco):")
    st.code(passo.sintaxe, language="text")
    for a in passo.avisos:
        st.warning(a)


def _rodape(operacao: str, coletar: Callable[[], Dict[str, Any]]) -> None:
    """Botões da janela, como no SPSS: OK · Ver syntax (Paste) · Reiniciar (Reset) · Cancelar."""
    st.divider()
    c1, c2, c3, c4 = st.columns(4)
    ok = c1.button("OK", type="primary", use_container_width=True, key=_k("ok"))
    ver = c2.button("Ver syntax", use_container_width=True, key=_k("ver_syntax"))
    reiniciar = c3.button("Reiniciar", use_container_width=True, key=_k("reiniciar"))
    cancelar = c4.button("Cancelar", use_container_width=True, key=_k("cancelar"))
    if cancelar:
        st.rerun()
    if reiniciar:
        st.session_state[K_RESET] = st.session_state.get(K_RESET, 0) + 1
        _rerun_da_janela()
    if ok or ver:
        try:
            params = coletar()
        except ErroTratamento as exc:
            st.error(str(exc))
            return
        if ok:
            _aplicar(operacao, **params)
        else:
            _previa(operacao, params)


def _contagem(nome: str, minimo: int) -> int:
    """Quantas linhas de campos a janela mostra (cresce com “Adicionar linha”; Reiniciar volta ao mínimo)."""
    return int(st.session_state.get(_k(nome), minimo))


def _botao_mais(nome: str, n: int) -> None:
    if st.button("＋ Adicionar linha", key=_k(f"{nome}_mais")):
        st.session_state[_k(nome)] = n + 1
        _rerun_da_janela()


def _fechar() -> None:
    if st.button("Fechar", use_container_width=True, key=_k("fechar")):
        st.rerun()


def _inserir(chave: str, trecho: str) -> None:
    st.session_state[chave] = dlg.inserir_trecho(st.session_state.get(chave, ""), trecho)


def _editor_expressao(b: Banco, nome: str, titulo: str, dica: str, altura: int = 100) -> str:
    """Caixa de expressão no estilo do Compute Variable: variáveis e funções à esquerda, teclado embaixo."""
    chave = _k(nome)
    esq, dire = st.columns([2, 3])
    with esq:
        st.markdown("**Variáveis**")
        var = st.selectbox("Variáveis do banco", b.colunas, format_func=lambda c: _rot(b, c),
                           key=_k(f"{nome}_var"), label_visibility="collapsed")
        st.button("Inserir variável ➜", key=_k(f"{nome}_ins_var"), use_container_width=True,
                  on_click=_inserir, args=(chave, var or ""))
        st.markdown("**Funções**")
        grupo = st.selectbox("Grupo de funções", list(dlg.FUNCOES), key=_k(f"{nome}_grp"),
                             label_visibility="collapsed")
        fn = st.selectbox("Função", dlg.FUNCOES[grupo], key=_k(f"{nome}_fn"), label_visibility="collapsed")
        st.button("Inserir função ➜", key=_k(f"{nome}_ins_fn"), use_container_width=True,
                  on_click=_inserir, args=(chave, fn))
    with dire:
        texto = st.text_area(titulo, key=chave, height=altura, placeholder=dica)
        for i, linha in enumerate(dlg.TECLADO):
            for col, (rotulo, trecho) in zip(st.columns(len(linha)), linha):
                col.button(rotulo, key=_k(f"{nome}_kp_{i}_{trecho}"), use_container_width=True,
                           on_click=_inserir, args=(chave, trecho))
        st.caption("Os botões inserem no fim da expressão. Veja todas as funções em Ajuda ▸ Funções e operadores.")
    return texto


# ----------------------------------------------------------------------------- janelas (uma por item de menu)
def _dlg_compute(b: Banco) -> None:
    st.caption(dlg.AJUDA["compute"])
    c1, c2 = st.columns(2)
    alvo = c1.text_input("Variável de destino (nova ou existente)", key=_k("alvo"), placeholder="ex.: Idade_meses")
    rot = c2.text_input("Rótulo da variável (opcional)", key=_k("rot"))
    expr = _editor_expressao(b, "expr", "Expressão numérica ou de texto", "ex.: MEAN(Q1 TO Q5)  ·  Idade * 12")
    with st.expander("Se… (calcular só para os casos que satisfazem uma condição)"):
        cond = st.text_input("Condição (opcional — equivale ao IF)", key=_k("cond"),
                             placeholder="ex.: Q1 = 1 AND Idade >= 18")
    _rodape("compute", lambda: dict(alvo=alvo, expressao=expr, condicao=cond or None, rotulo=rot or None))


def _dlg_recode(b: Banco, mesmas: bool) -> None:
    st.caption(dlg.AJUDA["recode_mesmas" if mesmas else "recode_outras"])
    origens = _multi(b, "Variável(is) de origem", _k("orig"))
    novos, rot = "", ""
    if mesmas:
        st.warning("Os valores originais serão substituídos. Para guardá-los, use Transformar ▸ "
                   "Recodificar em outras variáveis (ou Desfazer depois).")
    else:
        c1, c2 = st.columns(2)
        novos = c1.text_input("Nome(s) da(s) variável(is) de saída (um por origem, separados por espaço)",
                              key=_k("novos"), placeholder="ex.: Faixa_etaria")
        rot = c2.text_input("Rótulo da variável de saída (opcional, para uma só)", key=_k("rot"))
    st.markdown("**Valores antigos e novos**")
    modo = st.radio("Como informar as regras", ["Tabela", "Texto (syntax)"], horizontal=True, key=_k("modo"))
    linhas, texto = None, ""
    if modo == "Tabela":
        n = _contagem("n_regras", 4)
        linhas = []
        for col, titulo in zip(st.columns([3, 2, 2, 2]), ("Valor antigo", "De / valor", "Até", "Novo valor")):
            col.caption(titulo)
        for i in range(1, n + 1):
            c1, c2, c3, c4 = st.columns([3, 2, 2, 2])
            linhas.append({
                "Tipo": c1.selectbox(f"Valor antigo {i}", dlg.TIPOS_ANTIGO, key=_k(f"r_t{i}"), label_visibility="collapsed"),
                "De": c2.text_input(f"De {i}", key=_k(f"r_d{i}"), label_visibility="collapsed"),
                "Até": c3.text_input(f"Até {i}", key=_k(f"r_a{i}"), label_visibility="collapsed"),
                "Novo valor": c4.text_input(f"Novo valor {i}", key=_k(f"r_n{i}"), label_visibility="collapsed"),
            })
        _botao_mais("n_regras", n)
        st.caption("Uma regra por linha; a primeira que casar vale. Novo valor: um número, `SYSMIS` (ausente) ou "
                   "`COPY` (mantém o valor). “Todos os outros valores” vai sempre por último. Linhas em branco são ignoradas.")
    else:
        texto = st.text_area("Regras (uma por linha)", key=_k("regras"), height=120,
                             placeholder="LO THRU 29 = 1\n30 THRU 49 = 2\nELSE = COPY")
    rv = st.text_area("Rótulos dos valores novos (opcional)", key=_k("rv"), height=70,
                      placeholder="1 = Jovem\n2 = Adulto")

    def coletar() -> Dict[str, Any]:
        if not origens:
            raise ErroTratamento("Escolha ao menos uma variável de origem.")
        destinos = None if mesmas else _lista_nomes(novos)
        if not mesmas and not destinos:
            raise ErroTratamento("Informe o nome da variável de saída.")
        regras = dlg.regras_de_tabela(linhas) if linhas is not None else texto
        rotulos = td.analisar_rotulos_valor(rv, True) if rv.strip() else None
        return dict(origens=origens, regras=regras, destinos=destinos, rotulo_variavel=rot or None,
                    rotulos_valor=rotulos)

    _rodape("recode", coletar)


def _dlg_recode_mesmas(b: Banco) -> None:
    _dlg_recode(b, True)


def _dlg_recode_outras(b: Banco) -> None:
    _dlg_recode(b, False)


def _dlg_renomear(b: Banco) -> None:
    st.caption(dlg.AJUDA["renomear_variavel"])
    foco = _foco(b)
    atual = st.selectbox("Variável", _opcoes(b), index=_opcoes(b).index(foco) if foco else 0,
                         format_func=lambda c: _rot(b, c), key=_k("atual"))
    novo = st.text_input("Novo nome", key=_k("novo"))
    _rodape("renomear_variavel", lambda: dict(atual=atual, novo=novo))


def _dlg_excluir_vars(b: Banco) -> None:
    st.caption(dlg.AJUDA["excluir_variaveis"])
    nomes = _multi(b, "Variáveis a excluir", _k("nomes"))
    _rodape("excluir_variaveis", lambda: dict(nomes=nomes))


def _dlg_selecionar(b: Banco) -> None:
    st.caption(dlg.AJUDA["selecionar_casos"])
    modo = st.radio("Casos", ["Manter os casos que satisfazem a condição",
                              "Excluir os casos que satisfazem a condição"], key=_k("modo"))
    cond = _editor_expressao(b, "cond", "Condição", "ex.: Q1 = 1 OR Idade < 18")
    _rodape("selecionar_casos",
            lambda: dict(condicao=cond, modo="manter" if modo.startswith("Manter") else "excluir"))


def _dlg_ordenar(b: Banco) -> None:
    st.caption(dlg.AJUDA["ordenar_casos"])
    chaves = []
    for i in range(1, 4):
        c1, c2 = st.columns([3, 2])
        var = c1.selectbox(f"{i}ª variável", [""] + _opcoes(b), format_func=lambda c: _rot(b, c) if c else "—",
                           key=_k(f"v{i}"))
        sent = c2.radio("Ordem", ["A", "D"], horizontal=True, key=_k(f"d{i}"),
                        format_func=lambda x: "Crescente" if x == "A" else "Decrescente")
        if var:
            chaves.append((var, sent))
    _rodape("ordenar_casos", lambda: dict(chaves=chaves))


def _dlg_duplicados(b: Banco) -> None:
    st.caption(dlg.AJUDA["remover_duplicados"])
    chaves = _multi(b, "Variáveis que identificam o caso (ex.: telefone, ID)", _k("chaves"))
    manter = st.radio("Manter", ["primeiro", "ultimo"], horizontal=True, key=_k("manter"),
                      format_func=lambda x: "O primeiro caso de cada grupo" if x == "primeiro"
                      else "O último caso de cada grupo")
    _rodape("remover_duplicados", lambda: dict(chaves=chaves, manter=manter))


def _dlg_rotulo_var(b: Banco) -> None:
    st.caption(dlg.AJUDA["rotular_variavel"])
    var = _um(b, "Variável", _k("var"), padrao=_foco(b))
    rot = st.text_input("Rótulo (deixe em branco para remover)", value=b.rotulos_var.get(var or "", ""),
                        key=_k(f"rot_{var}"))
    _rodape("rotular_variavel", lambda: dict(nome=var, rotulo=rot))


def _dlg_rotulos_valor(b: Banco) -> None:
    st.caption(dlg.AJUDA["rotular_valores"])
    ref = _um(b, "Variável de referência (mostra os rótulos atuais)", _k("ref"), padrao=_foco(b))
    atual = b.rotulos_valor.get(ref or "", {})
    texto_atual = "\n".join(f"{_fmt_valor(k)} = {t}" for k, t in atual.items())
    alvo = st.multiselect("Aplicar às variáveis", _opcoes(b), default=[ref] if ref else [],
                          format_func=lambda c: _rot(b, c), key=_k(f"alvo_{ref}"))
    texto = st.text_area("Rótulos (um por linha: código = rótulo)", value=texto_atual, height=200,
                         key=_k(f"txt_{ref}"), placeholder="1 = Baixa\n2 = Média\n3 = Alta")
    modo = st.radio("Modo", ["substituir", "adicionar", "remover"], horizontal=True, key=_k("modo"),
                    format_func=lambda m: {"substituir": "Substituir todos (VALUE LABELS)",
                                           "adicionar": "Acrescentar/atualizar (ADD VALUE LABELS)",
                                           "remover": "Remover todos"}[m])

    def coletar() -> Dict[str, Any]:
        if not alvo:
            raise ErroTratamento("Escolha ao menos uma variável.")
        mapa: Dict[Any, str] = {}
        if modo != "remover":
            mapa = td.analisar_rotulos_valor(texto, not td.eh_texto(b.df[alvo[0]]))
        return dict(nomes=alvo, rotulos=mapa, modo=modo)

    _rodape("rotular_valores", coletar)


def _dlg_missing(b: Banco) -> None:
    st.caption(dlg.AJUDA["definir_missing"])
    foco = _foco(b)
    nomes = _multi(b, "Variáveis (todas do mesmo tipo)", _k("nomes"), default=[foco] if foco else [])
    modo = st.radio("Valores ausentes", list(dlg.MODOS_MISSING), format_func=dlg.MODOS_MISSING.get, key=_k("modo"))
    discretos, de, ate, extra = ["", "", ""], "", "", ""
    if modo == "discretos":
        discretos = [col.text_input(f"Valor {i}", key=_k(f"d{i}")) for i, col in enumerate(st.columns(3), 1)]
    elif modo == "faixa":
        c1, c2, c3 = st.columns(3)
        de = c1.text_input("Mínimo (vazio = menor valor)", key=_k("de"))
        ate = c2.text_input("Máximo (vazio = maior valor)", key=_k("ate"))
        extra = c3.text_input("Valor discreto (opcional)", key=_k("extra"))

    def coletar() -> Dict[str, Any]:
        if not nomes:
            raise ErroTratamento("Escolha ao menos uma variável.")
        spec = dlg.especificacao_missing(modo, discretos, de, ate, extra, texto=td.eh_texto(b.df[nomes[0]]))
        return dict(nomes=nomes, especificacao=spec)

    _rodape("definir_missing", coletar)


def _dlg_medida(b: Banco) -> None:
    st.caption(dlg.AJUDA["definir_medida"])
    foco = _foco(b)
    nomes = _multi(b, "Variáveis", _k("nomes"), default=[foco] if foco else [])
    medida = st.radio("Nível de medida", list(td.MEDIDAS), horizontal=True, key=_k("medida"),
                      format_func=lambda m: {"nominal": "Nominal", "ordinal": "Ordinal", "scale": "Escala"}[m])
    _rodape("definir_medida", lambda: dict(nomes=nomes, medida=medida))


def _dlg_add_casos(b: Banco) -> None:
    st.caption(dlg.AJUDA["adicionar_casos"])
    outro = _carregar_segundo_banco(_k("up2_casos"))
    if outro is None:
        _fechar()
        return
    usar = st.checkbox("Criar variável indicando a origem (0 = banco ativo, 1 = banco adicionado)", key=_k("usar"))
    nome = st.text_input("Nome da variável indicadora", value="Origem", key=_k("ind"))
    _rodape("adicionar_casos", lambda: dict(outro=outro, indicador=nome if usar else None))


def _dlg_add_vars(b: Banco) -> None:
    st.caption(dlg.AJUDA["adicionar_variaveis"])
    outro = _carregar_segundo_banco(_k("up2_vars"))
    if outro is None:
        _fechar()
        return
    comuns = [c for c in b.colunas if c.lower() in {x.lower() for x in outro.colunas}]
    chaves = st.multiselect("Chave(s) de ligação (variáveis presentes nos dois bancos)", comuns, key=_k("chaves"))
    tipo = st.radio("Quais casos ficam", ["consulta", "ambos"], key=_k("tipo"),
                    format_func=lambda t: "Só os casos do banco ativo (o segundo é uma tabela de consulta)"
                    if t == "consulta" else "Todos os casos dos dois bancos")
    trazer = st.multiselect("Variáveis a trazer (vazio = todas as novas)", outro.colunas, key=_k("trazer"))
    _rodape("adicionar_variaveis", lambda: dict(outro=outro, chaves=chaves, tipo=tipo, trazer=trazer or None))


def _dlg_agregar(b: Banco) -> None:
    st.caption(dlg.AJUDA["agregar"])
    quebras = _multi(b, "Variáveis de quebra (BREAK) — vazio = banco inteiro", _k("quebras"))
    st.markdown("**Resumos das variáveis**")
    n = _contagem("n_agr", 3)
    ags_campos = []
    for col, titulo in zip(st.columns([3, 2, 3]), ("Nova variável", "Função", "Variável")):
        col.caption(titulo)
    for i in range(1, n + 1):
        c1, c2, c3 = st.columns([3, 2, 3])
        ags_campos.append({
            "novo": c1.text_input(f"Nova variável {i}", key=_k(f"a_n{i}"), label_visibility="collapsed",
                                  placeholder="ex.: Idade_media"),
            "funcao": c2.selectbox(f"Função {i}", list(td.FUNCOES_AGREGACAO), key=_k(f"a_f{i}"),
                                   label_visibility="collapsed", help=" · ".join(f"{k}: {t}" for k, t in td.FUNCOES_AGREGACAO.items())),
            "variavel": c3.selectbox(f"Variável {i}", [""] + b.colunas, key=_k(f"a_v{i}"), label_visibility="collapsed",
                                     format_func=lambda c: _rot(b, c) if c else "— (não usada em N)"),
        })
    _botao_mais("n_agr", n)
    modo = st.radio("Resultado", ["adicionar", "substituir"], key=_k("modo"),
                    format_func=lambda m: "Acrescentar as variáveis agregadas a cada caso (MODE=ADDVARIABLES)"
                    if m == "adicionar" else "Substituir o banco por um banco agregado (uma linha por grupo)")

    def coletar() -> Dict[str, Any]:
        ags = [{"novo": str(r["novo"] or "").strip(), "funcao": r["funcao"], "variavel": r["variavel"] or ""}
               for r in ags_campos if str(r["novo"] or "").strip()]
        return dict(quebras=quebras, agregacoes=ags, modo=modo)

    _rodape("agregar", coletar)


# ----------------------------------------------------------------------------- janelas de arquivo e ajuda
def _dlg_salvar_sav(b: Banco) -> None:
    orig: Banco = st.session_state[K_ORIG]
    passos: List[Passo] = st.session_state[K_PASSOS]
    st.caption(dlg.AJUDA["salvar_sav"])
    nome = st.text_input("Nome do arquivo (sem extensão)", value=_nome_saida(orig), key=_k("nome_sav")).strip() \
        or _nome_saida(orig)
    from core.pyreadstat_seguro import pyreadstat
    st.caption("Gravação: " + ("escritor alternativo em Python (sem compressão) — o pyreadstat não está disponível aqui."
                               if pyreadstat.usando_leitor_puro else "pyreadstat (com compressão)."))
    assin = ("sav", len(passos), st.session_state.get(K_VER, 0), nome)
    saidas: Dict[str, Any] = st.session_state.setdefault(K_SAIDAS, {})
    if st.button("Gerar arquivo .sav", type="primary", use_container_width=True, key=_k("gerar_sav")):
        try:
            with st.spinner("Gravando o .sav..."):
                dados, avisos = td.gravar_sav(b)
            saidas["sav"] = {"assinatura": assin, "dados": dados, "avisos": avisos}
        except Exception as exc:
            st.error(f"Não foi possível gravar o .sav: {exc}")
    if saidas.get("sav", {}).get("assinatura") == assin:
        for a in saidas["sav"]["avisos"]:
            st.warning(a)
        _baixar("Baixar .sav", saidas["sav"]["dados"], f"{nome}.sav", "application/octet-stream",
                use_container_width=True)
    _fechar()


def _dlg_exportar(b: Banco) -> None:
    orig: Banco = st.session_state[K_ORIG]
    passos: List[Passo] = st.session_state[K_PASSOS]
    st.caption(dlg.AJUDA["exportar"])
    nome = st.text_input("Nome do arquivo (sem extensão)", value=_nome_saida(orig), key=_k("nome_exp")).strip() \
        or _nome_saida(orig)
    rotulos = st.checkbox("Mostrar os rótulos dos valores no lugar dos códigos", value=False, key=_k("rotulos"))
    saidas: Dict[str, Any] = st.session_state.setdefault(K_SAIDAS, {})
    assin = ("xlsx", len(passos), st.session_state.get(K_VER, 0), nome, rotulos)
    c1, c2 = st.columns(2)
    if c1.button("Gerar Excel", type="primary", use_container_width=True, key=_k("gerar_xlsx")):
        with st.spinner("Gerando o Excel..."):
            saidas["xlsx"] = {"assinatura": assin, "dados": td.exportar_excel(b, com_rotulos=rotulos)}
    if saidas.get("xlsx", {}).get("assinatura") == assin:
        with c1:
            _baixar("Baixar Excel (.xlsx)", saidas["xlsx"]["dados"], f"{nome}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c2:
        _baixar("Baixar CSV (;)", td.exportar_csv(b, com_rotulos=rotulos), f"{nome}.csv", "text/csv",
                use_container_width=True)
    _fechar()


def _dlg_salvar_sintaxe(b: Banco) -> None:
    orig: Banco = st.session_state[K_ORIG]
    passos: List[Passo] = st.session_state[K_PASSOS]
    st.caption(dlg.AJUDA["salvar_sintaxe"])
    if not passos:
        st.info("Ainda não há passos aplicados — a syntax fica disponível depois da primeira operação.")
    else:
        nome = _nome_saida(orig)
        sintaxe = td.gerar_sintaxe(orig.nome, passos, f"{nome}.sav")
        st.code(sintaxe, language="text")
        _baixar("Baixar syntax (.sps)", sintaxe.encode("utf-8-sig"), f"{nome}.sps", "text/plain",
                use_container_width=True)
    _fechar()


def _dlg_ajuda_expressoes(b: Banco) -> None:
    st.markdown(dlg.AJUDA_EXPRESSOES)
    _fechar()


def _dlg_ajuda_geral(b: Banco) -> None:
    st.markdown(dlg.AJUDA_GERAL)
    _fechar()


# id do menu -> (título da janela, largura, função)
JANELAS: Dict[str, Tuple[str, str, Callable[[Banco], None]]] = {
    "salvar_sav": ("Salvar como .sav", "medium", _dlg_salvar_sav),
    "exportar": ("Exportar para Excel / CSV", "medium", _dlg_exportar),
    "salvar_sintaxe": ("Salvar syntax (.sps)", "large", _dlg_salvar_sintaxe),
    "renomear_variavel": ("Renomear variável", "medium", _dlg_renomear),
    "excluir_variaveis": ("Excluir variáveis", "medium", _dlg_excluir_vars),
    "ordenar_casos": ("Ordenar casos", "medium", _dlg_ordenar),
    "selecionar_casos": ("Selecionar casos", "large", _dlg_selecionar),
    "remover_duplicados": ("Remover casos duplicados", "medium", _dlg_duplicados),
    "agregar": ("Agregar dados", "large", _dlg_agregar),
    "adicionar_casos": ("Adicionar casos", "medium", _dlg_add_casos),
    "adicionar_variaveis": ("Adicionar variáveis", "large", _dlg_add_vars),
    "rotular_variavel": ("Rótulo da variável", "medium", _dlg_rotulo_var),
    "rotular_valores": ("Rótulos dos valores", "large", _dlg_rotulos_valor),
    "definir_missing": ("Valores ausentes", "medium", _dlg_missing),
    "definir_medida": ("Nível de medida", "medium", _dlg_medida),
    "compute": ("Calcular variável", "large", _dlg_compute),
    "recode_mesmas": ("Recodificar nas mesmas variáveis", "large", _dlg_recode_mesmas),
    "recode_outras": ("Recodificar em outras variáveis", "large", _dlg_recode_outras),
    "ajuda_expressoes": ("Funções e operadores", "large", _dlg_ajuda_expressoes),
    "ajuda_geral": ("Como usar este módulo", "medium", _dlg_ajuda_geral),
}


def _abrir_janela(ident: str, b: Banco) -> None:
    titulo, largura, funcao = JANELAS[ident]
    st.dialog(titulo, width=largura)(funcao)(b)


# ----------------------------------------------------------------------------- faixas de comandos
def _titulo(nome_arquivo: str) -> None:
    st.markdown(f'<div class="td-titulo">{html.escape(nome_arquivo)} [Conjunto de dados] — Tratamento de Dados</div>',
                unsafe_allow_html=True)


def _faixas() -> Optional[str]:
    """Abas por tarefa; em cada uma, grupos com os comandos sempre visíveis. Devolve o id do comando clicado."""
    escolhido: Optional[str] = None
    with _caixa("td_faixa"):
        abas = st.tabs([nome for nome, _ in dlg.FAIXAS])
        for aba, (_, grupos) in zip(abas, dlg.FAIXAS):
            with aba:
                colunas = st.columns([1] * len(grupos) + [max(0, 4 - len(grupos)) or 0.01], gap="medium")
                for col, (grupo, comandos) in zip(colunas, grupos):
                    with col:
                        st.markdown(f'<div class="td-grupo">{html.escape(grupo)}</div>', unsafe_allow_html=True)
                        for ident, icone, rotulo, descricao in comandos:
                            if st.button(rotulo, key=f"td_c_{ident}", icon=f":material/{icone}:", help=descricao,
                                         use_container_width=True):
                                escolhido = ident
                            st.caption(descricao)
    return escolhido


def _acesso_rapido(tem_passos: bool) -> Optional[str]:
    """Desfazer sempre à vista, acima das abas."""
    escolhido: Optional[str] = None
    c1, c2, _ = st.columns([1.3, 1.1, 6])
    if c1.button("Desfazer", key="td_q_desfazer", icon=":material/undo:", disabled=not tem_passos,
                 help="Desfaz o último passo", use_container_width=True):
        escolhido = "desfazer"
    if c2.button("Desfazer tudo", key="td_q_desfazer_tudo", icon=":material/restart_alt:", disabled=not tem_passos,
                 help="Volta ao banco original", use_container_width=True):
        escolhido = "desfazer_tudo"
    return escolhido


def _barra_de_status(orig: Banco, b: Banco, passos: List[Passo]) -> None:
    partes = ['<span><span class="td-ok">●</span> Processador pronto</span>',
              f"<span>{html.escape(b.nome)}</span>",
              f"<span>{f'{len(b.df):,}'.replace(',', '.')} casos × {len(b.df.columns)} variáveis</span>",
              f"<span>{len(passos)} passo(s) aplicado(s)</span>"]
    if passos:
        partes.append(f"<span>Último: {html.escape(passos[-1].descricao[:90])}</span>")
    st.markdown(f'<div class="td-status">{"".join(partes)}</div>', unsafe_allow_html=True)


# ----------------------------------------------------------------------------- visualizações (abas de baixo)
def _visao_dados(b: Banco) -> None:
    c1, c2, _ = st.columns([1, 2, 3])
    n = c1.number_input("Linhas a mostrar", 5, 500, 100, step=50, key="td_linhas")
    rotulos = c2.checkbox("Mostrar os rótulos dos valores (Exibir ▸ Rótulos dos valores)", value=False,
                          key="td_ver_rotulos")
    st.markdown(dlg.tabela_html_dados(b, int(n), rotulos), unsafe_allow_html=True)
    st.caption(f"Mostrando {min(int(n), len(b.df))} de {len(b.df)} casos. Passe o mouse sobre o nome da coluna "
               "para ver o rótulo da variável.")


def _visao_variaveis(b: Banco) -> Optional[str]:
    """Variable View. Devolve o id da janela pedida pelos botões de edição da variável escolhida."""
    st.markdown(dlg.tabela_html_variaveis(b), unsafe_allow_html=True)
    c0, *botoes = st.columns([3.2, 1.2, 1.2, 1.2, 1.4, 1.2, 0.8])[:1 + len(dlg.ACOES_DA_VARIAVEL)]
    var = c0.selectbox("Variável escolhida", b.colunas, format_func=lambda c: _rot(b, c), key="td_foco_sel",
                       label_visibility="collapsed")
    pedido: Optional[str] = None
    for col, (rotulo, ident) in zip(botoes, dlg.ACOES_DA_VARIAVEL):
        if col.button(rotulo, key=f"td_vv_{ident}", use_container_width=True):
            pedido = ident
    if pedido:
        st.session_state[K_FOCO] = var
    st.caption("Escolha uma variável e clique em Nome, Rótulo, Valores, Ausentes ou Medida para editá-la.")
    return pedido


def _visao_syntax(orig: Banco, passos: List[Passo]) -> None:
    st.code(td.gerar_sintaxe(orig.nome, passos, f"{_nome_saida(orig)}.sav"), language="text")
    if not passos:
        st.info("Nenhuma operação aplicada ainda. Use os menus Transformar e Dados.")
        return
    st.markdown("**Passos aplicados**")
    for i, p in enumerate(passos, 1):
        st.markdown(f"{i}. {p.descricao}")
        for a in p.avisos:
            st.warning(a)
    c1, c2, _ = st.columns([1, 1, 3])
    if c1.button("Desfazer último passo", key="td_sy_desfazer"):
        _desfazer(1)
    if c2.button("Desfazer tudo", key="td_sy_desfazer_tudo"):
        _desfazer(None)


# ----------------------------------------------------------------------------- módulo
def modulo_tratamento_dados() -> None:
    st.markdown(CSS, unsafe_allow_html=True)
    st.header("Tratamento de Dados")

    carregado = K_ATUAL in st.session_state
    titulo = f"📂 {st.session_state[K_ORIG].nome} — trocar de arquivo" if carregado else "📂 Abrir arquivo de dados (.sav)"
    with st.expander(titulo, expanded=not carregado):
        up = st.file_uploader("Carregue o banco SPSS (.sav)", type=["sav"], key="td_upload")
    if up is None:
        _estado_vazio()
        st.session_state.pop(K_ASSIN, None)
        st.info("Envie um arquivo .sav para começar.")
        return
    try:
        _carregar(up)
    except Exception as exc:
        st.error(f"Não foi possível ler o arquivo SPSS: {exc}")
        return

    orig: Banco = st.session_state[K_ORIG]
    b: Banco = st.session_state[K_ATUAL]
    passos: List[Passo] = st.session_state[K_PASSOS]

    _titulo(b.nome)
    rapido = _acesso_rapido(bool(passos))
    escolhido = rapido or _faixas()
    if escolhido:
        st.session_state.pop(K_FOCO, None)           # abriu pelo menu/ferramenta: sem variável pré-selecionada
    if escolhido == "desfazer":
        _desfazer(1)
    elif escolhido == "desfazer_tudo":
        _desfazer(None)

    ultimo: Optional[Passo] = st.session_state.pop(K_MSG, None)
    if ultimo is not None:
        st.success(f"Aplicado: {ultimo.descricao}")
        for a in ultimo.avisos:
            st.warning(a)
        with st.expander("Syntax deste passo", expanded=False):
            st.code(ultimo.sintaxe, language="text")

    area = st.container()                     # a grade fica em cima; as abas de baixo vêm depois (como no SPSS)
    visao = st.segmented_control("Visualização", list(dlg.VISOES), default=dlg.VISOES[0], key="td_visao",
                                 label_visibility="collapsed") or dlg.VISOES[0]
    _barra_de_status(orig, b, passos)
    with area:
        if visao == dlg.VISOES[1]:
            pedido = _visao_variaveis(b)
            escolhido = pedido or escolhido
        elif visao == dlg.VISOES[2]:
            _visao_syntax(orig, passos)
        else:
            _visao_dados(b)

    if escolhido and escolhido not in dlg.ACOES_DIRETAS:
        _abrir_janela(escolhido, b)
