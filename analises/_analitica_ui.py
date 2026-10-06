"""
Peças de interface compartilhadas pelas ferramentas analíticas portadas
do "Relatória Analítica" (Cruzamentos, Análise Residual, Matriz de
Correlação, Mapa de Calor).
"""
from __future__ import annotations

import html
from typing import Callable, Dict, List, Optional

import streamlit as st

from core.analitica_dados import BancoCarregado, carregar_banco


def upload_banco(
    prefixo: str,
    rotulo: str = "Importar banco (CSV / XLSX / SAV)",
    *,
    tipos=("csv", "xlsx", "xls", "sav"),
    user_missing: bool = False,
    remover_colunas_vazias: bool = False,
    ao_trocar: Optional[Callable[[], None]] = None,
) -> Optional[BancoCarregado]:
    """Upload + leitura com cache na sessão (só relê quando o arquivo
    muda). `ao_trocar` é chamado quando um arquivo NOVO é carregado — é
    onde cada tela zera os resultados/fila, como o pickData original."""
    arquivo = st.file_uploader(rotulo, type=list(tipos), key=f"{prefixo}_upload")
    chave_estado = f"{prefixo}_banco"
    chave_id = f"{prefixo}_banco_id"

    if arquivo is None:
        return st.session_state.get(chave_estado)

    ident = (arquivo.name, arquivo.size)
    if st.session_state.get(chave_id) != ident:
        try:
            with st.spinner("Lendo o banco..."):
                banco = carregar_banco(
                    arquivo.getvalue(),
                    arquivo.name,
                    user_missing=user_missing,
                    remover_colunas_vazias=remover_colunas_vazias,
                )
        except Exception as e:  # noqa: BLE001
            st.error(f"Falha ao carregar dados:\n\n{e}")
            return None
        st.session_state[chave_estado] = banco
        st.session_state[chave_id] = ident
        if ao_trocar:
            ao_trocar()
    banco = st.session_state.get(chave_estado)
    if banco is not None:
        st.caption(f"📄 {banco.nome_arquivo} — {len(banco.df):,} casos · {banco.df.shape[1]} variáveis".replace(",", "."))
    return banco


def rotulo_var(nome: str, labels: Dict[str, str]) -> str:
    """Nome de exibição (rótulo do SPSS quando existir) — o `dispVar` do JS."""
    lab = (labels.get(nome) or "").strip()
    return lab if lab else nome


def formato_var(labels: Dict[str, str]) -> Callable[[str], str]:
    """format_func para selectbox/multiselect: 'Rótulo  ·  nome'."""
    def _f(nome: str) -> str:
        if not nome:
            return "—"
        lab = (labels.get(nome) or "").strip()
        return f"{lab}  ·  {nome}" if lab and lab != nome else nome
    return _f


def esc(s) -> str:
    return html.escape("" if s is None else str(s))


CSS_TABELA = """
<style>
.ra-wrap{overflow-x:auto;margin:4px 0 8px}
.ra-tab{border-collapse:collapse;font-size:13px;min-width:100%}
.ra-tab th,.ra-tab td{border:1px solid #d0d7de;padding:5px 8px;text-align:center;white-space:nowrap}
.ra-tab th{background:#f2f4f7;font-weight:700}
.ra-tab td.rowlab,.ra-tab th.rowlab{text-align:left;white-space:normal;min-width:180px}
.ra-tab tr.base td{font-weight:700;background:#fafbfc}
.ra-pill{display:inline-block;padding:2px 9px;border-radius:999px;border:1px solid #d0d7de;
         background:#fff;font-size:12px;font-weight:700;margin-right:6px}
</style>
"""


def css_tabelas() -> None:
    st.markdown(CSS_TABELA, unsafe_allow_html=True)


def lista_ordenavel(
    chave: str,
    itens: List[str],
    *,
    detalhes: Optional[List[str]] = None,
    permitir_remover: bool = False,
    permitir_renomear: bool = False,
):
    """Editor de ordem (substitui os modais com ↑/↓ do original).

    Mostra uma tabela editável com a coluna "Ordem" (e, se pedido,
    "Manter" e "Novo nome"). Devolve uma lista de tuplas
    (índice_original, novo_nome, manter) já na nova ordem — ou None se o
    editor não pôde ser lido.
    """
    import pandas as pd

    dados = {"Ordem": list(range(1, len(itens) + 1)), "Item": itens}
    if detalhes is not None:
        dados["Detalhe"] = detalhes
    if permitir_renomear:
        dados["Novo nome"] = list(itens)
    if permitir_remover:
        dados["Manter"] = [True] * len(itens)
    df = pd.DataFrame(dados)

    config = {
        "Ordem": st.column_config.NumberColumn("Ordem", min_value=1, step=1, width="small"),
        "Item": st.column_config.TextColumn("Item", disabled=True),
    }
    if detalhes is not None:
        config["Detalhe"] = st.column_config.TextColumn("Detalhe", disabled=True)
    if permitir_remover:
        config["Manter"] = st.column_config.CheckboxColumn("Manter", width="small")

    editado = st.data_editor(
        df, key=chave, hide_index=True,
        column_config=config, num_rows="fixed",
    )
    ordem = []
    for i in range(len(itens)):
        try:
            pos = float(editado.loc[i, "Ordem"])
        except Exception:
            pos = float(i + 1)
        nome = str(editado.loc[i, "Novo nome"]).strip() if permitir_renomear else itens[i]
        manter = bool(editado.loc[i, "Manter"]) if permitir_remover else True
        # empate na "Ordem": mantém a ordem original entre os empatados
        ordem.append((pos, i, nome or itens[i], manter))
    ordem.sort(key=lambda x: (x[0], x[1]))
    return [(i, nome, manter) for _, i, nome, manter in ordem]
