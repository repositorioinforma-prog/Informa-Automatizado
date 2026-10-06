"""
Pilha/fila de exportação na barra lateral (layout "cards expansíveis"),
usada pelos Cruzamentos e pelo Mapa de Calor.

- `componentes/relatoria_pilha/index.html` desenha o painel e os cards;
- a lixeira de um card manda {"nonce", "acoes":[{"id","acao":"remover","sig"}]};
  cada ação roda UMA vez (o ack por nonce volta para o componente);
- "Exportar" e "Limpar" são botões nativos do Streamlit, estilizados aqui.
"""
from __future__ import annotations

from typing import Any, Callable, Dict, List, Optional

import streamlit as st

from analises._ponte_relatoria import _componente

# os seletores pegam qualquer botão cuja key termine em _pilha_exportar / _pilha_limpar
CSS_PILHA = """
<style>
[data-testid="stSidebar"] [class*="_pilha_exportar"] button{
  width:100%; min-height:52px; border-radius:12px; font-size:1.05rem; font-weight:700;
  background:#2563eb !important; color:#fff !important; border:none !important;
}
[data-testid="stSidebar"] [class*="_pilha_exportar"] button:hover{ background:#1d4ed8 !important; }
[data-testid="stSidebar"] [class*="_pilha_limpar"] button{
  width:100%; min-height:48px; border-radius:12px; font-size:1.02rem; font-weight:700;
  background:#1b2434 !important; border:1px solid #2e3b55 !important;
}
[data-testid="stSidebar"] [class*="_pilha_limpar"] button,
[data-testid="stSidebar"] [class*="_pilha_limpar"] button *{ color:#f05252 !important; fill:#f05252 !important; }
[data-testid="stSidebar"] [class*="_pilha_limpar"] button:hover{ background:#2a1f2a !important; }
</style>
"""


def processar_acoes(ponte: Any, valor: Any, remover: Callable[[str], None], atributo: str = "acks_pilha") -> Dict[str, int]:
    """Executa as ações (lixeira) enviadas pelos cards; devolve os acks."""
    acks: Dict[str, int] = getattr(ponte, atributo, None) or {}
    if isinstance(valor, dict) and valor.get("nonce"):
        nonce = str(valor["nonce"])
        ultimo = acks.get(nonce, 0)
        for a in sorted(valor.get("acoes") or [], key=lambda x: int(x.get("id") or 0)):
            n = int(a.get("id") or 0)
            if n <= ultimo:
                continue
            if a.get("acao") == "remover":
                remover(str(a.get("sig") or ""))
            ultimo = n
        acks[nonce] = ultimo
        if len(acks) > 20:
            for k in list(acks)[:-20]:
                acks.pop(k, None)
    setattr(ponte, atributo, acks)
    return acks


def botao(fn, rotulo: str, icone: str, **kw):
    """st.button/st.download_button com ícone Material quando a versão permite."""
    try:
        return fn(rotulo, icon=icone, **kw)
    except TypeError:
        return fn(rotulo, **kw)


def pilha_lateral(
    ponte: Any,
    *,
    prefixo: str,
    itens: Any,
    remover: Callable[[str], None],
    titulo: str,
    vazio_msg: str,
    exportar: Any = None,
    limpar: Optional[Callable[[], None]] = None,
    rotulo_limpar: str = "Limpar pilha",
    subtitulo: Any = "",
    subtitulo_tipo: str = "",
) -> None:
    """Desenha a pilha na barra lateral.

    prefixo: base das keys (ex.: "cz" -> cz_pilha_comp, cz_pilha_exportar...).
    itens: cards (ver componentes/relatoria_pilha/index.html) — lista ou
        função; use função para refletir uma remoção feita nesta execução.
    exportar: {"rotulo", "dados" (bytes ou função), "arquivo", "mime"} ou
        função que devolve esse dicionário.
    """
    chave = f"{prefixo}_pilha_comp"
    acks = processar_acoes(ponte, st.session_state.get(chave), remover)
    # a lixeira pode ter removido algo agora: os itens precisam refletir isso
    itens = itens() if callable(itens) else itens
    exportar = exportar() if callable(exportar) else exportar
    subtitulo = subtitulo() if callable(subtitulo) else subtitulo
    with st.sidebar:
        st.markdown(CSS_PILHA, unsafe_allow_html=True)
        st.markdown("---")
        _componente("relatoria_pilha")(
            itens=itens, ack=acks, titulo=titulo, subtitulo=subtitulo, subtitulo_tipo=subtitulo_tipo,
            vazio_msg=vazio_msg, key=chave, default=None,
        )
        if not itens:
            return
        if exportar:
            try:
                dados = exportar["dados"]() if callable(exportar["dados"]) else exportar["dados"]
            except Exception as e:  # noqa: BLE001
                st.error(f"Falha ao montar o arquivo: {e}")
                dados = None
            if dados is not None:
                botao(st.download_button, exportar["rotulo"], ":material/download:", data=dados,
                      file_name=exportar["arquivo"], mime=exportar["mime"], key=f"{prefixo}_pilha_exportar")
        if limpar:
            botao(st.button, rotulo_limpar, ":material/delete:", key=f"{prefixo}_pilha_limpar", on_click=limpar)
