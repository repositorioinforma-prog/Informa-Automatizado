"""
Ponte Python <-> página original do Relatória Analítica, dentro do Streamlit.

As telas do app desktop eram páginas HTML/JS que chamavam o Python pelo
QWebChannel (`QtBridge.metodo(...)`). Aqui cada tela original roda num
componente Streamlit (pasta `componentes/<nome>/index.html`) e o
`componentes/ponte_qt.js` imita o QWebChannel. Este módulo é o outro lado:

- `PonteBase`: classe-base das "bridges". Os métodos públicos são os
  mesmos `@Slot` do original; em vez de Qt, usam:
    self.emitir(sinal, *args)       -> Signal.emit
    self.run_js(codigo)             -> page().runJavaScript
    self.aviso(texto, nivel)        -> QMessageBox
    self.download(nome, dados, mime)-> QFileDialog "Salvar como" + gravar
    self.notificar(texto, icone)    -> pop-up (st.toast) na tela do Streamlit
  Métodos que abriam o seletor de arquivo recebem `arquivo={"name","b64"}`.
- `rodar_ferramenta(...)`: processa os pedidos pendentes da página e
  desenha o componente com os eventos novos. Para pôr algo ENTRE as duas
  etapas (ex.: a barra lateral), use `obter_ponte`, `processar_pedidos`,
  `mostrar_notificacoes` e `desenhar_componente` separadamente.
"""
from __future__ import annotations

import base64
import hashlib
import os
import re
from typing import Any, Dict, List, Optional

import streamlit as st
import streamlit.components.v1 as components

_PASTA = os.path.join(os.path.dirname(os.path.dirname(__file__)), "componentes")
_componentes: Dict[str, Any] = {}

_MAX_EVENTOS = 60


def _assinatura_pasta(pasta: str) -> str:
    """Hash do conteúdo dos arquivos da pasta do componente."""
    h = hashlib.md5()
    for raiz, _dirs, arqs in sorted(os.walk(pasta)):
        for a in sorted(arqs):
            caminho = os.path.join(raiz, a)
            h.update(os.path.relpath(caminho, pasta).encode("utf-8"))
            with open(caminho, "rb") as f:
                h.update(f.read())
    return h.hexdigest()[:8]


def _componente(nome: str):
    """Declara o componente com o hash dos arquivos no NOME: o endereço muda
    sempre que a tela muda, então o navegador nunca usa uma cópia antiga."""
    pasta = os.path.join(_PASTA, nome)
    chave = f"{nome}_{_assinatura_pasta(pasta)}"
    if chave not in _componentes:
        _componentes[chave] = components.declare_component(chave, path=pasta)
    return _componentes[chave]


def versao_tela_no_servidor(nome: str) -> str:
    """Lê <meta name="versao-tela"> do index.html que está no servidor."""
    try:
        with open(os.path.join(_PASTA, nome, "index.html"), encoding="utf-8") as f:
            m = re.search(r'<meta\s+name="versao-tela"\s+content="([^"]*)"', f.read())
        return m.group(1) if m else ""
    except OSError:
        return ""


def conferir_versao_tela(ponte: "PonteBase", nome: str, esperada: str) -> bool:
    """Avisa quando a pasta componentes/<nome> ou o navegador estão com a tela
    antiga (o Python novo com a tela velha quebra o layout). True = tudo ok."""
    no_servidor = versao_tela_no_servidor(nome)
    if no_servidor != esperada:
        st.error(
            f"Os arquivos da pasta **componentes/{nome}/** estão desatualizados no servidor "
            f"(versão da tela: {no_servidor or 'antiga'}; esperada: {esperada}). "
            "Envie a pasta **componentes/** inteira junto com a pasta **analises/** para o repositório "
            "e reinicie o app."
        )
        return False
    no_navegador = getattr(ponte, "tela_navegador", None)
    if no_navegador is not None and no_navegador != esperada:
        st.warning(
            "O navegador ainda está com a tela antiga desta ferramenta. "
            "Recarregue a página com **Ctrl + F5**."
        )
        return False
    return True


class PonteBase:
    """Base das bridges. Subclasses implementam os métodos chamados pela página
    (mesmos nomes do `QtBridge` original) e, opcionalmente, `estado_inicial()`."""

    #: métodos que, no original, abriam o QFileDialog para ESCOLHER arquivo
    ARQUIVOS: Dict[str, Dict[str, str]] = {}

    def __init__(self):
        self.eventos: List[Dict[str, Any]] = []
        self.n_evento = 0
        self.acks: Dict[str, int] = {}
        self.ultimo_download: Optional[Dict[str, Any]] = None
        self.notificacoes: List[Dict[str, str]] = []

    # ---------------- saída para a página
    def _evento(self, ev: Dict[str, Any]) -> None:
        self.n_evento += 1
        ev["n"] = self.n_evento
        self.eventos.append(ev)
        # downloads são pesados: mantém só os 2 mais recentes na lista
        downloads = [e for e in self.eventos if e.get("tipo") == "download"]
        for velho in downloads[:-2]:
            self.eventos.remove(velho)
        if len(self.eventos) > _MAX_EVENTOS:
            self.eventos = self.eventos[-_MAX_EVENTOS:]

    def emitir(self, sinal: str, *args) -> None:
        self._evento({"tipo": "sinal", "nome": sinal, "args": list(args)})

    def run_js(self, codigo: str) -> None:
        self._evento({"tipo": "js", "codigo": codigo})

    def aviso(self, texto: str, nivel: str = "info") -> None:
        self._evento({"tipo": "aviso", "texto": texto, "nivel": nivel})

    def download(self, nome: str, dados: bytes, mime: str) -> None:
        self.ultimo_download = {"nome": nome, "dados": dados, "mime": mime}
        self._evento({"tipo": "download", "nome": nome, "mime": mime,
                      "b64": base64.b64encode(dados).decode("ascii")})

    def notificar(self, texto: str, icone: str = "ℹ️") -> None:
        """Pop-up rápido na tela do Streamlit (fica visível mesmo com a
        página rolada, ao contrário dos avisos dentro do componente)."""
        if not hasattr(self, "notificacoes"):
            self.notificacoes = []
        self.notificacoes.append({"texto": texto, "icone": icone})

    # ---------------- entrada
    def estado_inicial(self) -> None:
        """Chamado quando a página (re)abre: reenviar o estado atual."""

    def despachar(self, msg: Dict[str, Any]) -> None:
        metodo = str(msg.get("method") or "")
        if metodo == "__init__":
            self.estado_inicial()
            return
        if metodo.startswith("_") or not hasattr(self, metodo):
            self.aviso(f"Ação não suportada: {metodo}", "erro")
            return
        args = list(msg.get("args") or [])
        kwargs = {}
        if "file" in msg:
            f = msg["file"] or {}
            kwargs["arquivo"] = {"name": f.get("name", ""), "dados": base64.b64decode(f.get("b64") or "")}
        try:
            getattr(self, metodo)(*args, **kwargs)
        except Exception as e:  # noqa: BLE001 — o original mostrava QMessageBox
            self.aviso(f"Erro em {metodo}:\n{e}", "erro")


def obter_ponte(chave_estado: str, fabrica) -> "PonteBase":
    """Bridge guardada em st.session_state (criada na primeira vez).

    Quando o código é atualizado com o app aberto, o Streamlit recarrega os
    módulos mas mantém na sessão o objeto criado pela versão ANTIGA da classe
    (sem os métodos novos -> "Ação não suportada: ..."). Nesse caso a ponte é
    recriada com a classe nova, aproveitando os dados já carregados.
    """
    ponte = st.session_state.get(chave_estado)
    if ponte is None:
        ponte = fabrica()
        st.session_state[chave_estado] = ponte
    elif isinstance(fabrica, type) and type(ponte) is not fabrica:
        nova = fabrica()
        try:
            nova.__dict__.update(vars(ponte))  # mantém banco, resultados, filas...
        except TypeError:
            pass
        ponte = nova
        st.session_state[chave_estado] = ponte
    return ponte


def mostrar_notificacoes(ponte: "PonteBase") -> None:
    """Mostra (e esvazia) os pop-ups pedidos com `ponte.notificar`."""
    fila = getattr(ponte, "notificacoes", None) or []
    ponte.notificacoes = []
    toast = getattr(st, "toast", None)
    for n in fila:
        if toast is not None:
            try:
                toast(n["texto"], icon=n.get("icone") or None)
                continue
            except Exception:  # noqa: BLE001 — versões antigas / ícone inválido
                pass
        # sem st.toast: usa a faixa de aviso dentro da página
        ponte.aviso(n["texto"], "info")


def processar_pedidos(ponte: "PonteBase", key: str) -> None:
    """Executa os pedidos que a página enviou desde a última execução."""
    pedido = st.session_state.get(key)
    if isinstance(pedido, dict) and pedido.get("nonce"):
        # versão da tela que o NAVEGADOR está rodando (páginas antigas não mandam)
        ponte.tela_navegador = str(pedido.get("tela") or "antiga")
        # eventos que a página já executou não precisam ir de novo (downloads e
        # dados do Excel podem ter MB). Uma página nova ignora os antigos e
        # pede o estado com "__init__", então nada se perde.
        try:
            visto = int(pedido.get("visto") or 0)
        except (TypeError, ValueError):
            visto = 0
        if visto:
            ponte.eventos = [e for e in ponte.eventos if e["n"] > visto]
        nonce = str(pedido["nonce"])
        ultimo = ponte.acks.get(nonce, 0)
        for msg in sorted(pedido.get("pendentes") or [], key=lambda m: m.get("n", 0)):
            n = int(msg.get("n") or 0)
            if n > ultimo:
                ponte.despachar(msg)
                ultimo = n
        ponte.acks[nonce] = ultimo
        # guarda só os acks mais recentes (cada abertura da página gera um nonce)
        if len(ponte.acks) > 20:
            for k in list(ponte.acks)[:-20]:
                ponte.acks.pop(k, None)


def desenhar_componente(ponte: "PonteBase", nome_componente: str, *, altura: Any = 900, key: str) -> None:
    _componente(nome_componente)(
        eventos=ponte.eventos,
        ultimo_evento=ponte.n_evento,
        acks=ponte.acks,
        arquivos=ponte.ARQUIVOS,
        altura=altura,
        key=key,
        default=None,
    )


def rodar_ferramenta(nome_componente: str, chave_estado: str, fabrica, *, altura: int = 900, key: str):
    """Processa os pedidos da página e desenha o componente.

    nome_componente: pasta em componentes/ com o index.html da tela.
    chave_estado: onde a bridge fica guardada em st.session_state.
    fabrica: callable que cria a bridge (subclasse de PonteBase).
    """
    ponte = obter_ponte(chave_estado, fabrica)
    processar_pedidos(ponte, key)
    mostrar_notificacoes(ponte)
    desenhar_componente(ponte, nome_componente, altura=altura, key=key)
    return ponte
