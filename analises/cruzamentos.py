"""
Cruzamentos (Tabelas) — a tela ORIGINAL do Relatória Analítica dentro do
Streamlit.

A interface (HTML/CSS/JS) é a mesma do app desktop, sem alterações:
componentes/relatoria_cruzamentos/index.html. Ela conversa com a
`PonteCruzamentos` abaixo, que tem os mesmos métodos da `QtBridge` do
original (crosstabs_page.py). Os cálculos e a exportação Excel estão em
core/cruzamentos_math.py (copiados do original, com paridade conferida).

O que muda em relação ao desktop:
- "Carregar dados" abre o seletor de arquivo do navegador (no lugar da
  janela do Windows);
- a pilha de exportação fica na barra lateral do Streamlit (lista,
  remover item, Exportar Excel, Limpar pilha);
- adicionar à pilha mostra um pop-up (st.toast); o mesmo cruzamento
  (linhas × colunas, peso, % e NA) não entra duas vezes;
- "Exportar (1)" baixa o arquivo pelo navegador;
- os avisos (QMessageBox) aparecem como faixas no canto da página.
- Na página: lista de variáveis mais larga, Marcar tudo (verde, com
  contador) / Desmarcar (vermelho), e colunas com base < 30 ou "Branco/Nulo",
  "Não irá votar", "Não sabe" já vêm removidas (restauráveis).
"""
from __future__ import annotations

import hashlib
import json
from typing import Any, Dict, List, Optional, Tuple

import streamlit as st

from analises._pilha_lateral import pilha_lateral, processar_acoes
from analises._ponte_relatoria import (
    PonteBase, conferir_versao_tela, desenhar_componente, mostrar_notificacoes, obter_ponte,
    processar_pedidos,
)
from analises.tutoriais_analitica import botao_tutorial
from core.analitica_dados import carregar_banco
from core.cruzamentos_math import calcular_cruzamentos, exportar_excel

#: precisa bater com <meta name="versao-tela"> de componentes/relatoria_cruzamentos/index.html
VERSAO_TELA = "17"

_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class PonteCruzamentos(PonteBase):
    """Equivalente à `QtBridge` de crosstabs_page.py."""

    ARQUIVOS = {"pickData": {"accept": ".csv,.xlsx,.xls,.sav"}}

    def __init__(self):
        super().__init__()
        self.banco = None
        self.queue: List[Dict[str, Any]] = []
        self.ultimas_tabelas: List[Dict[str, Any]] = []

    # ---------------- estado ao (re)abrir a página
    def estado_inicial(self) -> None:
        if self.banco is not None:
            self._emitir_dados()
            if self.ultimas_tabelas:
                self.emitir("crosstabsReady", json.dumps({"tables": self.ultimas_tabelas}, ensure_ascii=False))
        self._push_queue_state()

    @staticmethod
    def assinatura(t: Dict[str, Any]) -> str:
        """Identifica a CONFIGURAÇÃO do cruzamento (igual ao sigOf da página)."""
        return "|".join([
            str(t.get("row_var") or ""), str(t.get("col_var") or ""), str(t.get("weight_var") or ""),
            str(t.get("pct_mode") or "col"), "1" if t.get("include_na") else "0",
        ])

    @staticmethod
    def titulo(t: Dict[str, Any]) -> str:
        return (t.get("title") or t.get("key") or "").strip()

    def _posicao_na_pilha(self, t: Dict[str, Any]) -> Optional[int]:
        sig = self.assinatura(t)
        for i, it in enumerate(self.queue):
            if self.assinatura(it) == sig:
                return i
        return None

    def _push_queue_state(self) -> None:
        items = [{"title": self.titulo(it), "sig": self.assinatura(it)} for it in self.queue]
        self.emitir("queueState", json.dumps({"items": items}, ensure_ascii=False))

    def _tentar_adicionar(self, t: Dict[str, Any]) -> Tuple[bool, Optional[int]]:
        pos = self._posicao_na_pilha(t)
        if pos is not None:
            return False, pos
        self.queue.append(t)
        return True, None

    def _emitir_dados(self) -> None:
        df = self.banco.df
        cols = [{"name": str(c), "dtype": str(dtype)} for c, dtype in df.dtypes.items()]
        labels = {str(c): (self.banco.var_labels.get(str(c)) or "") for c in df.columns}
        payload = {"file_info": self.banco.nome_arquivo, "columns": cols, "labels": labels}
        self.emitir("dataLoaded", json.dumps(payload, ensure_ascii=False))

    # ---------------- @Slot pickData
    def pickData(self, arquivo=None) -> None:
        if not arquivo:
            return
        try:
            self.banco = carregar_banco(arquivo["dados"], arquivo["name"])
        except Exception as e:  # noqa: BLE001
            self.aviso(f"Falha ao carregar dados:\n{e}", "erro")
            return
        self.ultimas_tabelas = []
        self._emitir_dados()

    # ---------------- @Slot computeCrosstabs
    def computeCrosstabs(self, rowVar: str, colVarsJson: str, weightVar: str, includeNA: bool, pctMode: str) -> None:
        if self.banco is None:
            self.aviso("Carregue um arquivo primeiro.", "aviso")
            return
        try:
            col_vars = json.loads(colVarsJson)
            tabelas = calcular_cruzamentos(
                self.banco.df, self.banco.var_labels, self.banco.value_labels,
                rowVar, col_vars, weightVar, bool(includeNA), pctMode,
            )
        except Exception as e:  # noqa: BLE001
            self.aviso(f"Falha ao calcular cruzamentos:\n{e}", "erro")
            return
        self.ultimas_tabelas = tabelas
        self.emitir("crosstabsReady", json.dumps({"tables": tabelas}, ensure_ascii=False))

    # ---------------- pilha
    def addToQueue(self, tableJson: str) -> None:
        try:
            t = json.loads(tableJson)
            if not isinstance(t, dict):
                return
        except Exception as e:  # noqa: BLE001
            self.aviso(f"Falha ao adicionar à pilha:\n{e}", "erro")
            return
        ok, pos = self._tentar_adicionar(t)
        if ok:
            self.notificar(f"Adicionado à pilha: {self.titulo(t)} — {len(self.queue)} tabela(s) na pilha.", "✅")
            self._push_queue_state()
        else:
            self.notificar(
                f"Esse cruzamento já está na pilha (item {pos + 1}): {self.titulo(self.queue[pos])}. "
                "Para trocar, remova o item na barra lateral e adicione de novo.", "⚠️")

    def addAllToQueue(self, tablesJson: str) -> None:
        try:
            tabelas = [t for t in json.loads(tablesJson) if isinstance(t, dict)]
        except Exception as e:  # noqa: BLE001
            self.aviso(f"Falha ao adicionar à pilha:\n{e}", "erro")
            return
        novas, repetidas = 0, 0
        for t in tabelas:
            ok, _ = self._tentar_adicionar(t)
            novas += ok
            repetidas += (not ok)
        if novas:
            msg = f"{novas} tabela(s) adicionada(s) à pilha"
            if repetidas:
                msg += f"; {repetidas} já estava(m) na pilha e não foi(ram) repetida(s)"
            self.notificar(msg + f". Total: {len(self.queue)}.", "✅")
            self._push_queue_state()
        elif repetidas:
            self.notificar("Todos esses cruzamentos já estão na pilha.", "⚠️")

    def removeQueueItem(self, index: int) -> None:
        try:
            index = int(index)
            if 0 <= index < len(self.queue):
                del self.queue[index]
            self._push_queue_state()
        except Exception as e:  # noqa: BLE001
            self.aviso(f"Falha ao remover item da pilha:\n{e}", "erro")

    def removerDaPilha(self, sig: str) -> None:
        """Remove pela assinatura (usado pela lixeira dos cards da barra lateral)."""
        for i, it in enumerate(self.queue):
            if self.assinatura(it) == sig:
                titulo = self.titulo(it)
                del self.queue[i]
                self.notificar(f"Removido da pilha: {titulo}", "🗑️")
                self._push_queue_state()
                return

    def clearQueue(self) -> None:
        if self.queue:
            self.notificar("Pilha esvaziada.", "🗑️")
        self.queue.clear()
        self._push_queue_state()

    def excel_da_pilha(self) -> bytes:
        """Excel da pilha inteira (guardado enquanto a pilha não muda)."""
        chave = hashlib.md5(json.dumps(self.queue, ensure_ascii=False, sort_keys=True, default=str)
                            .encode("utf-8")).hexdigest()
        cache = getattr(self, "_cache_excel", None)
        if not cache or cache[0] != chave:
            cache = (chave, exportar_excel(self.queue, nome_aba="Cruzamentos"))
            self._cache_excel = cache
        return cache[1]

    # ---------------- exportação
    def exportQueueExcel(self) -> None:
        if not self.queue:
            self.aviso("A pilha está vazia.", "info")
            return
        try:
            dados = exportar_excel(self.queue, nome_aba="Cruzamentos")
        except Exception as e:  # noqa: BLE001
            self.aviso(f"Falha ao exportar Excel:\n{e}", "erro")
            return
        self.download("cruzamentos.xlsx", dados, _MIME_XLSX)
        self.aviso(f"Arquivo gerado com {len(self.queue)} tabela(s).", "ok")

    def exportSingleExcel(self, tableJson: str) -> None:
        try:
            t = json.loads(tableJson)
            if not isinstance(t, dict):
                self.aviso("Tabela inválida para exportação.", "aviso")
                return
            dados = exportar_excel([t], nome_aba="Cruzamento")
        except Exception as e:  # noqa: BLE001
            self.aviso(f"Falha ao exportar:\n{e}", "erro")
            return
        self.download("cruzamento.xlsx", dados, _MIME_XLSX)
        self.aviso("Arquivo gerado.", "ok")


def processar_acoes_pilha(ponte: PonteCruzamentos, valor: Any) -> Dict[str, int]:
    """Lixeira dos cards da barra lateral (mantido para os testes)."""
    return processar_acoes(ponte, valor, ponte.removerDaPilha)


def _itens_pilha(ponte: PonteCruzamentos) -> List[Dict[str, Any]]:
    return [{
        "sig": ponte.assinatura(it),
        "title": ponte.titulo(it),
        "secoes": [
            {"rotulo": "Pergunta base", "texto": str(it.get("row_var_display") or it.get("row_var") or ""),
             "cor": "#60a5fa"},
            {"rotulo": "Cruzamento", "texto": str(it.get("col_var_display") or it.get("col_var") or ""),
             "cor": "#a78bfa"},
        ],
    } for it in ponte.queue]


def _pilha_na_barra_lateral(ponte: PonteCruzamentos) -> None:
    pilha_lateral(
        ponte, prefixo="cz", itens=lambda: _itens_pilha(ponte), remover=ponte.removerDaPilha,
        titulo="Pilha de exportação",
        vazio_msg="Vazia. Use “Adicionar à pilha” em cada tabela (ou “Adicionar todos à pilha”).",
        exportar=lambda: {"rotulo": f"Exportar Excel ({len(ponte.queue)})", "dados": ponte.excel_da_pilha,
                          "arquivo": "cruzamentos.xlsx", "mime": _MIME_XLSX},
        limpar=ponte.clearQueue, rotulo_limpar="Limpar pilha",
    )


def modulo_cruzamentos():
    st.header("Cruzamentos (Tabelas)")
    botao_tutorial("cruzamentos")

    ponte: PonteCruzamentos = obter_ponte("cz_ponte", PonteCruzamentos)
    processar_pedidos(ponte, "cz_componente")
    conferir_versao_tela(ponte, "relatoria_cruzamentos", VERSAO_TELA)
    _pilha_na_barra_lateral(ponte)
    mostrar_notificacoes(ponte)
    # altura "auto": a ferramenta cresce com o conteúdo e a PÁGINA rola
    # (sem uma "tela dentro da tela" com rolagem própria)
    desenhar_componente(ponte, "relatoria_cruzamentos", altura="auto", key="cz_componente")
