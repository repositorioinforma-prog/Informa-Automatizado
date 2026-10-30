"""Tela 'Exclusão de Colunas por Base': percorre o fluxo inteiro (upload -> regras -> prévia ->
aplicar -> download) com um Streamlit de mentira, para pegar erros de código da tela sem navegador."""
import contextlib
import importlib
import io
import sys
import types
import unittest
from unittest import mock

import openpyxl
from openpyxl import Workbook

from tests.test_exclusao_colunas import tabela
from core.exclusao_colunas_math import TIPO_BASE_REDUZIDA


class _Rerun(Exception):
    pass


class _Arquivo:
    def __init__(self, nome, dados):
        self.name, self._dados = nome, dados

    def getvalue(self):
        return self._dados


class _FakeSt(types.ModuleType):
    """Registra o que a tela desenha; widgets devolvem o valor de `valores[chave]` ou o padrão."""

    def __init__(self):
        super().__init__("streamlit")
        self.session_state = {}
        self.valores, self.apertados, self.arquivo = {}, set(), None
        self.msgs, self.tabelas, self.downloads, self.metricas = [], [], [], {}
        self.components = types.SimpleNamespace(v1=types.SimpleNamespace(html=lambda *a, **k: None))

    # --- texto
    def _msg(tipo):
        def f(self, texto, *a, **k):
            self.msgs.append((tipo, str(texto)))
        return f
    header = subheader = markdown = caption = _msg("texto")
    info, warning, success, error = _msg("info"), _msg("warning"), _msg("success"), _msg("error")

    # --- contêineres
    @contextlib.contextmanager
    def _ctx(self, *a, **k):
        yield self
    spinner = expander = container = _ctx

    def columns(self, spec, *a, **k):
        return [self for _ in range(spec if isinstance(spec, int) else len(spec))]

    # --- widgets
    def file_uploader(self, *a, **k):
        return self.arquivo

    def selectbox(self, rotulo, opcoes, key=None, **k):
        return self.valores.get(key, list(opcoes)[0])

    def number_input(self, rotulo, key=None, value=0, **k):
        return self.valores.get(key, value)

    def multiselect(self, rotulo, opcoes, key=None, default=None, **k):
        v = self.valores.get(key, self.session_state.get(key, default or []))
        assert all(x in opcoes for x in v), f"{key}: {v} fora das opções {opcoes}"
        return v

    def checkbox(self, rotulo, value=False, key=None, **k):
        return self.valores.get(key, value)

    def button(self, rotulo, key=None, **k):
        return key in self.apertados

    def metric(self, rotulo, valor, **k):
        self.metricas[rotulo] = valor

    def dataframe(self, df, **k):
        self.tabelas.append(df)

    def download_button(self, rotulo, data=None, file_name=None, **k):
        self.downloads.append((file_name, data))

    def rerun(self):
        raise _Rerun()


def _xlsx(*montadores):
    wb = Workbook()
    ws = wb.active
    for m in montadores:
        m(ws)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestTelaExclusaoColunas(unittest.TestCase):
    def setUp(self):
        self.st = _FakeSt()
        patcher = mock.patch.dict(sys.modules, {"streamlit": self.st})
        patcher.start()
        self.addCleanup(patcher.stop)
        sys.modules.pop("analises.exclusao_colunas", None)
        self.tela = importlib.import_module("analises.exclusao_colunas")
        self.addCleanup(sys.modules.pop, "analises.exclusao_colunas", None)

    def _rodar(self):
        try:
            self.tela.modulo_exclusao_colunas()
        except _Rerun:
            self.st.apertados = set()             # botão só vale True na execução do clique
            self.tela.modulo_exclusao_colunas()   # como o Streamlit: reexecuta o script

    def _arquivo(self):
        return _Arquivo("relatorio.xlsx", _xlsx(
            lambda ws: tabela(ws, topo=1, bases=(80, 10, 60, 40)),
            lambda ws: tabela(ws, topo=9, tipo="Base", bases=(5, 60, 40, 50)),
        ))

    def test_sem_arquivo(self):
        self._rodar()
        self.assertTrue(any("Envie" in m for _, m in self.st.msgs))

    def test_previa_aplicar_e_baixar(self):
        self.st.arquivo = self._arquivo()
        self._rodar()
        self.assertEqual(self.st.metricas["Colunas a excluir"], 2)
        self.assertEqual(self.st.metricas["Tabelas afetadas"], 2)
        previa = self.st.tabelas[-1]
        self.assertEqual(sorted(previa["Segmento"]), ["R1", "R2"])
        self.assertEqual(self.st.downloads, [])           # ainda não aplicou

        self.st.apertados = {"ec_aplicar"}
        self._rodar()
        nome, dados = self.st.downloads[-1]
        self.assertEqual(nome, "relatorio - colunas excluidas.xlsx")
        ws = openpyxl.load_workbook(io.BytesIO(dados)).active
        self.assertEqual([ws.cell(row=3, column=c).value for c in range(3, 7)], ["R1", "R3", "R4", None])
        self.assertEqual([ws.cell(row=11, column=c).value for c in range(3, 7)], ["R2", "R3", "R4", None])
        self.assertTrue(any(t == "success" and "2 coluna(s)" in m for t, m in self.st.msgs))

    def test_somente_base_reduzida_e_valor_de_corte(self):
        self.st.arquivo = self._arquivo()
        self.st.valores = {"ec_linhas_0": "Somente Base reduzida", "ec_limite_0": 50}
        self._rodar()
        # BR da 1ª tabela: 10 (<50) sai; 40 também. A tabela com "Base" (5) fica de fora.
        self.assertEqual(self.st.metricas["Colunas a excluir"], 2)
        self.assertEqual(self.st.metricas["Tabelas afetadas"], 1)
        self.assertEqual(set(self.st.tabelas[-1]["Linha de base"]), {"Base reduzida"})

    def test_adicionar_e_remover_regra(self):
        self.st.arquivo = self._arquivo()
        self.st.apertados = {"ec_add_regra"}
        self._rodar()                                      # adiciona (rerun) e redesenha
        self.assertEqual(self.st.session_state["ec_ids"], [0, 1])
        self.st.apertados = {"ec_rm_0"}
        self._rodar()
        self.assertEqual(self.st.session_state["ec_ids"], [1])

    def test_filtro_por_grupo_e_nenhuma_coluna(self):
        self.st.arquivo = self._arquivo()
        self._rodar()
        self.st.valores = {"ec_grupos_0": ["Região"], "ec_limite_0": 1}
        self._rodar()
        self.assertEqual(self.st.metricas["Colunas a excluir"], 0)
        self.assertTrue(any("Nenhuma coluna atende" in m for _, m in self.st.msgs))

    def test_resultado_desatualizado_avisa(self):
        self.st.arquivo = self._arquivo()
        self.st.apertados = {"ec_aplicar"}
        self._rodar()
        self.st.apertados = set()
        self.st.valores = {"ec_limite_0": 15}
        self._rodar()
        self.assertTrue(any(t == "warning" and "mudaram" in m for t, m in self.st.msgs))

    def test_arquivo_sem_base(self):
        self.st.arquivo = _Arquivo("x.xlsx", _xlsx(lambda ws: ws.append(["a", 1, 2])))
        self._rodar()
        self.assertTrue(any(t == "warning" and "Não encontrei" in m for t, m in self.st.msgs))

    def test_arquivo_invalido(self):
        self.st.arquivo = _Arquivo("x.xlsx", b"isto nao e um xlsx")
        self._rodar()
        self.assertTrue(any(t == "error" for t, _ in self.st.msgs))

    def test_novo_arquivo_zera_as_regras(self):
        self.st.arquivo = self._arquivo()
        self.st.apertados = {"ec_add_regra"}
        self._rodar()
        self.assertEqual(self.st.session_state["ec_ids"], [0, 1])
        self.st.apertados = set()
        self.st.arquivo = _Arquivo("outro.xlsx", _xlsx(lambda ws: tabela(ws, bases=(80, 10, 60, 40))))
        self._rodar()
        self.assertEqual(self.st.session_state["ec_ids"], [0])


if __name__ == "__main__":
    unittest.main()
