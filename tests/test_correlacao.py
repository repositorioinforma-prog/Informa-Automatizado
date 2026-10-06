"""Matriz de Correlação (portada do "Relatória Analítica").

A paridade com a CorrelationBridge original foi conferida na migração (CSV,
XLSX e SAV: mesmas colunas, valores únicos, matrizes nos 3 métodos e com
missing, CSV byte a byte e XLSX célula a célula); estes testes fixam as
regras principais para não regredirem.
"""
import base64
import io
import json
import os
import sys
import types
import unittest

import numpy as np
import openpyxl
import pandas as pd

if "streamlit" not in sys.modules:
    _st = types.ModuleType("streamlit")
    _st.session_state = {}
    _comp = types.ModuleType("streamlit.components")
    _v1 = types.ModuleType("streamlit.components.v1")
    _v1.declare_component = lambda *a, **k: (lambda **kw: None)
    _comp.v1 = _v1
    _st.components = _comp
    sys.modules.update({"streamlit": _st, "streamlit.components": _comp, "streamlit.components.v1": _v1})

from analises.correlacao import PonteCorrelacao, VERSAO_TELA  # noqa: E402
from core import correlacao_math as cm  # noqa: E402


def _csv():
    rng = np.random.default_rng(3)
    n = 300
    x = rng.normal(0, 1, n)
    df = pd.DataFrame({
        "a": x, "b": 2 * x + rng.normal(0, 0.3, n), "c": rng.normal(0, 1, n),
        "d": np.where(np.arange(n) % 10 == 0, 99, x), "Peso": rng.uniform(0.5, 1.5, n),
    })
    return df.to_csv(index=False).encode("utf-8")


def _args(js):
    return json.loads(json.loads(js[js.index("(") + 1:js.rindex(")")]))


class TestCorrelacao(unittest.TestCase):
    def setUp(self):
        self.p = PonteCorrelacao()
        self.p.pickData(arquivo={"name": "base.csv", "dados": _csv()})

    def _ultimo(self, nome):
        return _args([e["codigo"] for e in self.p.eventos if e["tipo"] == "js" and e["codigo"].startswith(nome)][-1])

    def test_colunas_e_peso(self):
        cols = self._ultimo("window.__setColumns")
        self.assertEqual(cols["weight_candidates"], ["Peso"])
        self.assertIn("a", cols["columns"])

    def test_pearson_ponderado_e_filtro_sig(self):
        self.p.computeCorr(json.dumps({"cols": ["a", "b", "c"], "weightVar": "Peso", "decimals": 3}))
        r = self._ultimo("window.__setCorrResult")
        self.assertTrue(r["ok"])
        m = np.array([[np.nan if v is None else v for v in row] for row in r["matrix"]])
        self.assertGreater(m[0, 1], 0.9)        # a x b: forte e significativa
        self.assertTrue(np.isnan(m[0, 2]))      # a x c: ruído -> filtrada pelo Sig.
        w = self.p.df["Peso"].to_numpy()
        self.assertAlmostEqual(m[0, 1], cm._weighted_pearson(self.p.df["a"].to_numpy(), self.p.df["b"].to_numpy(), w), 10)

    def test_missing_por_variavel_e_global(self):
        base = json.dumps({"cols": ["a", "d"], "weightVar": "Peso"})
        self.p.computeCorr(base)
        sem = self._ultimo("window.__setCorrResult")["matrix"][0][1]
        self.p.computeCorr(json.dumps({"cols": ["a", "d"], "weightVar": "Peso", "missingRules": {"d": ["99"]}}))
        regra = self._ultimo("window.__setCorrResult")["matrix"][0][1]
        self.p.computeCorr(json.dumps({"cols": ["a", "d"], "weightVar": "Peso", "missingVals": "99"}))
        glob = self._ultimo("window.__setCorrResult")["matrix"][0][1]
        self.assertAlmostEqual(regra, 1.0, 6)
        self.assertAlmostEqual(regra, glob, 10)
        self.assertNotAlmostEqual(sem or 0.0, regra, 3)

    def test_validacoes(self):
        self.p.computeCorr(json.dumps({"cols": ["a"], "weightVar": "Peso"}))
        self.assertEqual(self._ultimo("window.__setCorrResult")["error"], "Selecione pelo menos 2 colunas.")
        self.p.computeCorr(json.dumps({"cols": ["a", "b"], "weightVar": ""}))
        self.assertEqual(self._ultimo("window.__setCorrResult")["error"], "Selecione o peso (obrigatório).")

    def test_exportacoes(self):
        self.p.exportCorrXlsx("{}")
        self.assertEqual([e["tipo"] for e in self.p.eventos][-1], "aviso")  # sem calcular
        self.p.computeCorr(json.dumps({"cols": ["a", "b"], "weightVar": "Peso"}))
        self.p.exportCorrXlsx(json.dumps({"display": ["A", "B"]}))
        dl = [e for e in self.p.eventos if e["tipo"] == "download"][-1]
        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(dl["b64"])))
        self.assertEqual(wb.sheetnames, ["Correlações", "Sig. (2-tailed)"])
        self.assertEqual(wb["Correlações"]["B1"].value, "A")
        self.p.exportCorrCsv(json.dumps({"display": ["A", "B"]}))
        dl = [e for e in self.p.eventos if e["tipo"] == "download"][-1]
        self.assertTrue(base64.b64decode(dl["b64"]).startswith(b"\xef\xbb\xbf;A;B"))

    def test_ponte_de_versao_antiga_e_atualizada(self):
        """Código atualizado com o app aberto: a ponte antiga da sessão vira a
        nova (com os métodos novos) sem perder a base carregada."""
        import streamlit as st
        from analises._ponte_relatoria import PonteBase, obter_ponte

        class Antiga(PonteBase):
            pass
        velha = Antiga()
        velha.df, velha.data_label = self.p.df, "base.csv"
        st.session_state["corr_teste"] = velha
        nova = obter_ponte("corr_teste", PonteCorrelacao)
        self.assertIs(type(nova), PonteCorrelacao)
        self.assertEqual(nova.data_label, "base.csv")
        nova.despachar({"method": "saveUiState", "args": ['{"x": 1}']})
        self.assertEqual(nova.ui_state_json, '{"x": 1}')
        self.assertFalse([e for e in nova.eventos if e["tipo"] == "aviso"])
        del st.session_state["corr_teste"]

    def test_versao_da_tela(self):
        pasta = os.path.join(os.path.dirname(os.path.dirname(__file__)), "componentes", "relatoria_correlacao")
        html = open(os.path.join(pasta, "index.html"), encoding="utf-8").read()
        self.assertIn(f'<meta name="versao-tela" content="{VERSAO_TELA}" />', html)
        self.assertEqual(open(os.path.join(pasta, "ponte_qt.js"), "rb").read(),
                         open(os.path.join(os.path.dirname(pasta), "ponte_qt.js"), "rb").read())


if __name__ == "__main__":
    unittest.main()
