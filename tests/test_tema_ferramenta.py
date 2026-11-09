"""Tema escuro das ferramentas: só ferramentas que existem no MENU, cabeçalho com HTML escapado."""
import ast
import os
import sys
import types
import unittest
from unittest.mock import MagicMock

if "streamlit" not in sys.modules:
    _st = MagicMock()
    _comp = types.ModuleType("streamlit.components")
    _v1 = types.ModuleType("streamlit.components.v1")
    _v1.declare_component = lambda *a, **k: (lambda **kw: None)
    _comp.v1 = _v1
    _st.components = _comp
    sys.modules.update({"streamlit": _st, "streamlit.components": _comp, "streamlit.components.v1": _v1})

from analises import _tema_ferramenta as tf  # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(__file__))


class TestTemaFerramenta(unittest.TestCase):
    def test_ferramentas_tematizadas_existem_no_menu(self):
        arvore = ast.parse(open(os.path.join(RAIZ, "app.py"), encoding="utf-8").read())
        menu = []
        for no in arvore.body:
            if isinstance(no, ast.Assign) and getattr(no.targets[0], "id", "") == "MENU":
                menu = [e.value for e in no.value.elts if isinstance(e, ast.Constant)]
        for nome in tf.TEMA_ESCURO:
            self.assertIn(nome, menu)

    def test_cabecalho_escapa_html(self):
        h = tf.html_cabecalho("<b>x</b>", "a & b", "<path/>", "#fff")
        self.assertIn("&lt;b&gt;x&lt;/b&gt;", h)
        self.assertIn("a &amp; b", h)
        self.assertIn("tf-ancora", h)

    def test_css_cobre_os_componentes_usados(self):
        for sel in ("stFileUploaderDropzone", "stExpander", "stAlert", "stDataFrame", "tab-list", "stBaseButton-primary"):
            self.assertIn(sel, tf._CSS)


if __name__ == "__main__":
    unittest.main()
