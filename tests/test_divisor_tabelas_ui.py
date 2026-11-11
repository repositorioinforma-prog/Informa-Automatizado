import sys
import types
import unittest

try:
    import streamlit  # noqa: F401
except Exception:                      # ambiente de teste sem streamlit: o módulo só precisa importar
    sys.modules.setdefault("streamlit", types.ModuleType("streamlit"))

from analises import _assistente_ui as ui
from analises import _tema_ferramenta as tf
from analises import divisor_tabelas as dv


class StepperTests(unittest.TestCase):
    def test_tres_etapas_e_nomes_curtos(self):
        self.assertEqual([c for c, _ in dv._ETAPAS], ["enviar", "dividir", "indice"])
        for chave, _nome in dv._ETAPAS:
            self.assertIn(chave, dv._NOME_CURTO)

    def test_stepper_na_etapa_dividir(self):
        h = ui.html_stepper(dv._ETAPAS, "dividir", dv._NOME_CURTO)
        self.assertEqual(h.count("ta-feito"), 1)
        self.assertEqual(h.count("ta-atual"), 1)
        self.assertIn("ETAPA 2 DE 3", h)


class CartoesTests(unittest.TestCase):
    def test_tres_regras_com_os_limites_do_calculo(self):
        itens = dv._cartoes_regras()
        self.assertEqual(len(itens), 3)
        h = ui.html_cartoes(itens)
        self.assertEqual(h.count('class="ta-item"'), 3)
        self.assertIn(f"{dv.ALTURA_UTIL:g}", h)

    def test_escapa_html(self):
        h = ui.html_cartoes([("<", "<b>x</b>", "<i>y</i>")])
        self.assertNotIn("<b>x", h)
        self.assertNotIn("<i>y", h)


class TemaTests(unittest.TestCase):
    def test_esta_no_tema_escuro(self):
        sub, icone, cor = tf.TEMA_ESCURO["Divisor de Tabelas de Regioes"]
        self.assertTrue(sub and icone.startswith("<") and cor.startswith("#"))


if __name__ == "__main__":
    unittest.main()
