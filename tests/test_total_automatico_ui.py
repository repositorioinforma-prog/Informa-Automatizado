import sys
import types
import unittest

try:
    import streamlit  # noqa: F401
except Exception:                      # ambiente de teste sem streamlit: o módulo só precisa importar
    sys.modules.setdefault("streamlit", types.ModuleType("streamlit"))

from analises import _tema_ferramenta as tf
from analises import total_automatico as ta


class StepperTests(unittest.TestCase):
    def test_marca_feitas_atual_e_pendentes(self):
        h = ta._html_stepper("layout")
        self.assertEqual(h.count("ta-feito"), 4)           # limpeza, base, origem, termos
        self.assertEqual(h.count("ta-atual"), 1)
        self.assertEqual(h.count("ta-passo"), len(ta._ETAPAS_COM_PROGRESSO))
        self.assertIn("ETAPA 5 DE 9", h)
        self.assertIn('width:50%', h)

    def test_primeira_e_ultima_etapa(self):
        self.assertIn("ETAPA 1 DE 9", ta._html_stepper("limpeza_multiplas"))
        self.assertIn('width:0%', ta._html_stepper("limpeza_multiplas"))
        self.assertIn('width:100%', ta._html_stepper("final"))

    def test_etapa_fora_do_fluxo_nao_desenha(self):
        self.assertEqual(ta._html_stepper("aviso"), "")
        self.assertEqual(ta._html_stepper("upload"), "")

    def test_toda_etapa_tem_nome_curto_e_descricao(self):
        for chave, _nome in ta._ETAPAS_COM_PROGRESSO:
            self.assertIn(chave, ta._NOME_CURTO)
            self.assertIn(chave, ta._DESCRICAO_ETAPA)


class BoasVindasTests(unittest.TestCase):
    def test_lista_as_etapas_numeradas_sem_o_final(self):
        h = ta._html_boas_vindas()
        self.assertEqual(h.count('class="ta-item"'), len(ta._ETAPAS_COM_PROGRESSO) - 1)
        self.assertIn("Fora deste fluxo", h)
        self.assertIn("Legendas", h)

    def test_escapa_html(self):
        h = ta._html_boas_vindas()
        self.assertNotIn("<script", h)


class TemaTests(unittest.TestCase):
    def test_total_automatico_esta_no_tema_escuro(self):
        sub, icone, cor = tf.TEMA_ESCURO["Total Automático"]
        self.assertTrue(sub and icone.startswith("<") and cor.startswith("#"))
        self.assertIn("Total Automático", tf.html_cabecalho("Total Automático", sub, icone, cor))


if __name__ == "__main__":
    unittest.main()
