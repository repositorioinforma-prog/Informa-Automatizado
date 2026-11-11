import sys
import types
import unittest

try:
    import streamlit  # noqa: F401
except Exception:                      # ambiente de teste sem streamlit: o módulo só precisa importar
    sys.modules.setdefault("streamlit", types.ModuleType("streamlit"))

from analises import _assistente_ui as ui
from analises import _tema_ferramenta as tf
from analises import relatorio_automatizado as ra


class StepperTests(unittest.TestCase):
    def test_marca_feitas_atual_e_pendentes(self):
        h = ra._html_stepper("base_reduzida")               # 7ª etapa
        self.assertEqual(h.count("ta-feito"), 6)
        self.assertEqual(h.count("ta-atual"), 1)
        self.assertEqual(h.count("ta-passo"), len(ra._ETAPAS_COM_PROGRESSO))
        self.assertIn("ETAPA 7 DE 14", h)

    def test_primeira_e_ultima_etapa(self):
        self.assertIn("ETAPA 1 DE 14", ra._html_stepper("limpeza_multiplas"))
        self.assertIn("width:0%", ra._html_stepper("limpeza_multiplas"))
        self.assertIn("width:100%", ra._html_stepper("final"))

    def test_etapa_fora_do_fluxo_nao_desenha(self):
        self.assertEqual(ra._html_stepper("aviso"), "")
        self.assertEqual(ra._html_stepper("upload"), "")

    def test_toda_etapa_tem_nome_curto_e_descricao(self):
        for chave, _nome in ra._ETAPAS_COM_PROGRESSO:
            self.assertIn(chave, ra._NOME_CURTO)
            self.assertIn(chave, ra._DESCRICAO_ETAPA)


class BoasVindasTests(unittest.TestCase):
    def test_lista_as_etapas_numeradas_sem_o_final(self):
        h = ra._html_boas_vindas()
        self.assertEqual(h.count('class="ta-item"'), len(ra._ETAPAS_COM_PROGRESSO) - 1)
        self.assertIn("Fora deste fluxo", h)


class ModuloCompartilhadoTests(unittest.TestCase):
    def test_escapa_html_no_nome_do_arquivo_e_nas_etapas(self):
        self.assertNotIn("<script", ui.html_arquivo("<script>x</script>.xlsx"))
        h = ui.html_boas_vindas([("a", "<b>x</b>")], {"a": "<i>y</i>"}, ["<u>z</u>"], ignorar=())
        for tag in ("<b>x", "<i>y", "<u>z"):
            self.assertNotIn(tag, h)

    def test_arquivo_vazio(self):
        self.assertEqual(ui.html_arquivo(None), "")


class TemaTests(unittest.TestCase):
    def test_esta_no_tema_escuro(self):
        sub, icone, cor = tf.TEMA_ESCURO["Relatório Automatizado"]
        self.assertTrue(sub and icone.startswith("<") and cor.startswith("#"))
        self.assertIn("Relatório Automatizado", tf.html_cabecalho("Relatório Automatizado", sub, icone, cor))


if __name__ == "__main__":
    unittest.main()
