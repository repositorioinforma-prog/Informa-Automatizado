"""Análise das Perguntas: leitura do Total, regra dos labels, cruzamento e Excel (núcleo, sem Streamlit)."""
import io
import os
import sys
import unittest

import numpy as np
import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import analise_perguntas_math as ap  # noqa: E402

PROGRAMA = """* Total.
TABLES
  /FORMAT BLANK MISSING('.') TVARS(NAMES)
  /TABLE=Q1 BY $T
  /STATISTICS COUNT ( F5.0 ) 'Total' .
* Titulo.
TABLES
  /TABLE=Q1 BY $T
  /TITLE 'Titulo: Pergunta Um (Dep.)'
  /CAPTION 'Pergunta: Qual o seu voto?'.
CTABLES
  /TABLE $Q4R [COUNT F40.0] BY $T
  /TITLE 'Titulo: Múltipla'.
TABLES
  /TABLE=Q2 BY $T
  /TITLE 'Titulo: Segunda'.
TABLES
  /TABLE=Q1 BY $T
  /TITLE 'Titulo: Outro texto de Q1'.

**** COMPLETO ****
TABLES
  /TABLE=Q9 BY P1
  /TITLE 'Titulo: So no completo'.
"""


class TestLerTotal(unittest.TestCase):
    def test_lista_variaveis_do_total_sem_repetir_e_ignora_completo(self):
        t = ap.ler_perguntas_total(PROGRAMA)
        self.assertEqual([x.var for x in t], ["Q1", "Q4R", "Q2"])
        q1 = t[0]
        self.assertEqual(q1.titulo, "Pergunta Um (Dep.)")   # ponto dentro das aspas não fecha o comando
        self.assertEqual(q1.pergunta, "Qual o seu voto?")
        self.assertTrue(t[1].multipla)
        self.assertFalse(t[2].multipla)

    def test_sem_cabecalho_completo_tudo_e_total(self):
        t = ap.ler_perguntas_total("TABLES\n  /TABLE=Q7 BY $T\n  /TITLE 'Titulo: X'.\n")
        self.assertEqual([x.var for x in t], ["Q7"])


class TestLabels(unittest.TestCase):
    def setUp(self):
        self.tabs = ap.ler_perguntas_total(PROGRAMA)

    def test_banco_tem_label_vale_o_do_banco(self):
        lv = ap.resolver_labels(self.tabs, {"q1": "Label do banco"})[0]
        self.assertEqual((lv.label_usada, lv.fonte), ("Label do banco", ap.FONTE_BANCO))
        self.assertTrue(lv.difere)

    def test_sem_label_no_banco_usa_titulo_do_total(self):
        lv = ap.resolver_labels(self.tabs, {"Q1": "   "})[0]
        self.assertEqual((lv.label_usada, lv.fonte), ("Pergunta Um (Dep.)", ap.FONTE_TITULO))

    def test_sem_nada_usa_o_nome_da_variavel(self):
        lv = ap.resolver_labels([ap.TabelaTotal("Z9")], {})[0]
        self.assertEqual((lv.label_usada, lv.fonte), ("Z9", ap.FONTE_SEM_LABEL))

    def test_preferir_titulo(self):
        lv = ap.resolver_labels(self.tabs, {"Q1": "Label do banco"}, preferir_titulo=True)[0]
        self.assertEqual(lv.fonte, ap.FONTE_TITULO)


class TestCruzamentos(unittest.TestCase):
    def test_padrao_so_o_que_existe_no_banco(self):
        pres, aus = ap.cruzamentos_padrao_presentes(["p1", "P2", "VN", "x"])
        self.assertEqual(pres, ["p1", "P2", "VN"])      # nome real da coluna
        self.assertIn("P4", aus)
        self.assertNotIn("VN", aus)


def _banco():
    return pd.DataFrame({
        "Q": [1, 1, 1, 2, 2, 2, 1, 2, 2, np.nan],
        "P1": [1, 1, 2, 2, 2, 1, 1, 2, 2, 1],
        "peso": [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 1.5, 0.4, 1.0],
    })


class TestCruzar(unittest.TestCase):
    def test_peso_arredonda_celulas_e_exclui_peso_zero(self):
        cz = ap.cruzar(_banco(), "Q", "P1", "peso", {"Q": {1.0: "Sim", 2.0: "Não"}, "P1": {1.0: "M", 2.0: "F"}})
        self.assertEqual(cz.rot_linhas, ["Sim", "Não"])
        self.assertEqual(cz.rot_colunas, ["M", "F"])
        # Q=2,P1=2: 1+1+1.5+0.4 = 3.9 -> 4 ; Q=2,P1=1: 1 ; Q=1,P1=1: 2 (o de peso 0 sai) ; Q=1,P1=2: 1
        np.testing.assert_array_equal(cz.contagem, [[2, 1], [1, 4]])
        self.assertEqual(cz.base_total, 8)
        self.assertEqual(cz.n_total, 8)                  # 10 linhas - Q ausente - peso 0
        self.assertEqual(cz.casos_sem_peso, 1)
        np.testing.assert_allclose(cz.pct_coluna[:, 0], [200 / 3, 100 / 3])

    def test_residuo_bate_com_formula(self):
        cz = ap.cruzar(_banco(), "Q", "P1", None)
        o = cz.contagem
        e = np.outer(o.sum(1), o.sum(0)) / o.sum()
        r, c = o.sum(1, keepdims=True) / o.sum(), o.sum(0, keepdims=True) / o.sum()
        np.testing.assert_allclose(cz.z, (o - e) / np.sqrt(e * (1 - r) * (1 - c)))

    def test_categoria_que_arredonda_para_zero_continua_na_tabela(self):
        # como no SPSS: a categoria tem caso válido, então fica, com contagem 0, % 0 e resíduo vazio
        df = pd.DataFrame({"Q": [1, 1, 1, 1, 2], "P1": [1, 1, 2, 2, 1], "peso": [1.0, 1.0, 1.0, 1.0, 0.3]})
        cz = ap.cruzar(df, "Q", "P1", "peso")
        self.assertEqual(cz.rot_linhas, ["1", "2"])
        np.testing.assert_array_equal(cz.contagem, [[2, 2], [0, 0]])
        np.testing.assert_array_equal(cz.pct_coluna[1], [0.0, 0.0])
        self.assertTrue(np.isnan(cz.z[1]).all())
        self.assertEqual(cz.n_total, 5)

    def test_sem_casos_devolve_none(self):
        df = pd.DataFrame({"Q": [np.nan], "P1": [1]})
        self.assertIsNone(ap.cruzar(df, "Q", "P1"))

    def test_pergunta_nao_cruza_com_ela_mesma(self):
        res = ap.calcular_analise(_banco(), ["Q"], ["Q", "P1"], "peso")
        self.assertEqual([c.coluna for c in res.por_pergunta["Q"]], ["P1"])


class TestExcel(unittest.TestCase):
    def _planilha(self, limiar=1.85):
        df = _banco()
        res = ap.calcular_analise(df, ["Q"], ["P1"], None, {"Q": {1.0: "Sim", 2.0: "Não"}, "P1": {1.0: "M", 2.0: "F"}})
        tabs = ap.ler_perguntas_total(PROGRAMA)
        labs = ap.resolver_labels([t for t in tabs if t.var == "Q1"], {})
        labs[0].var = "Q"
        data = ap.montar_excel(res, labs, {"P1": "Sexo"}, peso=None, limiar_z=limiar, perguntas_total=tabs,
                               nome_banco="b.sav", nome_syntax="t.sps")
        return openpyxl.load_workbook(io.BytesIO(data)), res

    def test_abas_e_layout_do_spss(self):
        wb, _ = self._planilha()
        self.assertEqual(wb.sheetnames, [ap.NOME_ABA, "Resumo"])
        ws = wb[ap.NOME_ABA]
        self.assertEqual(ws["A1"].value, "Pergunta Um (Dep.) * Sexo Crosstabulation")
        self.assertEqual(ws["D2"].value, "Sexo")
        self.assertEqual((ws["D3"].value, ws["E3"].value, ws["F2"].value), ("M", "F", "Total"))
        self.assertEqual((ws["B4"].value, ws["C4"].value, ws["C5"].value), ("Sim", "% within Sexo", "Adjusted Residual"))
        self.assertEqual(ws["A8"].value, "Total")
        self.assertEqual([ws.cell(8, k).value for k in (4, 5, 6)], [1.0, 1.0, 1.0])
        self.assertEqual(ws["D4"].number_format, "###0.0%")
        self.assertEqual(ws["D5"].number_format, "###0.0")
        textos = [c for linha in ws.iter_rows(values_only=True) for c in linha if isinstance(c, str)]
        self.assertNotIn("Contagem", textos)          # sem linhas de contagem, como na exportação do SPSS

    def test_pintura_como_o_vba(self):
        wb, res = self._planilha(limiar=0.5)
        ws = wb[ap.NOME_ABA]
        cz = res.por_pergunta["Q"][0]
        for i in range(cz.z.shape[0]):
            for j in range(cz.z.shape[1]):
                esperado = "FFC00000" if cz.z[i, j] <= -0.5 else "FF002060" if cz.z[i, j] >= 0.5 else None
                for r in (4 + 2 * i, 5 + 2 * i):      # % da alternativa e resíduo
                    c = ws.cell(r, 4 + j)
                    got = c.fill.fgColor.rgb if c.fill.fill_type else None
                    self.assertEqual(got, esperado, (r, j))
                    if esperado:
                        self.assertTrue(c.font.b)
                        self.assertEqual(c.font.color.rgb, "FFFFFFFF")
        # a coluna Total nunca é pintada e o resíduo dela fica vazio
        self.assertIsNone(ws["F5"].value)
        self.assertIsNone(ws["F4"].fill.fill_type)

    def test_limiar_alto_nao_pinta_nada(self):
        wb, _ = self._planilha(limiar=5.0)
        ws = wb[ap.NOME_ABA]
        self.assertFalse(any(c.fill.fill_type for linha in ws.iter_rows() for c in linha))

    def test_com_linha_count_como_no_spss(self):
        df = _banco()
        res = ap.calcular_analise(df, ["Q"], ["P1"], None, {"Q": {1.0: "Sim", 2.0: "Não"}, "P1": {1.0: "M", 2.0: "F"}})
        data = ap.montar_excel(res, [ap.LabelVar("Q", label_usada="Pergunta")], {"P1": "Sexo"}, peso=None,
                               limiar_z=0.5, incluir_contagem=True)
        ws = openpyxl.load_workbook(io.BytesIO(data))[ap.NOME_ABA]
        cz = res.por_pergunta["Q"][0]
        self.assertEqual([ws.cell(r, 3).value for r in (4, 5, 6)], ["Count", "% within Sexo", "Adjusted Residual"])
        self.assertEqual([ws.cell(4, c).value for c in (4, 5, 6)], [cz.contagem[0, 0], cz.contagem[0, 1], cz.total_linhas[0]])
        self.assertEqual(ws["D4"].number_format, "###0")
        # Total: linha Count (bases) + linha % within (100 %)
        self.assertEqual((ws["A10"].value, ws["C10"].value, ws["C11"].value), ("Total", "Count", "% within Sexo"))
        self.assertEqual([ws.cell(10, c).value for c in (4, 5, 6)], [cz.base_colunas[0], cz.base_colunas[1], cz.base_total])
        # a linha Count nunca é pintada; % e resíduo sim
        pintadas = [r for r in range(4, 10) if ws.cell(r, 4).fill.fill_type]
        self.assertTrue(pintadas)
        self.assertNotIn(4, pintadas)
        self.assertNotIn(7, pintadas)

    def test_previa_da_tela_e_o_mesmo_bloco_do_excel(self):
        res = ap.calcular_analise(_banco(), ["Q"], ["P1"], None, {"Q": {1.0: "Sim", 2.0: "Não"}, "P1": {1.0: "M", 2.0: "F"}})
        cz = res.por_pergunta["Q"][0]
        data = ap.montar_excel(res, [ap.LabelVar("Q", label_usada="Pergunta")], {"P1": "Sexo"}, peso=None, limiar_z=0.5,
                               incluir_contagem=True)
        no_excel = openpyxl.load_workbook(io.BytesIO(data))[ap.NOME_ABA]
        ws = ap.bloco_como_planilha(cz, "Pergunta", "Sexo", 0.5, True)
        self.assertEqual(ws.max_row, no_excel.max_row)
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                a, b = ws.cell(r, c), no_excel.cell(r, c)
                if isinstance(a.value, float):
                    self.assertAlmostEqual(a.value, b.value, places=9)
                else:
                    self.assertEqual(a.value, b.value)
                self.assertEqual((a.fill.fgColor.rgb if a.fill.fill_type else None), (b.fill.fgColor.rgb if b.fill.fill_type else None))

    def test_previa_formata_numeros_como_o_excel(self):
        cz = ap.cruzar(_banco(), "Q", "P1", None)
        ws = ap.bloco_como_planilha(cz, "P", "S", 1.85, True, como_texto=True)
        self.assertEqual(ws["D4"].value, str(int(cz.contagem[0, 0])))                       # Count: inteiro
        self.assertEqual(ws["D5"].value, f"{cz.pct_coluna[0, 0]:.1f}".replace(".", ",") + "%")  # % within
        self.assertTrue(ws["D6"].value is None or "," in ws["D6"].value or ws["D6"].value.lstrip("-").isdigit())
        self.assertEqual(ap._texto_formatado(0.0641, "###0.0%"), "6,4%")
        self.assertEqual(ap._texto_formatado(-1.926, "###0.0"), "-1,9")

    def test_valor_ausente_fica_vazio_e_nao_zero(self):
        wb, _ = self._planilha()
        ws = wb[ap.NOME_ABA]
        valores = [c for linha in ws.iter_rows(values_only=True) for c in linha if c is not None]
        self.assertFalse(any(isinstance(v, float) and np.isnan(v) for v in valores))


if __name__ == "__main__":
    unittest.main()
