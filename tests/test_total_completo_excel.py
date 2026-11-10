import io
import unittest

import openpyxl

from core.planilha_utils import worksheet_para_html
from core.total_completo_excel_math import (
    escolher,
    montar_excel_total_completo,
    nao_geradas,
    nome_arquivo_total_completo,
    planilha_tabelas,
)
from core.total_completo_math import comparar_com_descritivo, ler_descritivo
from tests.test_total_completo import SPS, _calc


def _abre(res, secoes=None):
    return openpyxl.load_workbook(io.BytesIO(montar_excel_total_completo(res, secoes)))


class LayoutTests(unittest.TestCase):
    def setUp(self):
        self.res = _calc()
        self.wb = _abre(self.res)
        self.ws = self.wb["Total e Completo"]

    def test_bloco_simples_total(self):
        ws = self.ws
        self.assertEqual(ws["A1"].value, "Titulo: Pergunta 1 (Estimulada e Única)")
        self.assertTrue(ws["A1"].font.b)
        self.assertIn("A1:B1", {str(m) for m in ws.merged_cells.ranges})
        self.assertEqual(ws["B2"].value, "Total")
        self.assertEqual([ws.cell(r, 1).value for r in (3, 4, 5, 6)], ["Bom", "Ruim", "NS", "Base"])
        self.assertAlmostEqual(ws["B3"].value, 3 / 7 * 100)            # valor inteiro, sem arredondar
        self.assertEqual(ws["B3"].number_format, "###0.0")
        self.assertEqual(ws["B6"].value, 6)                              # base sem peso
        self.assertEqual(ws["B6"].number_format, "###0")
        self.assertEqual(ws["A7"].value, "Pergunta: Qual a sua opinião?")
        self.assertEqual(ws["B3"].border.top.style, "medium")
        self.assertEqual(ws["B6"].border.bottom.style, "medium")
        self.assertEqual(ws["A3"].font.name, "Arial")
        self.assertEqual(ws["B3"].font.name, "Calibri")

    def test_multipla_com_categoria_vazia(self):
        ws = self.ws
        # tem a coluna Sexo → cabeçalho de 2 linhas; os dados começam em r + 3
        achou = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Titulo: Meios (Estimulada e Múltipla)"]
        self.assertEqual(len(achou), 1)
        r = achou[0]
        self.assertEqual([ws.cell(r + 3 + i, 1).value for i in range(5)], ["TV", "Rádio", "Internet", "Jornal", "Base"])
        self.assertEqual(ws.cell(r + 6, 2).value, 0.0)                   # EMPTY=INCLUDE: aparece com 0,0

    def test_cabecalho_com_grupo_em_duas_linhas(self):
        ws = self.ws
        r = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Titulo: Pergunta 1 (Estimulada e Única)"][-1]
        merges = {str(m) for m in ws.merged_cells.ranges}
        self.assertIn(f"B{r + 1}:B{r + 2}", merges)                      # «Total» ocupa as 2 linhas
        self.assertIn(f"C{r + 1}:D{r + 1}", merges)                      # «Sexo» sobre Masculino/Feminino
        self.assertEqual(ws.cell(r + 1, 3).value, "Sexo")
        self.assertEqual([ws.cell(r + 2, c).value for c in (3, 4)], ["Masculino", "Feminino"])
        self.assertEqual(ws.cell(r + 3, 3).value, 25.0)

    def test_ida_e_volta_com_o_leitor_do_descritivo(self):
        lido = ler_descritivo(montar_excel_total_completo(self.res))
        comp = comparar_com_descritivo(self.res, lido)
        self.assertEqual([c.status for c in comp if c.ordem], ["igual", "igual", "igual"])


class SecoesETextoTests(unittest.TestCase):
    def test_filtra_secao(self):
        res = _calc()
        self.assertEqual(len(escolher(res)), 3)
        self.assertEqual([c.tabela.secao for c in escolher(res, ["Completo"])], ["Completo"])
        ws = _abre(res, ["Total"])["Total e Completo"]
        titulos = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value).startswith("Titulo:")]
        self.assertEqual(len(titulos), 2)

    def test_aba_nao_geradas(self):
        sps = SPS.replace("/TABLE= Q1 + $tQ1 BY total\n", "/TABLE= QX + $tQX BY total\n", 1).replace(
            'cpct( Q1 ( F5.1 ) " ":total )', 'cpct( QX ( F5.1 ) " ":total )', 1)
        res = _calc(sps=sps)
        self.assertEqual(len(nao_geradas(res)), 1)
        wb = _abre(res)
        self.assertEqual(wb.sheetnames, ["Total e Completo", "Não geradas"])
        self.assertEqual(wb["Não geradas"]["C2"].value, "QX")
        self.assertNotIn("Não geradas", _abre(_calc()).sheetnames)

    def test_previa_em_texto_formatado(self):
        ws = planilha_tabelas(escolher(_calc())[:1], como_texto=True)
        self.assertEqual(ws["B3"].value, "42,9")                         # vírgula decimal, 1 casa
        self.assertEqual(ws["B6"].value, "6")
        html, _l, _c = worksheet_para_html(ws, max_linhas=50)
        self.assertIn("Titulo: Pergunta 1", html)
        self.assertIn("42,9", html)

    def test_nome_do_arquivo(self):
        n = "Total_e_Completo_Angra_dos_Reis_RJ_Set26.sps"
        self.assertEqual(nome_arquivo_total_completo(n), "Total e Completo_Angra_dos_Reis_RJ_Set26.xlsx")
        self.assertEqual(nome_arquivo_total_completo(n, ["Total"]), "Total_Angra_dos_Reis_RJ_Set26.xlsx")
        self.assertEqual(nome_arquivo_total_completo("Completo X (1).sps", ["Completo"]), "Completo_X.xlsx")
        self.assertEqual(nome_arquivo_total_completo(""), "Total e Completo.xlsx")


if __name__ == "__main__":
    unittest.main()
