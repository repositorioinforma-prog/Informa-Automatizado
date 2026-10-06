"""Regra de altura das labels da coluna A (23 caracteres -> 15, 46 -> 30...)."""
import unittest

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from core.planilha_utils import altura_label_por_caracteres, linhas_por_largura
from core.relatorio_automatizado_math import aplicar_codigo_09
from core.total_automatico_math import aplicar_codigo_09_total


def _din10(ws):
    """Fonte do relatório real (layout do código 05 / Total): DIN 10."""
    for linha in ws.iter_rows():
        for cel in linha:
            if cel.value is not None and not type(cel).__name__ == "MergedCell":
                cel.font = Font(name="DIN", size=10)


def _planilha_exemplo():
    wb = Workbook()
    ws = wb.active
    ws.column_dimensions["A"].width = 21.78
    # linha 1: pergunta mesclada (não é label)
    ws["A1"] = "Pergunta: " + "x" * 120
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 45
    # linha 2: cabeçalho da tabela
    ws["C2"] = "Total"
    ws["D2"] = "Masculino"
    # labels
    ws["A3"] = "Sim"                                   # 3 chars   -> 15
    ws["A4"] = "Branco / Nulo / Nenhum."               # 23 chars, cabe -> 15
    ws["A5"] = "a" * 24                                # 24 chars  -> 30
    # 41 chars (regra -> 30), mas a quebra real por palavra dá 3 linhas:
    # "Não sabe / Não" | "respondeu ao" | "entrevistador" -> 45
    ws["A6"] = "Não sabe / Não respondeu ao entrevistador"
    ws["A7"] = "b" * 47                                # 47 chars  -> 45
    for r in range(3, 8):
        ws.cell(row=r, column=3, value=0.5)
        ws.cell(row=r, column=4, value=0.25)
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, horizontal="left", indent=1)
        ws.row_dimensions[r].height = 16.0  # altura "estranha" antiga
    # linha 8: título mesclado
    ws["A8"] = "Titulo: Avaliação"
    ws.merge_cells("A8:F8")
    # linha 9: legenda solta na coluna A, sem dados na tabela -> não mexe
    ws["A9"] = "* Legenda longa que não pertence a nenhuma linha de tabela"
    ws["A9"].alignment = Alignment(wrap_text=False)
    ws.row_dimensions[9].height = 13
    _din10(ws)
    return wb, ws


class TestRegraCaracteres(unittest.TestCase):
    def test_faixas(self):
        self.assertEqual(altura_label_por_caracteres("a"), 15)
        self.assertEqual(altura_label_por_caracteres("a" * 22), 15)
        self.assertEqual(altura_label_por_caracteres("a" * 23), 15)
        self.assertEqual(altura_label_por_caracteres("a" * 24), 30)
        self.assertEqual(altura_label_por_caracteres("a" * 46), 30)
        self.assertEqual(altura_label_por_caracteres("a" * 47), 45)
        self.assertEqual(altura_label_por_caracteres("a" * 69), 45)
        self.assertEqual(altura_label_por_caracteres("a" * 70), 60)

    def test_espacos_contam(self):
        self.assertEqual(altura_label_por_caracteres("abc " * 5 + "abc"), 15)   # 23
        self.assertEqual(altura_label_por_caracteres("abc " * 5 + "abcd"), 30)  # 24

    def test_quebra_manual_soma_trechos(self):
        self.assertEqual(altura_label_por_caracteres(" Aluno em domicílio \n(Educação Pública)"), 30)
        self.assertEqual(altura_label_por_caracteres("a" * 30 + "\nb"), 45)  # 2 + 1 faixas

    def test_espacos_nas_pontas_ignorados(self):
        self.assertEqual(altura_label_por_caracteres("  " + "a" * 23 + "   "), 15)


class TestCodigo09(unittest.TestCase):
    def _verificar(self, fn):
        wb, ws = _planilha_exemplo()
        fn(ws)
        esperado = {3: 15, 4: 15, 5: 30, 6: 45, 7: 45}
        for r, h in esperado.items():
            self.assertEqual(ws.row_dimensions[r].height, h, f"linha {r}")
            a = ws.cell(row=r, column=1).alignment
            self.assertTrue(a.wrap_text)
            self.assertEqual(a.indent, 1)          # alinhamento preservado
            self.assertEqual(a.horizontal, "left")
        self.assertEqual(ws.row_dimensions[1].height, 45)  # pergunta intocada
        self.assertEqual(ws.row_dimensions[9].height, 13)  # legenda intocada
        self.assertEqual(ws["A8"].value, "Titulo: Avaliação")

    def test_relatorio_automatizado(self):
        self._verificar(aplicar_codigo_09)

    def test_total_automatico(self):
        self._verificar(aplicar_codigo_09_total)

    def test_texto_longo_em_outra_coluna_prevalece(self):
        wb, ws = _planilha_exemplo()
        ws["C3"] = "Texto\nem\ntrês"
        ws["C3"].alignment = Alignment(wrap_text=True)
        aplicar_codigo_09(ws)
        self.assertEqual(ws.row_dimensions[3].height, 45)



class TestLarguraReal(unittest.TestCase):
    """Caso real: 23 caracteres, mas letras largas -> o Excel quebra a linha."""

    def test_waguinho_republicanos_vai_para_30(self):
        wb, ws = _planilha_exemplo()
        ws["A3"] = "Waguinho - Republicanos"
        aplicar_codigo_09(ws)
        self.assertEqual(ws.row_dimensions[3].height, 30)
        wb, ws = _planilha_exemplo()
        ws["A3"] = "Waguinho - Republicanos"
        aplicar_codigo_09_total(ws)
        self.assertEqual(ws.row_dimensions[3].height, 30)

    def test_labels_comuns_de_22_23_continuam_15(self):
        for texto in ("Aprova muito o governo", "Branco / Nulo / Nenhum", "Eduardo Paes - PSD"):
            wb, ws = _planilha_exemplo()
            ws["A3"] = texto
            aplicar_codigo_09(ws)
            self.assertEqual(ws.row_dimensions[3].height, 15, texto)

    def test_regra_de_caracteres_continua_como_piso(self):
        # 24 caracteres estreitos: pela largura caberia, mas a regra manda 30
        wb, ws = _planilha_exemplo()
        ws["A3"] = "il " * 8
        ws["A3"] = ws["A3"].value.strip() + "lil"
        aplicar_codigo_09(ws)
        self.assertEqual(ws.row_dimensions[3].height, 30)

    def test_quebra_por_palavra(self):
        self.assertEqual(linhas_por_largura("Waguinho - Republicanos", 11000), 2)
        self.assertEqual(linhas_por_largura("Sim", 11000), 1)
        self.assertEqual(linhas_por_largura("W" * 30, 11000), 3)  # palavra enorme é partida

class TestDenominacoes(unittest.TestCase):
    def _planilha(self):
        wb = Workbook()
        ws = wb.active
        ws["C1"] = "Sexo"
        ws["E1"] = "Denominações Católicas"
        ws.merge_cells("E1:G1")
        ws["E1"].alignment = Alignment(wrap_text=True)
        ws.row_dimensions[1].height = 32
        ws["C2"] = "DENOMINACOES\nEVANGÉLICAS"
        ws["C2"].alignment = Alignment(wrap_text=True)
        ws.row_dimensions[2].height = 45
        # label de tabela na coluna A também é sobreposta pela regra fixa
        ws["A3"] = "Denominações Evangélicas " + "x" * 40
        ws["C3"] = 0.5
        return wb, ws

    def _verificar(self, fn):
        wb, ws = self._planilha()
        fn(ws)
        for r in (1, 2, 3):
            self.assertEqual(ws.row_dimensions[r].height, 15, f"linha {r}")

    def test_relatorio_automatizado(self):
        self._verificar(aplicar_codigo_09)

    def test_total_automatico(self):
        self._verificar(aplicar_codigo_09_total)



class TestFormatosTotalAutomatico(unittest.TestCase):
    """Tabelas estreitas (só Total): label mesclada em A:B e/ou dado na coluna B."""

    def _planilha(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Titulo: Avaliação do governo"   # título mesclado A:B
        ws.merge_cells("A1:B1")
        ws.row_dimensions[1].height = 30
        ws["A2"] = "Pergunta: " + "y" * 80           # pergunta mesclada larga
        ws.merge_cells("A2:C2")
        ws.row_dimensions[2].height = 45
        # labels mescladas em A:B, dado na coluna C
        ws["A3"] = "a" * 23
        ws.merge_cells("A3:B3")
        ws["C3"] = 0.4
        ws["A4"] = "a" * 30
        ws.merge_cells("A4:B4")
        ws["C4"] = 0.6
        # label sem mesclagem, dado na coluna B
        ws["A5"] = "b" * 50
        ws["B5"] = 0.1
        # label mesclada verticalmente (A6:A7) -> fora da regra
        ws["A6"] = "c" * 60
        ws.merge_cells("A6:A7")
        ws["C6"] = 0.2
        ws.row_dimensions[6].height = 20
        # nota mesclada A:B sem dado na tabela -> fora da regra
        ws["A8"] = "* nota " + "z" * 60
        ws.merge_cells("A8:B8")
        ws.row_dimensions[8].height = 13
        for r in range(3, 6):
            ws.row_dimensions[r].height = 16.0
        ws.column_dimensions["A"].width = 21.78
        ws.column_dimensions["B"].width = 8.78
        _din10(ws)
        return wb, ws

    def _verificar(self, fn):
        wb, ws = self._planilha()
        fn(ws)
        self.assertEqual(ws.row_dimensions[3].height, 15)
        self.assertEqual(ws.row_dimensions[4].height, 30)
        self.assertEqual(ws.row_dimensions[5].height, 45)
        self.assertTrue(ws["A3"].alignment.wrap_text)
        self.assertEqual(ws.row_dimensions[2].height, 45)   # pergunta intocada
        self.assertEqual(ws.row_dimensions[8].height, 13)   # nota intocada
        self.assertNotEqual(ws.row_dimensions[6].height, 60)  # vertical fora da regra

    def test_total_automatico(self):
        self._verificar(aplicar_codigo_09_total)

    def test_relatorio_automatizado(self):
        self._verificar(aplicar_codigo_09)


if __name__ == "__main__":
    unittest.main()
