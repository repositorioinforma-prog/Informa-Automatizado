"""Base reduzida: quando a ÚLTIMA coluna da tabela é excluída, a nova última
coluna fica com a borda direita média ("Bordas Externas Espessas")."""
import io
import unittest

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Border, Side

from core.planilha_utils import excluir_colunas_base_reduzida

FINA, MEDIA = Side(style="thin"), Side(style="medium")


def _tabela(ws, bases):
    """Título (A..H mesclado), cabeçalho Total + segmento (2 linhas), 2 linhas de dados e Base reduzida."""
    ult = 2 + len(bases)
    ws.cell(row=1, column=1, value="Intenção de voto")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ult)
    ws.cell(row=2, column=2, value="Total")
    ws.cell(row=2, column=3, value="Região")
    for i in range(len(bases)):
        ws.cell(row=3, column=3 + i, value=f"R{i + 1}")
    for r, rot in ((4, "Candidato A"), (5, "Candidato B")):
        ws.cell(row=r, column=1, value=rot)
        for c in range(2, ult + 1):
            ws.cell(row=r, column=c, value=50.0)
    ws.cell(row=6, column=1, value="Base reduzida")
    ws.cell(row=6, column=2, value=sum(bases))
    for i, b in enumerate(bases):
        ws.cell(row=6, column=3 + i, value=b)
    for r in range(2, 7):
        for c in range(1, ult + 1):
            ws.cell(row=r, column=c).border = Border(left=MEDIA if c == 1 else FINA,
                                                     right=MEDIA if c == ult else FINA, top=FINA, bottom=FINA)
    # como no Excel: a borda externa de uma mesclagem fica guardada na âncora
    ws.cell(row=2, column=3).border = Border(left=FINA, right=MEDIA, top=FINA, bottom=FINA)
    ws.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
    ws.merge_cells(start_row=2, start_column=2, end_row=3, end_column=2)
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=ult)
    return ult


def _recarregar(ws):
    buf = io.BytesIO()
    ws.parent.save(buf)
    return openpyxl.load_workbook(io.BytesIO(buf.getvalue())).active


def _direita(ws, r, c):
    s = ws.cell(row=r, column=c).border.right
    return s.style if s is not None else None


class TestBordaDireitaBaseReduzida(unittest.TestCase):
    def test_ultima_coluna_excluida(self):
        ws = Workbook().active
        ult = _tabela(ws, [80, 60, 40, 10])   # R4 (última) sai
        self.assertEqual(excluir_colunas_base_reduzida(ws, 25), 1)
        ws = _recarregar(_recarregar(ws))
        nova = ult - 1
        for r in range(2, 7):
            self.assertEqual(_direita(ws, r, nova), "medium", f"linha {r}")
        self.assertEqual(_direita(ws, 4, nova - 1), "thin")       # divisória interna intocada
        self.assertIsNone(ws.cell(row=4, column=ult).value)

    def test_duas_ultimas_excluidas(self):
        ws = Workbook().active
        ult = _tabela(ws, [80, 60, 12, 10])
        self.assertEqual(excluir_colunas_base_reduzida(ws, 25), 2)
        ws = _recarregar(ws)
        for r in range(2, 7):
            self.assertEqual(_direita(ws, r, ult - 2), "medium", f"linha {r}")

    def test_coluna_do_meio_nao_muda_nada(self):
        ws = Workbook().active
        ult = _tabela(ws, [80, 10, 60, 40])   # R2 sai; a última continua sendo R4
        excluir_colunas_base_reduzida(ws, 25)
        ws = _recarregar(ws)
        for r in range(2, 7):
            self.assertEqual(_direita(ws, r, ult - 1), "medium")
        self.assertEqual(_direita(ws, 4, ult - 2), "thin")


if __name__ == "__main__":
    unittest.main()
