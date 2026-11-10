import io
import unittest

import numpy as np
import openpyxl
import pandas as pd

from core.total_completo_math import (
    achados_da_comparacao_tc,
    calcular_total_e_completo,
    comparar_com_descritivo,
    ler_descritivo,
    ler_tabelas,
)

SPS = """
* Encoding: UTF-8.
***********TOTAL***********
TEMPORARY.
  VARIABLE LABELS  Q1 '' .
TABLES
  /FORMAT BLANK Zero
  /FTOTAL= $tQ1 'Base'
  /TABLE= Q1 + $tQ1 BY total
  /STATISTICS
   cpct( Q1 ( F5.1 ) " ":total )
   UNW count( $tQ1 ( F5.0 ) '')
 /TITLE  'Titulo: Pergunta 1 (Estimulada e Única)'
/CAPTION 'Pergunta: Qual a sua opinião?'.

TEMPORARY.
  VARIABLE LABELS  $M '' .
CTABLES
  /VLABELS VARIABLES=$M TOTAL DISPLAY=NONE
  /TABLE $M [C][COLPCT.COUNT COMMA40.1, TOTALS[COUNT F40.0]] BY TOTAL [C] + SEXO [C]
  /SLABELS POSITION=LAYER VISIBLE=NO
  /CATEGORIES VARIABLES=$M  EMPTY=INCLUDE TOTAL=YES  LABEL ='Base' POSITION=AFTER
  /CATEGORIES VARIABLES=TOTAL ORDER=A KEY=VALUE EMPTY=INCLUDE
  /TITLES
       TITLE='Titulo: Meios (Estimulada e Múltipla)'
    CAPTION='Pergunta: Quais meios?'.

***********COMPLETO***********
TEMPORARY.
  VARIABLE LABELS  Q1 '' .
TABLES
  /FORMAT BLANK Zero
  /FTOTAL= $tQ1 'Base'
  /TABLE= Q1 + $tQ1 BY total+SEXO
  /STATISTICS
   cpct( Q1 ( F5.1 ) " ":total SEXO)
   UNW count( $tQ1 ( F5.0 ) '')
 /TITLE  'Titulo: Pergunta 1 (Estimulada e Única)'
/CAPTION 'Pergunta: Qual a sua opinião?'.
"""


def _df():
    return pd.DataFrame({
        "Q1": [1.0, 1.0, 2.0, 2.0, 3.0, 1.0, np.nan],
        "SEXO": [1.0, 2.0, 1.0, 2.0, 1.0, 2.0, 1.0],
        "M1": [1.0, 2.0, np.nan, 1.0, np.nan, 2.0, np.nan],
        "M2": [2.0, np.nan, np.nan, np.nan, np.nan, 3.0, np.nan],
        "peso": [1.0, 1.0, 2.0, 1.0, 1.0, 1.0, 1.0],
    })


VL = {"Q1": {1: "Bom", 2: "Ruim", 3: "NS"}, "SEXO": {1: "Masculino", 2: "Feminino"},
      "M1": {1: "TV", 2: "Rádio", 3: "Internet", 4: "Jornal"}}
VAR_LABELS = {"SEXO": "Sexo"}
MR = {"$M": ["M1", "M2"]}


def _calc(df=None, sps=SPS, peso="peso", vl=None):
    return calcular_total_e_completo(sps.encode(), df if df is not None else _df(), vl or VL, VAR_LABELS, MR, peso)


class LerTabelasTests(unittest.TestCase):
    def test_le_secoes_comandos_e_colunas(self):
        tabs, _ = ler_tabelas(SPS.encode())
        self.assertEqual([(t.secao, t.comando, t.var_linha, t.multipla) for t in tabs],
                         [("Total", "TABLES", "Q1", False), ("Total", "CTABLES", "M", True), ("Completo", "TABLES", "Q1", False)])
        self.assertEqual(tabs[1].colunas, ["TOTAL", "SEXO"])
        self.assertEqual(tabs[2].colunas, ["total", "SEXO"])
        self.assertEqual(tabs[0].titulo, "Pergunta 1 (Estimulada e Única)")
        self.assertEqual(tabs[0].pergunta, "Qual a sua opinião?")

    def test_estatistica_nao_suportada_vira_aviso(self):
        sps = SPS.replace("cpct( Q1 ( F5.1 ) \" \":total )", "mean( Q1 ( F5.1 ) )")
        tabs, achados = ler_tabelas(sps.encode())
        self.assertFalse(tabs[0].suportada)
        self.assertTrue(any(f.code == "TC_TABELA_NAO_SUPORTADA" for f in achados))


class CalculoTests(unittest.TestCase):
    def test_percentual_da_coluna_ponderado_e_base_sem_peso(self):
        r = _calc()
        c = r.calculadas[0]
        # casos válidos (Q1 não vazio): pesos 1,1,2,1,1,1 = 7 ; Q1=1 → 1+1+1=3 ; Q1=2 → 2+1=3 ; Q1=3 → 1
        self.assertEqual(c.rotulos, ["Bom", "Ruim", "NS"])
        np.testing.assert_allclose(c.valores[:, 0], [3 / 7 * 100, 3 / 7 * 100, 1 / 7 * 100])
        self.assertEqual(c.colunas[0].base, 6)            # Base = casos sem peso

    def test_segmento_e_soma_100(self):
        r = _calc()
        c = r.calculadas[2]
        self.assertEqual([(x.grupo, x.rotulo) for x in c.colunas],
                         [("Total", "Total"), ("Sexo", "Masculino"), ("Sexo", "Feminino")])
        for j in range(3):
            self.assertAlmostEqual(float(np.nansum(c.valores[:, j])), 100.0)
        # masculino: casos 1,3,5 → Q1=1(w1), 2(w2), 3(w1) → 25 / 50 / 25
        np.testing.assert_allclose(c.valores[:, 1], [25.0, 50.0, 25.0])
        self.assertEqual([x.base for x in c.colunas], [6, 3, 3])
        self.assertFalse(any(f.code == "TC_COLUNA_NAO_SOMA_100" for f in r.findings))

    def test_caso_sem_peso_sai_de_tudo(self):
        df = _df()
        df.loc[0, "peso"] = np.nan
        r = _calc(df)
        self.assertEqual(r.calculadas[0].colunas[0].base, 5)
        self.assertTrue(any(f.code == "TC_CASOS_SEM_PESO" for f in r.findings))

    def test_sem_peso(self):
        r = _calc(peso=None)
        # sem peso: Q1=1 → 3 casos, Q1=2 → 2, Q1=3 → 1 (de 6 válidos)
        np.testing.assert_allclose(r.calculadas[0].valores[:, 0], [50.0, 2 / 6 * 100, 1 / 6 * 100])

    def test_multipla_percentual_de_casos_e_categoria_vazia(self):
        r = _calc()
        c = r.calculadas[1]
        # respondentes (≥1 resposta): casos 1,2,4,6 → pesos 1,1,1,1 = 4
        # TV(1): casos 1,4 → 2 ; Rádio(2): casos 1,2,6 → 3 ; Internet(3): caso 6 → 1 ; Jornal(4): ninguém (EMPTY=INCLUDE)
        self.assertEqual(c.rotulos, ["TV", "Rádio", "Internet", "Jornal"])
        np.testing.assert_allclose(c.valores[:, 0], [50.0, 75.0, 25.0, 0.0])
        self.assertEqual(c.colunas[0].base, 4)

    def test_variavel_inexistente_e_critico(self):
        sps = SPS.replace("/TABLE= Q1 + $tQ1 BY total\n", "/TABLE= QX + $tQX BY total\n", 1).replace("cpct( Q1 ( F5.1 ) \" \":total )", "cpct( QX ( F5.1 ) \" \":total )", 1)
        r = _calc(sps=sps)
        self.assertIsNone(r.calculadas[0])
        self.assertTrue(any(f.code == "TC_TABELA_COM_ERRO" and f.severity == "critico" for f in r.findings))

    def test_codigo_sem_rotulo(self):
        vl = dict(VL)
        vl["Q1"] = {1: "Bom", 2: "Ruim"}
        r = _calc(vl=vl)
        self.assertTrue(any(f.code == "TC_CODIGO_SEM_ROTULO" for f in r.findings))

    def test_titulo_pergunta_e_repeticao(self):
        sps = SPS + SPS.split("***********COMPLETO***********")[1]
        r = _calc(sps=sps)
        self.assertTrue(any(f.code == "TC_TABELA_REPETIDA" for f in r.findings))


def _xlsx(linhas):
    wb = openpyxl.Workbook()
    ws = wb.active
    for l in linhas:
        ws.append(l)
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


class ComparacaoTests(unittest.TestCase):
    def _descritivo(self, valor_bom_masc=25.0):
        return _xlsx([
            ["Titulo: Pergunta 1 (Estimulada e Única)", None, None, None],
            [None, "Total", None, None],
            ["Bom", 3 / 7 * 100, None, None],
            ["Ruim", 3 / 7 * 100, None, None],
            ["NS", 1 / 7 * 100, None, None],
            ["Base", 6, None, None],
            ["Pergunta: Qual a sua opinião?", None, None, None],
            [None, None, None, None],
            ["Titulo: Pergunta 1 (Estimulada e Única)", None, None, None],
            [None, "Total", "Sexo", None],
            [None, None, "Masculino", "Feminino"],
            ["Bom", 3 / 7 * 100, valor_bom_masc, 200 / 3],
            ["Ruim", 3 / 7 * 100, 50.0, 100 / 3],
            ["NS", 1 / 7 * 100, 25.0, 0.0],
            ["Base", 6, 3, 3],
            ["Pergunta: Qual a sua opinião?", None, None, None],
        ])

    def test_le_o_descritivo(self):
        d = ler_descritivo(self._descritivo())
        self.assertEqual(len(d), 2)
        self.assertEqual(d[0].colunas, ["Total"])
        self.assertEqual(d[1].grupos, ["Total", "Sexo", "Sexo"])
        self.assertEqual(d[1].colunas, ["Total", "Masculino", "Feminino"])
        self.assertEqual([x[0] for x in d[1].linhas], ["Bom", "Ruim", "NS"])
        self.assertEqual(d[1].base, [6.0, 3.0, 3.0])

    def test_tabelas_iguais(self):
        r = _calc()
        comp = comparar_com_descritivo(r, ler_descritivo(self._descritivo()))
        por_ordem = {c.ordem: c.status for c in comp if c.ordem}
        self.assertEqual(por_ordem[1], "igual")
        self.assertEqual(por_ordem[3], "igual")
        self.assertEqual(por_ordem[2], "so_no_app")        # a múltipla não está neste Descritivo

    def test_acha_numero_diferente(self):
        r = _calc()
        comp = comparar_com_descritivo(r, ler_descritivo(self._descritivo(valor_bom_masc=30.0)))
        dif = [c for c in comp if c.status == "diferente"]
        self.assertEqual(len(dif), 1)
        self.assertEqual(dif[0].celulas_diferentes, 1)
        self.assertAlmostEqual(dif[0].maior_diferenca, 5.0)
        self.assertTrue(any(f.severity == "critico" for f in achados_da_comparacao_tc(comp)))

    def test_linha_base_reduzida_do_descritivo_e_ignorada(self):
        linhas = [["Titulo: Pergunta 1 (Estimulada e Única)", None],
                  [None, "Total"], ["Bom", 3 / 7 * 100], ["Ruim", 3 / 7 * 100], ["NS", 1 / 7 * 100],
                  ["Base reduzida", 1.0], ["Base", 6], ["Pergunta: x", None]]
        comp = comparar_com_descritivo(_calc(), ler_descritivo(_xlsx(linhas)))
        self.assertEqual({c.ordem: c.status for c in comp if c.ordem}[1], "igual")


if __name__ == "__main__":
    unittest.main()
