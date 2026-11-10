import io
import unittest

import numpy as np
import openpyxl
import pandas as pd

from core.missings_cruzamento_math import (
    aplicar_missings_do_banco,
    listar_categorias,
    marcados_do_painel,
    missings_para_calculo,
    sintaxe_missing,
    sugestao,
    tabela_para_painel,
    variaveis_de_segmento,
)
from core.total_completo_math import calcular_total_e_completo, comparar_com_descritivo, ler_descritivo, ler_tabelas

SPS = """
***********COMPLETO***********
TEMPORARY.
  VARIABLE LABELS  Q1 '' .
TABLES
  /FORMAT BLANK Zero
  /FTOTAL= $tQ1 'Base'
  /TABLE= Q1 + $tQ1 BY total+P1+SEG+$MS
  /STATISTICS
   cpct( Q1 ( F5.1 ) " ":total P1 SEG $MS)
   UNW count( $tQ1 ( F5.0 ) '')
 /TITLE  'Titulo: Pergunta 1 (Estimulada e Única)'
/CAPTION 'Pergunta: Qual a sua opinião?'.
"""
SPS_TOTAL = SPS.replace("COMPLETO", "TOTAL").replace("BY total+P1+SEG+$MS", "BY total+SEG").replace(" P1 SEG $MS)", " SEG)")


def _df():
    n = 100
    rng = np.random.default_rng(1)
    seg = np.array([1] * 60 + [2] * 30 + [3] * 7 + [4] * 3)          # 3 e 4 têm base < 25
    return pd.DataFrame({
        "Q1": rng.integers(1, 4, n).astype(float),
        "P1": ([1, 2] * 50),
        "SEG": seg.astype(float),
        "M1": [1.0] * 80 + [2.0] * 5 + [np.nan] * 15,
        "M2": [np.nan] * 100,
        "peso": [1.0] * n,
    })


VL = {"Q1": {1: "A", 2: "B", 3: "C"}, "P1": {1: "Masc", 2: "Fem"}, "SEG": {1: "S1", 2: "S2", 3: "S3", 4: "S4", 99: "NS"},
      "M1": {1: "m1", 2: "m2", 3: "m3"}}
LB = {"P1": "Sexo", "SEG": "Segmento", "$MS": "Meios"}
MR = {"$MS": ["M1", "M2"]}


class ListarTests(unittest.TestCase):
    def setUp(self):
        self.tabs, _ = ler_tabelas(SPS.encode())
        self.cats = listar_categorias(self.tabs, _df(), VL, LB, MR, "peso")

    def test_variaveis_de_segmento_sem_total_e_so_do_completo(self):
        self.assertEqual(variaveis_de_segmento(self.tabs), ["P1", "SEG", "$MS"])
        tabs_t, _ = ler_tabelas(SPS_TOTAL.encode())
        self.assertEqual(variaveis_de_segmento(tabs_t), [])
        self.assertEqual(variaveis_de_segmento(tabs_t, ("Total",)), ["SEG"])

    def test_categorias_com_codigo_rotulo_e_casos(self):
        seg = {c.codigo: c for c in self.cats if c.var == "SEG"}
        self.assertEqual({k: v.n for k, v in seg.items()}, {1.0: 60, 2.0: 30, 3.0: 7, 4.0: 3, 99.0: 0})
        self.assertEqual(seg[3.0].rotulo, "S3")
        self.assertEqual(seg[1.0].rotulo_var, "Segmento")

    def test_regra_base_menor_que_25_com_excecoes(self):
        sug = {(v, k) for v, k in sugestao(self.cats)}
        self.assertEqual({k for v, k in sug if v == "SEG"}, {3.0, 4.0})      # 99 tem 0 casos: não é sugerido
        # P1 (exceção) nunca entra, mesmo com base pequena
        df = _df(); df["P1"] = [1] * 97 + [2] * 3
        cats = listar_categorias(self.tabs, df, VL, LB, MR, "peso")
        self.assertFalse(any(v == "P1" for v, _ in sugestao(cats)))
        cats2 = listar_categorias(self.tabs, df, VL, LB, MR, "peso", excecoes=[])
        self.assertIn(("P1", 2.0), sugestao(cats2))

    def test_base_minima_configuravel(self):
        cats = listar_categorias(self.tabs, _df(), VL, LB, MR, "peso", base_minima=5)
        self.assertEqual({k for v, k in sugestao(cats) if v == "SEG"}, {4.0})

    def test_conjunto_multiplo(self):
        ms = {c.codigo: c for c in self.cats if c.var == "$MS"}
        self.assertEqual({k: v.n for k, v in ms.items()}, {1.0: 80, 2.0: 5, 3.0: 0})
        self.assertTrue(ms[2.0].sugerido)

    def test_caso_sem_peso_nao_conta(self):
        df = _df(); df.loc[:9, "peso"] = np.nan
        seg1 = next(c for c in listar_categorias(self.tabs, df, VL, LB, MR, "peso") if c.var == "SEG" and c.codigo == 1.0)
        self.assertEqual(seg1.n, 50)

    def test_ja_missing_do_sav(self):
        rng = {"SEG": [{"lo": 4.0, "hi": 4.0}]}
        cats = listar_categorias(self.tabs, _df(), VL, LB, MR, "peso", missing_ranges=rng)
        c4 = next(c for c in cats if c.var == "SEG" and c.codigo == 4.0)
        self.assertTrue(c4.ja_missing)
        self.assertFalse(c4.sugerido)
        self.assertEqual(c4.n, 3)                       # mostra quantos casos tinha antes de virar missing
        out = aplicar_missings_do_banco(_df(), rng)
        self.assertEqual(int(out["SEG"].isna().sum()), 3)
        # o que já é missing no .sav não vai de novo para o cálculo (já chega vazio); só o que foi marcado a mais
        self.assertEqual(missings_para_calculo(cats, sugestao(cats)).get("SEG"), {3.0})

    def test_missing_de_texto(self):
        df = pd.DataFrame({"T": ["a", "-", "b"]})
        out = aplicar_missings_do_banco(df, {"T": [{"lo": "-", "hi": "-"}]})
        self.assertEqual(out["T"].isna().tolist(), [False, True, False])


class PainelESintaxeTests(unittest.TestCase):
    def setUp(self):
        self.tabs, _ = ler_tabelas(SPS.encode())
        self.cats = listar_categorias(self.tabs, _df(), VL, LB, MR, "peso")

    def test_painel_e_leitura_de_volta(self):
        painel = tabela_para_painel(self.cats, sugestao(self.cats))
        self.assertEqual(list(painel.columns), ["Variável", "Rótulo da variável", "Código", "Rótulo", "Casos", "Situação", "Missing"])
        marc = marcados_do_painel(painel)
        self.assertEqual(marc, sugestao(self.cats))
        painel.loc[(painel["Variável"] == "SEG") & (painel["Código"] == 3.0), "Missing"] = False      # a pessoa desmarca
        self.assertNotIn(("SEG", 3.0), marcados_do_painel(painel))

    def test_sintaxe(self):
        txt = sintaxe_missing(self.cats, sugestao(self.cats), MR)
        self.assertIn("MISSING VALUES SEG (3, 4).", txt)
        self.assertIn("(2)", txt)                        # conjunto múltiplo: vai nas variáveis-membro
        self.assertIn("M1", txt)

    def test_sintaxe_mantem_o_que_ja_era_missing(self):
        rng = {"SEG": [{"lo": 99.0, "hi": 99.0}]}
        cats = listar_categorias(self.tabs, _df(), VL, LB, MR, "peso", missing_ranges=rng)
        txt = sintaxe_missing(cats, sugestao(cats), MR, rng)
        self.assertIn("MISSING VALUES SEG (99, 3, 4).", txt)

    def test_sintaxe_avisa_mais_de_3_discretos(self):
        df = _df(); df["SEG"] = np.array([1] * 40 + [2, 3, 4, 5] * 5 + [6] * 20 + [7] * 20).astype(float)
        vl = {**VL, "SEG": {k: str(k) for k in range(1, 8)}}
        cats = listar_categorias(self.tabs, df, vl, LB, MR, "peso")
        txt = sintaxe_missing(cats, sugestao(cats), MR)
        self.assertIn("ATENÇÃO", txt)


class CalculoComMissingsTests(unittest.TestCase):
    def _res(self, missings=None):
        return calcular_total_e_completo(SPS.encode(), _df(), VL, LB, MR, "peso", missings)

    def test_categoria_marcada_nao_vira_coluna(self):
        sem = self._res().calculadas[0]
        com = self._res({"SEG": {3.0, 4.0}}).calculadas[0]
        rot = lambda c: [x.rotulo for x in c.colunas]                    # noqa: E731
        self.assertIn("S3", rot(sem))
        self.assertNotIn("S3", rot(com))
        self.assertNotIn("S4", rot(com))
        self.assertEqual(rot(com).count("S1"), 1)
        self.assertEqual(com.colunas[0].base, 100)                      # os casos continuam no Total
        self.assertEqual([(v, r, n) for v, _g, r, n in com.fora], [("SEG", "S3", 7), ("SEG", "S4", 3)])

    def test_missing_de_conjunto_multiplo(self):
        com = self._res({"$MS": {2.0}}).calculadas[0]
        self.assertNotIn("m2", [x.rotulo for x in com.colunas])
        self.assertIn("m1", [x.rotulo for x in com.colunas])

    def test_achado_e_sem_falso_alarme_de_segmento_que_nao_fecha(self):
        r = self._res({"SEG": {3.0, 4.0}})
        self.assertTrue(any(f.code == "TC_MISSING_APLICADO" for f in r.findings))
        self.assertFalse(any(f.code == "TC_SEGMENTO_NAO_FECHA" for f in r.findings))
        self.assertEqual(r.missings, {"SEG": {3.0, 4.0}})
        self.assertFalse(any(f.code == "TC_MISSING_APLICADO" for f in self._res().findings))

    def test_comparacao_nao_acusa_coluna_que_saiu_por_missing(self):
        sem = self._res()
        wb = openpyxl.Workbook(); ws = wb.active
        c = sem.calculadas[0]
        ws.append(["Titulo: Pergunta 1 (Estimulada e Única)"])
        ws.append([None] + [x.grupo for x in c.colunas])
        ws.append([None] + [x.rotulo for x in c.colunas])
        for i, rot in enumerate(c.rotulos):
            ws.append([rot] + [float(v) for v in c.valores[i]])
        ws.append(["Base"] + [x.base for x in c.colunas])
        ws.append(["Pergunta: Qual a sua opinião?"])
        b = io.BytesIO(); wb.save(b)
        d = ler_descritivo(b.getvalue())
        com = self._res({"SEG": {3.0, 4.0}})
        comp = comparar_com_descritivo(com, d)
        self.assertEqual([x.status for x in comp], ["igual"])           # sem «colunas do Descritivo sem tabela»


if __name__ == "__main__":
    unittest.main()
