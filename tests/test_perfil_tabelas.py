"""Perfil (tabelas geradas do banco): conta do TABLES com peso, Base sem ponderar, Média só no presencial."""
import io
import os
import sys
import unittest

import numpy as np
import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import perfil_tabelas_math as pt  # noqa: E402
from core.perfil_math import ler_perfil  # noqa: E402

VL = {
    "P1": {1: "Masculino", 2: "Feminino"},
    "P2": {1: "18 a 29", 2: "30 a 39", 3: "40 ou mais"},
    "P3": {1: "Até 2", 2: "2 a 5"},
    "P4": {1: "Fundamental", 2: "Médio", 3: "Superior"},
    "juncaoreligiao": {1: "Católica", 4: "Evangélica", 9: "Ateu"},
}


def _banco():
    return pd.DataFrame(
        {
            "P1": [1, 1, 2, 2, 2, 1, np.nan, 2],
            "P2": [1, 2, 3, 1, 2, 3, 1, 1],
            "P3": [1, 1, 2, 2, 1, 1, 2, 1],
            "P4": [1, 2, 3, 3, 2, 1, 2, 2],
            "juncaoreligiao": [1, 4, 4, 1, 9, 1, 4, 1],
            "MEP2": [25.0, 35.0, 55.0, 22.0, 31.0, 48.0, 27.0, 29.0],
            "peso": [1.0, 2.0, 0.5, 1.5, 1.0, 1.0, 1.0, 2.0],
        }
    )


class TestCalculo(unittest.TestCase):
    def setUp(self):
        self.df = _banco()

    def test_percentual_ponderado_e_base_sem_ponderar(self):
        r = pt.calcular_perfil(self.df, VL, "peso", presencial=False)
        sexo = next(b for b in r.blocos if b.id == "sexo")
        # válidos (P1 não ausente): pesos M = 1+2+1 = 4 ; F = .5+1.5+1+2 = 5 ; total 9
        self.assertEqual(sexo.categorias, ["Masculino", "Feminino"])
        self.assertEqual(sexo.pct, [44.4, 55.6])
        self.assertEqual(sexo.base, 7)          # casos válidos, sem ponderar
        self.assertIsNone(sexo.media)

    def test_ordem_dos_blocos_e_titulos(self):
        r = pt.calcular_perfil(self.df, VL, "peso", presencial=False)
        self.assertEqual([b.titulo for b in r.blocos], ["Sexo", "Idade", "Renda Familiar", "Escolaridade", "Religiões"])
        self.assertEqual([b.var for b in r.blocos], ["P1", "P2", "P3", "P4", "juncaoreligiao"])

    def test_media_so_no_presencial_e_ponderada(self):
        pres = pt.calcular_perfil(self.df, VL, "peso", presencial=True)
        idade = next(b for b in pres.blocos if b.id == "idade")
        esperado = np.average(self.df["MEP2"], weights=self.df["peso"])
        self.assertAlmostEqual(idade.media, round(esperado, 2), places=2)
        self.assertTrue(all(b.media is None for b in pres.blocos if b.id != "idade"))
        onl = pt.calcular_perfil(self.df, VL, "peso", presencial=False)
        self.assertTrue(all(b.media is None for b in onl.blocos))

    def test_presencial_sem_mep2_avisa_e_nao_inventa(self):
        r = pt.calcular_perfil(self.df.drop(columns="MEP2"), VL, "peso", presencial=True)
        self.assertIsNone(next(b for b in r.blocos if b.id == "idade").media)
        self.assertTrue(any("MEP2" in a for a in r.avisos))

    def test_variavel_ausente_fica_de_fora_com_aviso(self):
        r = pt.calcular_perfil(self.df.drop(columns="P4"), VL, "peso", presencial=False)
        self.assertNotIn("escolaridade", [b.id for b in r.blocos])
        self.assertEqual(r.ausentes, [("P4", "Escolaridade")])
        self.assertTrue(any("P4" in a for a in r.avisos))

    def test_coluna_em_minusculas_e_achada(self):
        df = self.df.rename(columns={"P1": "p1", "juncaoreligiao": "JuncaoReligiao"})
        r = pt.calcular_perfil(df, VL, "peso", presencial=False)
        self.assertEqual(r.ausentes, [])

    def test_peso_ausente_ou_zero_sai_da_conta(self):
        df = self.df.copy()
        df.loc[0, "peso"] = np.nan
        df.loc[1, "peso"] = 0
        r = pt.calcular_perfil(df, VL, "peso", presencial=False)
        self.assertEqual(r.casos_sem_peso, 2)
        sexo = next(b for b in r.blocos if b.id == "sexo")
        self.assertEqual(sexo.base, 5)          # 7 válidos - 2 sem peso
        self.assertTrue(any("peso" in a for a in r.avisos))

    def test_sem_peso_conta_cada_caso_uma_vez(self):
        r = pt.calcular_perfil(self.df, VL, None, presencial=False)
        sexo = next(b for b in r.blocos if b.id == "sexo")
        self.assertEqual(sexo.pct, [42.9, 57.1])    # 3/7 e 4/7

    def test_categoria_sem_caso_nao_aparece(self):
        df = self.df.copy()
        df["P3"] = 1                           # a categoria 2 (rótulo existe) não tem caso
        r = pt.calcular_perfil(df, VL, "peso", presencial=False)
        renda = next(b for b in r.blocos if b.id == "renda")
        self.assertEqual(renda.categorias, ["Até 2"])
        self.assertEqual(renda.pct, [100.0])

    def test_variavel_sem_caso_valido_vai_para_sem_dados(self):
        df = self.df.copy()
        df["P4"] = np.nan
        r = pt.calcular_perfil(df, VL, "peso", presencial=False)
        self.assertEqual(r.sem_dados, [("P4", "Escolaridade")])
        self.assertNotIn("escolaridade", [b.id for b in r.blocos])

    def test_peso_inexistente_e_erro_claro(self):
        with self.assertRaises(ValueError):
            pt.calcular_perfil(self.df, VL, "peso_x", presencial=False)

    def test_rotulo_sem_value_label_usa_o_codigo(self):
        r = pt.calcular_perfil(self.df, {}, "peso", presencial=False)
        self.assertEqual(next(b for b in r.blocos if b.id == "sexo").categorias, ["1", "2"])


class TestExcel(unittest.TestCase):
    def setUp(self):
        self.res = pt.calcular_perfil(_banco(), VL, "peso", presencial=True)
        self.xlsx = pt.montar_excel_perfil(self.res, "Banco_RJ_Set26.sav", {"Cidade": "X", "UF": "RJ", "Mês/Ano": "Set26"})

    def test_o_leitor_do_slide_le_de_volta(self):
        lido = ler_perfil(self.xlsx)
        self.assertEqual(lido["pendentes"], [])
        for b in self.res.blocos:
            l = lido["blocos"][b.id]
            self.assertEqual(l["categorias"], b.categorias)
            for a, e in zip(l["valores"], b.pct):
                self.assertAlmostEqual(a, e, places=6)
            self.assertAlmostEqual(l["base"], b.base)
        idade = lido["blocos"]["idade"]
        self.assertAlmostEqual(idade["media"], next(b for b in self.res.blocos if b.id == "idade").media, places=6)
        self.assertIsNone(lido["blocos"]["sexo"]["media"])

    def test_online_nao_tem_linha_media(self):
        res = pt.calcular_perfil(_banco(), VL, "peso", presencial=False)
        wb = openpyxl.load_workbook(io.BytesIO(pt.montar_excel_perfil(res)))
        textos = [c.value for linha in wb[pt.NOME_ABA].iter_rows() for c in linha]
        self.assertNotIn("Média", textos)
        self.assertIn("Base", textos)

    def test_abas_e_conferencia(self):
        wb = openpyxl.load_workbook(io.BytesIO(self.xlsx))
        self.assertEqual(wb.sheetnames, [pt.NOME_ABA, pt.NOME_ABA_CONFERENCIA])
        cf = [[c.value for c in l] for l in wb[pt.NOME_ABA_CONFERENCIA].iter_rows()]
        achatado = [str(v) for l in cf for v in l if v is not None]
        self.assertIn("Presencial (rua)", achatado)
        self.assertIn("peso", achatado)

    def test_pre_visualizacao_mostra_texto_formatado(self):
        ws = pt.planilha_perfil(self.res, como_texto=True)
        valores = [c.value for linha in ws.iter_rows() for c in linha if c.value is not None]
        self.assertTrue(all(isinstance(v, str) for v in valores))
        self.assertIn("Base", valores)


class TestNomes(unittest.TestCase):
    def test_nome_do_arquivo(self):
        self.assertEqual(pt.nome_arquivo_perfil("Paty do Alferes", "rj", "Set26"), "Perfil_Paty_do_Alferes_RJ_Set26.xlsx")
        self.assertEqual(pt.nome_arquivo_perfil("São Paulo", "SP", "Jul26"), "Perfil_São_Paulo_SP_Jul26.xlsx")
        self.assertEqual(pt.nome_arquivo_perfil("", "RJ", "Set26"), "Perfil_RJ_Set26.xlsx")
        self.assertEqual(pt.nome_arquivo_perfil("", "", ""), "Perfil.xlsx")
        self.assertNotIn("/", pt.nome_arquivo_perfil("a/b", "RJ", "Set26"))

    def test_sugestao_de_uf_e_mes(self):
        self.assertEqual(pt.sugerir_identificacao("Banco_com_peso_de_voto_Deputados_RJ_Set26.sav"), ("RJ", "Set26"))
        self.assertEqual(pt.sugerir_identificacao("Total e Completo_Rio de Janeiro Deputados_RJ_Set26.sps"), ("RJ", "Set26"))
        self.assertEqual(pt.sugerir_identificacao("banco.sav"), ("", ""))


if __name__ == "__main__":
    unittest.main()
