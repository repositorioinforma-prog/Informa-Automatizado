import unittest

import numpy as np
import pandas as pd

from core.analitica_dados import BancoCarregado, carregar_banco
from core.tratamento_banco_math import (
    banco_sem_geradas,
    comparar_com_spss,
    exportar_sav,
    nome_arquivo_tratado,
    tratar_banco,
)

SPS = """
recode SEXO (1=1) (2=2) into P1.
var label P1 'Sexo'.
value label /P1 1 'Masculino' 2 'Feminino'.
recode Q4 (1 thru 3=copy) (97=97) (98=98) (99=99) into Q4R.
var label Q4R 'Voto Q4'.
value label /Q4R 1 'A' 2 'B' 3 'C' 97 'Branco/Nulo' 98 'Não irá votar' 99 'Não sabe'.
recode Q5 (1 thru 3=copy) (97=97) (98=98) (99=99) into Q5R.
var label Q5R 'Voto Q5'.
value label /Q5R 1 'A' 2 'B' 3 'C' 97 'Branco/Nulo' 98 'Não irá votar' 99 'Não sabe'.
recode Q6 (1 thru 3=copy) (97=97) (98=98) (99=99) into Q6R.
var label Q6R 'Voto Q6'.
value label /Q6R 1 'A' 2 'B' 3 'C' 97 'Branco/Nulo' 98 'Não irá votar' 99 'Não sabe'.
recode Q7 (1 thru 3=copy) (97=97) (98=98) (99=99) into Q7R.
var label Q7R 'Voto Q7'.
value label /Q7R 1 'A' 2 'B' 3 'C' 98 'Branco/Nulo' 97 'Não irá votar' 99 'Não sabe'.
"""


def _banco(com_geradas=False):
    df = pd.DataFrame({
        "SEXO": [1.0, 2.0, 2.0, 1.0],
        "Q4": [1.0, 2.0, 97.0, 99.0],
        "Q5": [3.0, 2.0, 98.0, 1.0],
        "Q6": [1.0, 1.0, 97.0, 99.0],
        "Q7": [2.0, 3.0, 97.0, 98.0],
    })
    return BancoCarregado(df=df, nome_arquivo="Banco_Original_X.sav")


class TratarBancoTests(unittest.TestCase):
    def setUp(self):
        self.banco = _banco()
        self.res = tratar_banco(SPS.encode(), self.banco)

    def test_cria_as_variaveis(self):
        self.assertEqual(self.res.execucao.criadas, ["P1", "Q4R", "Q5R", "Q6R", "Q7R"])
        self.assertEqual(self.res.execucao.df["Q4R"].tolist(), [1.0, 2.0, 97.0, 99.0])

    def test_pega_codigo_97_98_trocado_entre_variaveis(self):
        f = [x for x in self.res.achados_do_tratamento if x.code == "TRAT_CODIGO_INCONSISTENTE"]
        self.assertEqual({x.variable for x in f}, {"Q7R"})
        self.assertEqual(len(f), 2)                   # 'Branco/Nulo' e 'Não irá votar'
        self.assertFalse(self.res.aprovado)

    def test_aprova_quando_esta_tudo_coerente(self):
        sps = SPS.replace("98 'Branco/Nulo' 97 'Não irá votar'", "97 'Branco/Nulo' 98 'Não irá votar'")
        r = tratar_banco(sps.encode(), self.banco)
        self.assertEqual([x.code for x in r.achados_do_tratamento if x.severity in ("critico", "possivel_erro")], [])
        self.assertTrue(r.aprovado)

    def test_codigo_sem_rotulo_e_rotulo_sem_caso(self):
        sps = "recode Q4 (1 thru 3=copy) (97=97) into Q4R.\nvalue label /Q4R 1 'A' 2 'B' 5 'E'."
        r = tratar_banco(sps.encode(), self.banco)
        codigos = {x.code for x in r.achados_do_tratamento}
        self.assertIn("TRAT_CODIGO_SEM_ROTULO", codigos)      # 3 e 97 sem rótulo
        self.assertIn("TRAT_ROTULO_SEM_CASO", codigos)        # 5 sem caso
        self.assertIn("TRAT_SEM_ROTULO_DE_VARIAVEL", codigos)

    def test_variavel_vazia(self):
        sps = "recode Q4 (50=1) (51=2) into Q4R.\nvar label Q4R 'x'.\nvalue label /Q4R 1 'a' 2 'b'."
        r = tratar_banco(sps.encode(), self.banco)
        f = [x for x in r.achados_do_tratamento if x.code == "TRAT_VARIAVEL_VAZIA"]
        self.assertEqual(f[0].severity, "possivel_erro")
        sps1 = "recode Q4 (50=1) into Q4R.\nvar label Q4R 'x'.\nvalue label /Q4R 1 'a'."
        f1 = [x for x in tratar_banco(sps1.encode(), self.banco).achados_do_tratamento if x.code == "TRAT_VARIAVEL_VAZIA"]
        self.assertEqual(f1[0].severity, "revisar")       # só separa uma categoria

    def test_mrset_com_membro_inexistente(self):
        sps = SPS + "\nMRSETS /MCGROUP NAME=$X VARIABLES= Q4R FANTASMA /DISPLAY NAME=[$X].\n"
        r = tratar_banco(sps.encode(), self.banco)
        self.assertTrue(any(x.code == "TRAT_MRSET_MEMBRO_AUSENTE" and x.severity == "critico"
                            for x in r.achados_do_tratamento))


class RotulosParecidosTests(unittest.TestCase):
    def test_pega_grafia_diferente_entre_variaveis(self):
        df = pd.DataFrame({"Q1": [1.0, 2.0], "Q2": [1.0, 2.0]})
        sps = """
recode Q1 (1=1) (2=2) into Q1R.
var label Q1R 'a'.
value label /Q1R 1 'Alex da Confecção (AVANTE)' 2 'Célia Jordão (PSD)'.
recode Q2 (1=1) (2=2) into Q2R.
var label Q2R 'b'.
value label /Q2R 1 'Alexda Confecção (AVANTE)' 2 'Célia Jordão (PSD)'.
"""
        r = tratar_banco(sps.encode(), BancoCarregado(df, "x.sav"))
        f = [x for x in r.achados_do_tratamento if x.code == "TRAT_ROTULO_PARECIDO"]
        self.assertEqual(len(f), 1)
        self.assertEqual(f[0].severity, "revisar")
        self.assertIn("Alexda", f[0].title)

    def test_opcoes_parecidas_na_mesma_variavel_nao_contam(self):
        df = pd.DataFrame({"Q1": [1.0, 2.0]})
        sps = """
recode Q1 (1=1) (2=2) into Q1R.
var label Q1R 'a'.
value label /Q1R 1 'Carlos Silva (PL)' 2 'Carlos Silva (PL) '.
"""
        r = tratar_banco(sps.encode(), BancoCarregado(df, "x.sav"))
        self.assertFalse(any(x.code == "TRAT_ROTULO_PARECIDO" for x in r.achados))

    def test_so_maiuscula_e_informativo(self):
        df = pd.DataFrame({"Q1": [1.0], "Q2": [1.0]})
        sps = """
recode Q1 (1=1) into A.
var label A 'a'.
value label /A 1 'Não sabe'.
recode Q2 (1=1) into B.
var label B 'b'.
value label /B 1 'Não Sabe'.
"""
        r = tratar_banco(sps.encode(), BancoCarregado(df, "x.sav"))
        f = [x for x in r.achados_do_tratamento if x.code == "TRAT_ROTULO_PARECIDO"]
        self.assertEqual([x.severity for x in f], ["informativo"])


class ComparacaoComSpssTests(unittest.TestCase):
    def _spss(self, res):
        e = res.execucao
        return BancoCarregado(
            df=e.df.copy(), nome_arquivo="spss.sav",
            var_labels={k: v for k, v in e.var_labels.items()},
            value_labels={k: dict(v) for k, v in e.value_labels.items()},
        )

    def setUp(self):
        self.banco = _banco()
        self.base = tratar_banco(SPS.encode(), self.banco)

    def test_igual(self):
        r = tratar_banco(SPS.encode(), self.banco, self._spss(self.base))
        self.assertEqual(len(r.comparacao), 5)
        self.assertEqual(r.divergentes, [])

    def test_valor_diferente(self):
        spss = self._spss(self.base)
        spss.df.loc[1, "P1"] = 1.0
        r = tratar_banco(SPS.encode(), self.banco, spss)
        d = {c.variavel: c for c in r.divergentes}
        self.assertEqual(set(d), {"P1"})
        self.assertEqual(d["P1"].casos_diferentes, 1)
        self.assertIn("caso 2", d["P1"].exemplo)
        self.assertTrue(any(x.code == "TRAT_VALORES_DIFERENTES" and x.severity == "critico" for x in r.achados))
        self.assertFalse(r.aprovado)

    def test_rotulo_diferente(self):
        spss = self._spss(self.base)
        spss.value_labels["P1"][2] = "Mulher"
        spss.var_labels["Q4R"] = "outro texto"
        r = tratar_banco(SPS.encode(), self.banco, spss)
        d = {c.variavel: c for c in r.divergentes}
        self.assertEqual(set(d), {"P1", "Q4R"})
        self.assertFalse(d["P1"].rotulos_de_valor_iguais)
        self.assertFalse(d["Q4R"].rotulo_de_variavel_igual)

    def test_variavel_ausente_no_spss(self):
        spss = self._spss(self.base)
        spss.df = spss.df.drop(columns=["Q6R"])
        r = tratar_banco(SPS.encode(), self.banco, spss)
        self.assertEqual([c.variavel for c in r.divergentes], ["Q6R"])
        self.assertFalse(r.divergentes[0].existe_no_spss)

    def test_numero_de_casos_diferente(self):
        spss = self._spss(self.base)
        spss.df = spss.df.iloc[:2]
        c = comparar_com_spss(self.base.execucao, spss)
        self.assertTrue(all(not x.valores_iguais for x in c))


class PartirDoZeroTests(unittest.TestCase):
    def test_remove_as_variaveis_que_a_syntax_cria(self):
        banco = _banco()
        banco.df["P1"] = 9.0
        banco.var_labels["P1"] = "velho"
        limpo, removidas = banco_sem_geradas(SPS.encode(), banco)
        self.assertEqual(removidas, ["P1"])
        self.assertNotIn("P1", limpo.df.columns)
        self.assertNotIn("P1", limpo.var_labels)
        self.assertIn("P1", banco.df.columns)         # o original não é mexido

    def test_rodando_sobre_banco_ja_tratado_da_o_mesmo_resultado(self):
        banco = _banco()
        primeiro = tratar_banco(SPS.encode(), banco)
        ja_tratado = BancoCarregado(primeiro.execucao.df.copy(), "t.sav",
                                    dict(primeiro.execucao.var_labels),
                                    {k: dict(v) for k, v in primeiro.execucao.value_labels.items()})
        segundo = tratar_banco(SPS.encode(), ja_tratado, ja_tratado)
        self.assertEqual(segundo.divergentes, [])
        self.assertTrue(any(x.code == "TRAT_GERADAS_REMOVIDAS" for x in segundo.achados))


class ExportarSavTests(unittest.TestCase):
    def test_ida_e_volta_preserva_valores_e_rotulos(self):
        res = tratar_banco(SPS.encode(), _banco())
        dados = exportar_sav(res.execucao)
        lido = carregar_banco(dados, "t.sav", user_missing=True)
        for v in ("P1", "Q4R", "Q7R"):
            a, b = res.execucao.df[v], lido.df[v]
            self.assertTrue(((a == b) | (a.isna() & b.isna())).all(), v)
        self.assertEqual(lido.var_labels["P1"], "Sexo")
        self.assertEqual({float(k): x for k, x in lido.value_labels["Q4R"].items()}[97.0], "Branco/Nulo")

    def test_nome_do_arquivo(self):
        self.assertEqual(nome_arquivo_tratado("Banco_Original_Angra_dos_Reis_RJ_Set26.sav"),
                         "Banco_Tratado_Angra_dos_Reis_RJ_Set26.sav")
        self.assertEqual(nome_arquivo_tratado("meu.sav"), "meu_tratado.sav")


if __name__ == "__main__":
    unittest.main()
