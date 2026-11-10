import unittest

import numpy as np
import pandas as pd

from core.syntax_executor_math import (
    ErroDeSyntax,
    aplicar_regras,
    executar_syntax,
    parse_recode_completo,
    variaveis_geradas,
)


def _sps(texto: str) -> bytes:
    return texto.encode("utf-8")


def _df():
    return pd.DataFrame({
        "SEXO": [1.0, 2.0, 2.0, np.nan],
        "IDADE": [1.0, 3.0, 6.0, 5.0],
        "Q1": [1.0, 2.0, 3.0, 9999.0],
        "ESC": [1.0, 2.0, 5.0, 9.0],
    })


class ParseRecodeTests(unittest.TestCase):
    def test_into_origem_destino_e_regras(self):
        c = parse_recode_completo(1, "recode SEXO (1=1) (2=2) into P1.")
        self.assertEqual(c.origens, ["SEXO"])
        self.assertEqual(c.destinos, ["P1"])
        self.assertTrue(c.com_into)
        self.assertEqual(len(c.regras[0]), 2)

    def test_sem_into_grava_no_lugar(self):
        c = parse_recode_completo(1, "RECODE Q1 (1=2).")
        self.assertEqual(c.destinos, ["Q1"])
        self.assertFalse(c.com_into)

    def test_faixa_lista_else_e_copy(self):
        c = parse_recode_completo(1, "recode ESC (1=1) (2 thru 3=2) (4, 5=3) (LO THRU 0=SYSMIS) (else=copy) into P4.")
        tipos = [[x.tipo for x in r.condicoes] for r in c.regras[0]]
        self.assertEqual(tipos, [["val"], ["faixa"], ["val", "val"], ["faixa"], ["else"]])

    def test_varias_origens_e_barra(self):
        c = parse_recode_completo(1, "RECODE A B (1=2) /C (3=4) INTO X Y Z.")
        self.assertEqual(c.origens, ["A", "B", "C"])
        self.assertEqual(c.destinos, ["X", "Y", "Z"])

    def test_into_com_quantidade_errada(self):
        with self.assertRaises(ErroDeSyntax):
            parse_recode_completo(1, "RECODE A B (1=2) INTO X.")

    def test_valor_decimal_e_texto(self):
        c = parse_recode_completo(1, "RECODE A (1.5=2) ('x'='y') INTO B.")
        self.assertEqual(c.regras[0][0].condicoes[0].a, 1.5)
        self.assertEqual(c.regras[0][1].valor, "y")


class AplicarRegrasTests(unittest.TestCase):
    def _rodar(self, recode, serie, atual=None):
        cmd = parse_recode_completo(1, recode)
        return aplicar_regras(serie, cmd.regras[0], atual)

    def test_primeira_regra_vence_sem_cascata(self):
        # (13=11)(11=90): o 13 vira 11 e NÃO é recodificado de novo para 90
        s = pd.Series([13.0, 11.0, 12.0])
        res, _ = self._rodar("RECODE X (13=11) (11=90) (ELSE=COPY) INTO Y.", s)
        self.assertEqual(res.tolist(), [11.0, 90.0, 12.0])

    def test_nao_citado_com_into_vira_sysmis(self):
        s = pd.Series([1.0, 2.0, 7.0])
        res, info = self._rodar("RECODE X (1=1) (2=2) INTO Y.", s)
        self.assertTrue(np.isnan(res.iloc[2]))
        self.assertEqual(info["nao_citados"], 1)

    def test_nao_citado_no_lugar_fica_como_estava(self):
        s = pd.Series([1.0, 2.0, 7.0])
        res, _ = self._rodar("RECODE X (1=10) INTO Y.", s, atual=s.copy())
        self.assertEqual(res.tolist(), [10.0, 2.0, 7.0])

    def test_else_sysmis_e_faixa(self):
        s = pd.Series([1.0, 3.0, 6.0, np.nan])
        res, _ = self._rodar("RECODE X (1 thru 5=copy) (else=sysmis) INTO Y.", s)
        self.assertEqual(res.iloc[:3].tolist()[0:2], [1.0, 3.0])
        self.assertTrue(np.isnan(res.iloc[2]) and np.isnan(res.iloc[3]))

    def test_lo_hi_e_missing(self):
        s = pd.Series([-5.0, 0.0, 50.0, np.nan])
        res, _ = self._rodar("RECODE X (LO THRU 0=1) (1 THRU HI=2) (MISSING=9) INTO Y.", s)
        self.assertEqual(res.tolist(), [1.0, 1.0, 2.0, 9.0])

    def test_texto(self):
        s = pd.Series(["a", "b", ""])
        res, _ = self._rodar("RECODE X ('a'='sim') ('b'='nao') INTO Y.", s)
        self.assertEqual(res.tolist()[:2], ["sim", "nao"])


class ExecutarSyntaxTests(unittest.TestCase):
    SPS = """
* Encoding: UTF-8.
Compute total=1.
var label total ''.
value label /total 1 'Total'.

recode SEXO (1=1) (2=2) into P1.
var label P1 'Sexo'.
value label /P1
1 'Masculino'
2 'Feminino'.
EXECUTE.

recode ESC (1=1) (2 thru 3=2) (4 thru 5=3) (6 thru 8=4) (9=1) (else=sysmis) into P4.
var label P4 'Escolaridade'.
value label /P4 1 'Sem instrução' 2 'Fund.' 3 'Médio' 4 'Superior'.

recode Q1 (1 thru 3=copy) into Q1R.
VARIABLE LABELS Q1R 'Pergunta 1'.
VALUE LABELS Q1R 1 'Branco/Nulo' 2 'Sim/Não' 3 'Outro'.
"""

    def test_cria_variaveis_com_valores_e_rotulos(self):
        r = executar_syntax(_sps(self.SPS), _df())
        self.assertEqual(r.df["P1"].tolist()[:3], [1.0, 2.0, 2.0])
        self.assertTrue(np.isnan(r.df["P1"].iloc[3]))
        self.assertEqual(r.df["TOTAL"].tolist(), [1.0] * 4)
        self.assertEqual(r.df["P4"].dropna().tolist(), [1.0, 2.0, 3.0, 1.0])
        self.assertEqual(r.var_labels["P1"], "Sexo")
        self.assertEqual(r.value_labels["P1"], {1: "Masculino", 2: "Feminino"})
        self.assertEqual(r.criadas, ["TOTAL", "P1", "P4", "Q1R"])

    def test_rotulo_com_barra_nao_e_cortado(self):
        r = executar_syntax(_sps(self.SPS), _df())
        self.assertEqual(r.value_labels["Q1R"], {1: "Branco/Nulo", 2: "Sim/Não", 3: "Outro"})

    def test_nao_altera_o_banco_original(self):
        df = _df()
        executar_syntax(_sps(self.SPS), df)
        self.assertNotIn("P1", df.columns)

    def test_variavel_de_origem_inexistente_e_critico(self):
        r = executar_syntax(_sps("recode NAOEXISTE (1=1) into X."), _df())
        self.assertTrue(any(f.code == "EXEC_ORIGEM_AUSENTE" and f.severity == "critico" for f in r.findings))
        self.assertEqual(r.registros[0].status, "erro")

    def test_comando_desconhecido_nao_passa_em_silencio(self):
        r = executar_syntax(_sps("IF (SEXO=1) X=2.\nSELECT IF (SEXO=1).\nFREQUENCIES SEXO.\nXYZ ABC."), _df())
        nao_rodou = [f for f in r.findings if f.code == "EXEC_COMANDO_NAO_SUPORTADO"]
        self.assertEqual(len(nao_rodou), 3)           # IF, SELECT IF e XYZ; FREQUENCIES só lê
        sev = {f.snippet.split()[0].upper(): f.severity for f in nao_rodou}
        self.assertEqual(sev["IF"], "possivel_erro")
        self.assertEqual(sev["SELECT"], "possivel_erro")
        self.assertEqual(sev["XYZ"], "revisar")
        self.assertEqual([x.status for x in r.registros if x.tipo == "FREQUENCIES"], ["ignorado"])

    def test_variavel_ja_existente_e_avisada(self):
        df = _df()
        df["P1"] = [9.0, 9.0, 9.0, 9.0]
        r = executar_syntax(_sps("recode SEXO (1=1) (2=2) into P1."), df)
        self.assertTrue(any(f.code == "EXEC_VARIAVEL_SOBRESCRITA" for f in r.findings))

    def test_valor_9999_sem_regra_e_informativo(self):
        r = executar_syntax(_sps("recode Q1 (1 thru 3=copy) into Q1R."), _df())
        f = [x for x in r.findings if x.code == "EXEC_VALOR_SEM_REGRA"]
        self.assertEqual(len(f), 1)
        self.assertEqual(f[0].severity, "informativo")
        self.assertIn("9999", f[0].message)

    def test_valor_comum_esquecido_e_revisar(self):
        df = pd.DataFrame({"A": [1.0] * 48 + [2.0] * 48 + [3.0, 4.0, 5.0, 6.0]})
        r = executar_syntax(_sps("recode A (1=1) (2=2) (3=3) into B."), df)
        f = [x for x in r.findings if x.code == "EXEC_VALOR_SEM_REGRA"]
        self.assertEqual(f[0].severity, "revisar")

    def test_mrsets_e_redefinicao(self):
        sps = """
recode Q1 (1=1) into A.
recode Q1 (2=2) into B.
MRSETS
  /MCGROUP NAME=$P6 LABEL=' ' VARIABLES= A B
  /DISPLAY NAME=[$P6].
MRSETS
  /MCGROUP NAME=$P6 LABEL=' ' VARIABLES= A
  /DISPLAY NAME=[$P6].
"""
        r = executar_syntax(_sps(sps), _df())
        self.assertEqual(r.mrsets["$P6"], ["A"])
        self.assertTrue(any(f.code == "EXEC_MRSET_REDEFINIDO" for f in r.findings))

    def test_compute_aritmetico(self):
        r = executar_syntax(_sps("compute dobro = Q1 * 2.\ncompute um = 1."), _df())
        self.assertEqual(r.df["DOBRO"].tolist()[:3], [2.0, 4.0, 6.0])

    def test_variaveis_geradas(self):
        self.assertEqual(variaveis_geradas(_sps(self.SPS)), ["TOTAL", "P1", "P4", "Q1R"])


if __name__ == "__main__":
    unittest.main()
