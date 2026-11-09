"""Tratamento de Dados: expressões, operações estilo SPSS, histórico, syntax e ida e volta do .sav."""
import math
import unittest

import numpy as np
import pandas as pd

from core import tratamento_dados_math as td
from core.tratamento_dados_math import Banco, ErroTratamento


def _banco() -> Banco:
    df = pd.DataFrame({
        "ID": [1., 2., 3., 4., 5., 6.],
        "Q1": [1., 2., 3., 99., 1., np.nan],
        "Q2": [5., 4., np.nan, 2., 1., 3.],
        "Idade": [20., 35., 40., 52., 61., 28.],
        "Cidade": ["Rio", "Niterói", "Rio", None, "Rio", "Niterói"],
    })
    return Banco(
        df=df,
        rotulos_var={"Q1": "Satisfação"},
        rotulos_valor={"Q1": {1.0: "Baixa", 2.0: "Média", 3.0: "Alta"}},
        missing={"Q1": [{"lo": 99.0, "hi": 99.0}]},
        medida={"ID": "scale", "Q1": "ordinal", "Q2": "scale", "Idade": "scale", "Cidade": "nominal"},
        nome="teste.sav",
    )


def _col(b: Banco, nome: str):
    return b.df[nome].tolist()


def _igual(a, b):
    assert len(a) == len(b), (a, b)
    for x, y in zip(a, b):
        if isinstance(y, float) and math.isnan(y):
            assert isinstance(x, float) and math.isnan(x), (a, b)
        else:
            assert x == y, (a, b)


class TestExpressoes(unittest.TestCase):
    def setUp(self):
        self.b = _banco()

    def v(self, expr):
        r = td.avaliar(self.b, expr)
        return r.a.tolist()

    def test_aritmetica_e_precedencia(self):
        _igual(self.v("Idade + 2 * 3"), [26., 41., 46., 58., 67., 34.])
        _igual(self.v("(Idade + 2) * 3"), [66., 111., 126., 162., 189., 90.])
        _igual(self.v("-2 ** 2"), [-4.] * 6)
        _igual(self.v("2 ** 3 ** 2")[:1], [512.])

    def test_missing_de_usuario_vira_ausente_na_conta(self):
        _igual(self.v("Q1 + 1"), [2., 3., 4., math.nan, 2., math.nan])

    def test_missing_sysmis_value(self):
        _igual(self.v("MISSING(Q1)"), [0., 0., 0., 1., 0., 1.])
        _igual(self.v("SYSMIS(Q1)"), [0., 0., 0., 0., 0., 1.])
        _igual(self.v("VALUE(Q1)"), [1., 2., 3., 99., 1., math.nan])

    def test_logica_de_tres_valores(self):
        _igual(self.v("Q1 = 1"), [1., 0., 0., math.nan, 1., math.nan])
        _igual(self.v("NOT (Q1 = 1)"), [0., 1., 1., math.nan, 0., math.nan])
        _igual(self.v("Q1 = 1 OR Idade > 50"), [1., 0., 0., 1., 1., math.nan])
        # ausente AND verdadeiro = ausente; ausente AND falso = falso
        _igual(self.v("Q1 = 3 AND Idade > 50"), [0., 0., 0., math.nan, 0., 0.])
        _igual(self.v("Q1 EQ 1 & Idade GE 20"), [1., 0., 0., math.nan, 1., math.nan])

    def test_any_e_range(self):
        _igual(self.v("ANY(Q1, 1, 3)"), [1., 0., 1., math.nan, 1., math.nan])
        _igual(self.v("RANGE(Idade, 30, 50, 60, 70)"), [0., 1., 1., 0., 1., 0.])

    def test_funcoes_agregadas_com_to_e_minimo_valido(self):
        # Q1 TO Q2 = Q1, Q2 (ordem do arquivo; ID fica fora)
        _igual(self.v("MEAN(Q1 TO Q2)"), [3., 3., 3., 2., 1., 3.])
        _igual(self.v("SUM(Q1, Q2)"), [6., 6., 3., 2., 2., 3.])
        _igual(self.v("MEAN.2(Q1, Q2)"), [3., 3., math.nan, math.nan, 1., math.nan])
        _igual(self.v("NVALID(Q1, Q2)"), [2., 2., 1., 1., 2., 1.])
        _igual(self.v("MAX(Q1, Q2, 4)"), [5., 4., 4., 4., 4., 4.])

    def test_funcoes_matematicas(self):
        _igual(self.v("RND(2.5)")[:1], [3.])
        _igual(self.v("RND(-2.5)")[:1], [-3.])
        _igual(self.v("TRUNC(2.9)")[:1], [2.])
        _igual(self.v("ABS(-4)")[:1], [4.])
        _igual(self.v("SQRT(16)")[:1], [4.])
        _igual(self.v("MOD(Idade, 7)"), [6., 0., 5., 3., 5., 0.])
        self.assertTrue(math.isnan(self.v("1 / 0")[0]))

    def test_texto(self):
        r = td.avaliar(self.b, "UPCASE(Cidade)")
        self.assertEqual(r.a.tolist(), ["RIO", "NITERÓI", "RIO", None, "RIO", "NITERÓI"])
        r = td.avaliar(self.b, "CONCAT(Cidade, '-', 'X')")
        self.assertEqual(r.a[0], "Rio-X")
        r = td.avaliar(self.b, "SUBSTR(Cidade, 1, 2)")
        self.assertEqual(r.a.tolist()[:2], ["Ri", "Ni"])
        _igual(self.v("LENGTH(Cidade)"), [3., 7., 3., math.nan, 3., 7.])
        _igual(self.v("Cidade = 'Rio'"), [1., 0., 1., math.nan, 1., 0.])

    def test_casenum_e_sysmis(self):
        _igual(self.v("$CASENUM"), [1., 2., 3., 4., 5., 6.])
        self.assertTrue(math.isnan(self.v("$SYSMIS")[0]))

    def test_ponto_final_e_nomes_sem_caixa(self):
        _igual(self.v("idade + 1."), [21., 36., 41., 53., 62., 29.])

    def test_erros_claros(self):
        for expr in ["Inexistente + 1", "Idade +", "Idade ) 2", "Cidade + 1", "Idade = 'a'", "FUNCAOFALSA(1)",
                     "ANY(Idade)", "1 ? 2", ""]:
            with self.assertRaises(ErroTratamento, msg=expr):
                td.avaliar(self.b, expr)

    def test_nao_executa_codigo_python(self):
        for expr in ["__import__('os').system('echo oi')", "open('x')", "().__class__", "exec('1')"]:
            with self.assertRaises(ErroTratamento, msg=expr):
                td.avaliar(self.b, expr)


class TestCompute(unittest.TestCase):
    def test_compute_novo_com_rotulo_e_sintaxe(self):
        b = _banco()
        r = td.compute(b, "IdadeMes", "Idade * 12", rotulo="Idade em meses")
        _igual(_col(r.banco, "IdadeMes"), [240., 420., 480., 624., 732., 336.])
        self.assertEqual(r.banco.rotulos_var["IdadeMes"], "Idade em meses")
        self.assertEqual(r.banco.medida["IdadeMes"], "scale")
        self.assertIn("COMPUTE IdadeMes = Idade * 12.", r.sintaxe)
        self.assertIn("VARIABLE LABELS IdadeMes 'Idade em meses'.", r.sintaxe)
        self.assertNotIn("IdadeMes", b.df.columns)  # o original não muda

    def test_if_so_altera_onde_verdadeiro(self):
        b = _banco()
        r = td.compute(b, "Faixa", "1", condicao="Idade < 40")
        _igual(_col(r.banco, "Faixa"), [1., 1., math.nan, math.nan, math.nan, 1.])
        self.assertIn("IF (Idade < 40) Faixa = 1.", r.sintaxe)
        r2 = td.compute(r.banco, "Faixa", "2", condicao="Idade >= 40")
        _igual(_col(r2.banco, "Faixa"), [1., 1., 2., 2., 2., 1.])
        self.assertTrue(any("sobrescrita" in a for a in r2.avisos))

    def test_if_com_condicao_ausente_nao_atribui(self):
        r = td.compute(_banco(), "X", "1", condicao="Q1 = 1")
        _igual(_col(r.banco, "X"), [1., math.nan, math.nan, math.nan, 1., math.nan])

    def test_compute_texto_novo(self):
        r = td.compute(_banco(), "Cid2", "UPCASE(Cidade)")
        self.assertEqual(r.banco.df["Cid2"].tolist()[0], "RIO")
        self.assertIn("STRING Cid2 (A7).", r.sintaxe)
        self.assertEqual(r.banco.medida["Cid2"], "nominal")

    def test_tipos_incompativeis_e_nomes_invalidos(self):
        b = _banco()
        with self.assertRaises(ErroTratamento):
            td.compute(b, "Idade", "UPCASE(Cidade)")
        with self.assertRaises(ErroTratamento):
            td.compute(b, "Cidade", "1")
        for nome in ["1abc", "a b", "AND", "", "x" * 70, "abc."]:
            with self.assertRaises(ErroTratamento, msg=nome):
                td.compute(b, nome, "1")


class TestRecode(unittest.TestCase):
    def test_no_lugar_mantem_nao_citados(self):
        r = td.recode(_banco(), ["Idade"], "(LO THRU 29 = 1) (30 THRU 49 = 2)")
        _igual(_col(r.banco, "Idade"), [1., 2., 2., 52., 61., 1.])
        self.assertIn("RECODE Idade (LO THRU 29=1) (30 THRU 49=2).", r.sintaxe)

    def test_into_novo_nao_citados_viram_sysmis(self):
        r = td.recode(_banco(), ["Idade"], "(LO THRU 29 = 1) (30 THRU 49 = 2)", destinos=["Faixa"])
        _igual(_col(r.banco, "Faixa"), [1., 2., 2., math.nan, math.nan, 1.])
        _igual(_col(r.banco, "Idade"), [20., 35., 40., 52., 61., 28.])   # origem intacta
        self.assertIn("INTO Faixa.", r.sintaxe)

    def test_else_copy_e_else_valor(self):
        r = td.recode(_banco(), ["Idade"], "(LO THRU 29 = 1) (ELSE = COPY)", destinos=["F2"])
        _igual(_col(r.banco, "F2"), [1., 35., 40., 52., 61., 1.])
        r = td.recode(_banco(), ["Idade"], "(LO THRU 29 = 1) (ELSE = 9)", destinos=["F3"])
        _igual(_col(r.banco, "F3"), [1., 9., 9., 9., 9., 1.])

    def test_primeira_regra_que_casa_vale_e_troca_de_valores(self):
        r = td.recode(_banco(), ["Q2"], "(1 = 2) (2 = 1)")
        _igual(_col(r.banco, "Q2"), [5., 4., math.nan, 1., 2., 3.])
        r = td.recode(_banco(), ["Idade"], "(LO THRU 40 = 1) (30 THRU 50 = 2) (ELSE = 3)", destinos=["F"])
        _igual(_col(r.banco, "F"), [1., 1., 1., 3., 3., 1.])

    def test_missing_e_sysmis_e_user_missing(self):
        # MISSING pega 99 (user) e ausente (sysmis); SYSMIS só o ausente
        r = td.recode(_banco(), ["Q1"], "(MISSING = 0) (ELSE = COPY)", destinos=["A"])
        _igual(_col(r.banco, "A"), [1., 2., 3., 0., 1., 0.])
        r = td.recode(_banco(), ["Q1"], "(SYSMIS = 0) (ELSE = COPY)", destinos=["B"])
        _igual(_col(r.banco, "B"), [1., 2., 3., 99., 1., 0.])
        # valor explícito também enxerga o código bruto 99
        r = td.recode(_banco(), ["Q1"], "(99 = 9)")
        _igual(_col(r.banco, "Q1"), [1., 2., 3., 9., 1., math.nan])

    def test_ate_valor_inclui_user_missing_mas_nao_sysmis(self):
        r = td.recode(_banco(), ["Q1"], "(2 THRU HI = 5)", destinos=["Z"])
        _igual(_col(r.banco, "Z"), [math.nan, 5., 5., 5., math.nan, math.nan])

    def test_into_existente_preserva_o_que_nao_foi_citado(self):
        b = td.compute(_banco(), "Alvo", "-1").banco
        r = td.recode(b, ["Idade"], "(LO THRU 29 = 1)", destinos=["Alvo"])
        _igual(_col(r.banco, "Alvo"), [1., -1., -1., -1., -1., 1.])

    def test_varias_origens_e_destinos(self):
        r = td.recode(_banco(), ["Q2", "Idade"], "(LO THRU 10 = 1) (ELSE = 0)", destinos=["a", "b"])
        _igual(_col(r.banco, "a"), [1., 1., 0., 1., 1., 1.])     # ELSE também pega o ausente
        _igual(_col(r.banco, "b"), [0., 0., 0., 0., 0., 0.])

    def test_texto(self):
        r = td.recode(_banco(), ["Cidade"], "('Rio' = 'RJ') ('Niterói' = 'NIT')", destinos=["UF"])
        self.assertEqual(r.banco.df["UF"].tolist(), ["RJ", "NIT", "RJ", None, "RJ", "NIT"])
        r = td.recode(_banco(), ["Cidade"], "('Rio' = 1) (ELSE = 2)", destinos=["N"])
        _igual(_col(r.banco, "N"), [1., 2., 1., 2., 1., 2.])

    def test_regras_por_linha_e_rotulos(self):
        r = td.recode(_banco(), ["Idade"], "LO THRU 29 = 1\n30 THRU HI = 2", destinos=["F"],
                      rotulo_variavel="Faixa etária", rotulos_valor={1: "Jovem", 2: "Adulto"})
        _igual(_col(r.banco, "F"), [1., 2., 2., 2., 2., 1.])
        self.assertEqual(r.banco.rotulos_var["F"], "Faixa etária")
        self.assertEqual(r.banco.rotulos_valor["F"], {1.0: "Jovem", 2.0: "Adulto"})
        self.assertIn("VALUE LABELS F 1 'Jovem' 2 'Adulto'.", r.sintaxe)

    def test_erros(self):
        b = _banco()
        for regras, kw in [("", {}), ("(1 2)", {}), ("(ELSE = 1) (2 = 3)", {}), ("(1 = 'a') (2 = 3)", {}),
                           ("(1 THRU = 2)", {}), ("('a' = 1)", {})]:
            with self.assertRaises(ErroTratamento, msg=regras):
                td.recode(b, ["Q2"], regras, **kw)
        with self.assertRaises(ErroTratamento):       # trocar tipo exige INTO
            td.recode(b, ["Cidade"], "('Rio' = 1)")
        with self.assertRaises(ErroTratamento):       # destinos de menos
            td.recode(b, ["Q2", "Idade"], "(1 = 2)", destinos=["x"])


class TestCasosEVariaveis(unittest.TestCase):
    def test_select_if_manter(self):
        r = td.selecionar_casos(_banco(), "Idade >= 40")
        _igual(_col(r.banco, "ID"), [3., 4., 5.])
        self.assertIn("SELECT IF (Idade >= 40).", r.sintaxe)

    def test_excluir_remove_tambem_os_ausentes_e_avisa(self):
        r = td.selecionar_casos(_banco(), "Q1 = 1", modo="excluir")
        _igual(_col(r.banco, "ID"), [2., 3.])        # 1 e 5 (verdadeiro) saem; 4 e 6 (ausentes) também
        self.assertTrue(any("ausente" in a for a in r.avisos))
        self.assertIn("SELECT IF NOT (Q1 = 1).", r.sintaxe)

    def test_ordenar_varias_chaves_e_ausentes(self):
        r = td.ordenar_casos(_banco(), [("Cidade", "A"), ("Idade", "D")])
        self.assertEqual(_col(r.banco, "ID"), [4., 2., 6., 5., 3., 1.])
        self.assertIn("SORT CASES BY Cidade (A) Idade (D).", r.sintaxe)
        r = td.ordenar_casos(_banco(), [("Q2", "A")])
        self.assertEqual(_col(r.banco, "ID")[0], 3.)       # ausente primeiro (crescente)
        r = td.ordenar_casos(_banco(), [("Q2", "D")])
        self.assertEqual(_col(r.banco, "ID")[-1], 3.)      # ausente por último (decrescente)

    def test_remover_duplicados(self):
        b = td.compute(_banco(), "Chave", "1", condicao="Cidade = 'Rio'").banco
        r = td.remover_duplicados(b, ["Cidade"], manter="primeiro")
        self.assertEqual(len(r.banco.df), 3)    # Rio, Niterói e o ausente
        self.assertIn("/FIRST=PrimeiroCaso", r.sintaxe)
        r2 = td.remover_duplicados(_banco(), ["Cidade"], manter="ultimo")
        self.assertEqual(sorted(_col(r2.banco, "ID")), [4., 5., 6.])
        self.assertIn("/LAST=UltimoCaso", r2.sintaxe)

    def test_excluir_e_renomear_variaveis_levam_metadados(self):
        r = td.renomear_variavel(_banco(), "q1", "Satisf")
        self.assertIn("Satisf", r.banco.df.columns)
        self.assertEqual(r.banco.rotulos_var["Satisf"], "Satisfação")
        self.assertIn("Satisf", r.banco.missing)
        self.assertNotIn("Q1", r.banco.rotulos_valor)
        with self.assertRaises(ErroTratamento):
            td.renomear_variavel(_banco(), "Q1", "Q2")
        r = td.excluir_variaveis(_banco(), ["Q1", "Q2"])
        self.assertNotIn("Q1", r.banco.df.columns)
        self.assertNotIn("Q1", r.banco.rotulos_var)
        self.assertIn("DELETE VARIABLES Q1 Q2.", r.sintaxe)
        with self.assertRaises(ErroTratamento):
            td.excluir_variaveis(_banco(), list(_banco().df.columns))


class TestRotulosMissingsMedida(unittest.TestCase):
    def test_rotulo_de_variavel(self):
        r = td.rotular_variavel(_banco(), "Idade", "Idade em anos")
        self.assertEqual(r.banco.rotulos_var["Idade"], "Idade em anos")
        self.assertIn("VARIABLE LABELS Idade 'Idade em anos'.", r.sintaxe)
        r = td.rotular_variavel(r.banco, "Idade", "")
        self.assertNotIn("Idade", r.banco.rotulos_var)

    def test_rotulos_de_valor(self):
        mapa = td.analisar_rotulos_valor("1 = Baixo\n2 = 'Médio alto'\n3\tAlto", numerico=True)
        self.assertEqual(mapa, {1.0: "Baixo", 2.0: "Médio alto", 3.0: "Alto"})
        r = td.rotular_valores(_banco(), ["Q2"], mapa)
        self.assertEqual(r.banco.rotulos_valor["Q2"][2.0], "Médio alto")
        self.assertIn("VALUE LABELS Q2 1 'Baixo' 2 'Médio alto' 3 'Alto'.", r.sintaxe)
        r = td.rotular_valores(_banco(), ["Q1"], {99: "Não sabe"}, modo="adicionar")
        self.assertEqual(len(r.banco.rotulos_valor["Q1"]), 4)
        self.assertIn("ADD VALUE LABELS Q1 99 'Não sabe'.", r.sintaxe)
        r = td.rotular_valores(_banco(), ["Q1"], {}, modo="remover")
        self.assertNotIn("Q1", r.banco.rotulos_valor)
        with self.assertRaises(ErroTratamento):
            td.analisar_rotulos_valor("abc = x", numerico=True)
        with self.assertRaises(ErroTratamento):
            td.rotular_valores(_banco(), ["Q2", "Cidade"], {1: "a"})   # tipos misturados
        m = td.analisar_rotulos_valor("R = Rio", numerico=False)
        self.assertEqual(m, {"R": "Rio"})

    def test_missing_valores_faixas_e_limites(self):
        self.assertEqual(td.analisar_missing("99, 98"), [{"lo": 99.0, "hi": 99.0}, {"lo": 98.0, "hi": 98.0}])
        f = td.analisar_missing("LO THRU 0, 99")
        self.assertEqual(f[0]["lo"], td.LOWEST)
        self.assertEqual(td.analisar_missing(""), [])
        for ruim in ["1,2,3,4", "1 THRU 2, 5 THRU 6", "5 THRU 1", "a", "1 THRU"]:
            with self.assertRaises(ErroTratamento, msg=ruim):
                td.analisar_missing(ruim)
        r = td.definir_missing(_banco(), ["Q2"], "5")
        self.assertEqual(r.banco.missing["Q2"], [{"lo": 5.0, "hi": 5.0}])
        self.assertIn("MISSING VALUES Q2 (5).", r.sintaxe)
        _igual(td.avaliar(r.banco, "MEAN(Q2)").a.tolist(), [math.nan, 4., math.nan, 2., 1., 3.])
        r2 = td.definir_missing(r.banco, ["Q2"], "")
        self.assertNotIn("Q2", r2.banco.missing)
        r3 = td.definir_missing(_banco(), ["Cidade"], "'Rio'")
        self.assertEqual(r3.banco.missing["Cidade"], [{"lo": "Rio", "hi": "Rio"}])
        self.assertIn("MISSING VALUES Cidade ('Rio').", r3.sintaxe)

    def test_medida(self):
        r = td.definir_medida(_banco(), ["Q2", "Idade"], "ordinal")
        self.assertEqual(r.banco.medida["Q2"], "ordinal")
        self.assertIn("VARIABLE LEVEL Q2 Idade (ORDINAL).", r.sintaxe)
        with self.assertRaises(ErroTratamento):
            td.definir_medida(_banco(), ["Q2"], "xyz")


class TestMerge(unittest.TestCase):
    def _outro(self):
        df = pd.DataFrame({"id": [2., 4., 9.], "Renda": [1000., 2000., 3000.], "Idade": [99., 99., 99.]})
        return Banco(df=df, rotulos_var={"Renda": "Renda mensal"}, medida={"id": "scale", "Renda": "scale", "Idade": "scale"},
                     nome="renda.sav")

    def test_add_files_com_indicador_e_variaveis_so_de_um_lado(self):
        b = _banco()
        outro = Banco(df=pd.DataFrame({"ID": [7., 8.], "Idade": [30., 31.], "Extra": [1., 2.]}), nome="mais.sav",
                      rotulos_var={"Extra": "Extra!"})
        r = td.adicionar_casos(b, outro, indicador="Origem")
        self.assertEqual(len(r.banco.df), 8)
        _igual(_col(r.banco, "Origem"), [0.] * 6 + [1.] * 2)
        _igual(_col(r.banco, "Extra"), [math.nan] * 6 + [1., 2.])
        self.assertEqual(r.banco.rotulos_var["Extra"], "Extra!")
        self.assertIn("ADD FILES /FILE=* /FILE='mais.sav' /IN=Origem.", r.sintaxe)
        self.assertEqual(len(r.avisos), 2)
        with self.assertRaises(ErroTratamento):
            td.adicionar_casos(b, Banco(df=pd.DataFrame({"Cidade": [1.]}), nome="x.sav"))
        with self.assertRaises(ErroTratamento):
            td.adicionar_casos(b, Banco(df=pd.DataFrame({"zzz": [1.]}), nome="x.sav"))

    def test_match_files_consulta(self):
        r = td.adicionar_variaveis(_banco(), self._outro(), ["ID"], "consulta")
        self.assertEqual(len(r.banco.df), 6)
        _igual(_col(r.banco, "Renda"), [math.nan, 1000., math.nan, 2000., math.nan, math.nan])
        _igual(_col(r.banco, "Idade"), [20., 35., 40., 52., 61., 28.])    # nome repetido: vale o ativo
        self.assertEqual(r.banco.rotulos_var["Renda"], "Renda mensal")
        self.assertIn("/TABLE='renda.sav' /BY ID", r.sintaxe)
        self.assertTrue(any("mesmo nome" in a for a in r.avisos))

    def test_match_files_ambos_e_chave_duplicada_e_sem_chave(self):
        r = td.adicionar_variaveis(_banco(), self._outro(), ["ID"], "ambos")
        self.assertEqual(len(r.banco.df), 7)     # +caso 9 só do segundo banco
        self.assertIn("/FILE='renda.sav' /BY ID", r.sintaxe)
        dup = self._outro()
        dup.df.loc[1, "id"] = 2.
        with self.assertRaises(ErroTratamento):
            td.adicionar_variaveis(_banco(), dup, ["ID"], "consulta")
        r = td.adicionar_variaveis(_banco(), self._outro(), [], "consulta", trazer=["Renda"])
        self.assertEqual(len(r.banco.df), 6)
        self.assertTrue(any("ORDEM" in a for a in r.avisos))
        with self.assertRaises(ErroTratamento):
            td.adicionar_variaveis(_banco(), self._outro(), ["Cidade"], "consulta")


class TestAggregate(unittest.TestCase):
    def test_adicionar_variaveis_por_grupo(self):
        r = td.agregar(_banco(), ["Cidade"], [{"novo": "IdadeMedia", "funcao": "MEAN", "variavel": "Idade"},
                                              {"novo": "N_grupo", "funcao": "N"}], modo="adicionar")
        _igual(_col(r.banco, "IdadeMedia"), [(20 + 40 + 61) / 3, (35 + 28) / 2, (20 + 40 + 61) / 3, 52., (20 + 40 + 61) / 3, (35 + 28) / 2])
        _igual(_col(r.banco, "N_grupo"), [3., 2., 3., 1., 3., 2.])
        self.assertIn("MODE=ADDVARIABLES", r.sintaxe)
        self.assertIn("/BREAK=Cidade", r.sintaxe)
        self.assertIn("/IdadeMedia = MEAN(Idade)", r.sintaxe)
        self.assertIn("/N_grupo = N.", r.sintaxe)

    def test_substituir_pelo_banco_agregado(self):
        r = td.agregar(_banco(), ["Cidade"], [{"novo": "Soma", "funcao": "SUM", "variavel": "Idade"},
                                              {"novo": "Ausentes", "funcao": "NMISS", "variavel": "Q1"},
                                              {"novo": "Validos", "funcao": "NVALID", "variavel": "Q1"},
                                              {"novo": "Maior", "funcao": "MAX", "variavel": "Idade"},
                                              {"novo": "DP", "funcao": "SD", "variavel": "Idade"},
                                              {"novo": "Med", "funcao": "MEDIAN", "variavel": "Idade"},
                                              {"novo": "Prim", "funcao": "FIRST", "variavel": "Idade"},
                                              {"novo": "Ult", "funcao": "LAST", "variavel": "Idade"}], modo="substituir")
        d = r.banco.df
        self.assertEqual(len(d), 3)
        self.assertTrue(pd.isna(d["Cidade"].iloc[0]))                  # ausente primeiro
        lin = d[d["Cidade"] == "Rio"].iloc[0]
        self.assertEqual((lin["Soma"], lin["Maior"], lin["Med"], lin["Prim"], lin["Ult"]), (121., 61., 40., 20., 61.))
        self.assertAlmostEqual(lin["DP"], np.std([20, 40, 61], ddof=1))
        self.assertEqual((lin["Ausentes"], lin["Validos"]), (0., 3.))
        nit = d[d["Cidade"] == "Niterói"].iloc[0]
        self.assertEqual((nit["Ausentes"], nit["Validos"]), (1., 1.))     # Q1 ausente em 1 dos 2 casos
        sem = d[d["Cidade"].isna()].iloc[0]
        self.assertEqual((sem["Ausentes"], sem["Validos"]), (1., 0.))     # 99 é user-missing em Q1
        self.assertNotIn("Idade", d.columns)
        self.assertNotIn("/OUTFILE=* MODE", r.sintaxe)

    def test_sem_quebra_e_erros(self):
        r = td.agregar(_banco(), [], [{"novo": "Tot", "funcao": "SUM", "variavel": "Idade"}], modo="substituir")
        self.assertEqual(r.banco.df["Tot"].tolist(), [236.])
        for ag in [[{"novo": "x", "funcao": "MEAN", "variavel": "Cidade"}],
                   [{"novo": "x", "funcao": "XYZ", "variavel": "Idade"}],
                   [{"novo": "Idade", "funcao": "SUM", "variavel": "Idade"}], []]:
            with self.assertRaises(ErroTratamento):
                td.agregar(_banco(), ["Cidade"], ag, modo="adicionar")


class TestHistoricoEIO(unittest.TestCase):
    def test_historico_desfazer_e_sintaxe(self):
        original = _banco()
        b, p1 = td.aplicar(original, "compute", alvo="Dobro", expressao="Idade * 2")
        b, p2 = td.aplicar(b, "recode", origens=["Dobro"], regras="(LO THRU 80 = 1) (ELSE = 2)", destinos=["Grupo"])
        b, p3 = td.aplicar(b, "selecionar_casos", condicao="Grupo = 2")
        self.assertEqual(len(b.df), 2)
        desfeito = td.reaplicar(original, [p1, p2])
        self.assertEqual(len(desfeito.df), 6)
        self.assertIn("Grupo", desfeito.df.columns)
        refeito = td.reaplicar(original, [p1, p2, p3])
        pd.testing.assert_frame_equal(refeito.df, b.df)
        sint = td.gerar_sintaxe("teste.sav", [p1, p2, p3])
        self.assertTrue(sint.startswith("* Syntax gerada"))
        self.assertIn("GET FILE='teste.sav'.", sint)
        self.assertIn("COMPUTE Dobro = Idade * 2.", sint)
        self.assertIn("* Passo 3: SELECT IF (Grupo = 2)", sint)
        self.assertTrue(sint.rstrip().endswith("SAVE OUTFILE='teste_tratado.sav'."))
        with self.assertRaises(ErroTratamento):
            td.aplicar(original, "nao_existe")

    def test_ida_e_volta_sav_preserva_rotulos_missings_e_medida(self):
        b = _banco()
        b.df["Cidade"] = b.df["Cidade"].fillna("")      # no .sav, texto vazio é "em branco" (não existe nulo)
        b, _ = td.aplicar(b, "definir_missing", nomes=["Q2"], especificacao="98 THRU 99")
        dados, avisos = td.gravar_sav(b)
        self.assertEqual(avisos, [])
        lido = td.carregar_banco(dados, "volta.sav")
        self.assertEqual(lido.colunas, b.colunas)
        pd.testing.assert_frame_equal(lido.df.reset_index(drop=True), b.df.reset_index(drop=True), check_dtype=False)
        self.assertEqual(lido.rotulos_var.get("Q1"), "Satisfação")
        self.assertEqual(lido.rotulos_valor["Q1"], {1.0: "Baixa", 2.0: "Média", 3.0: "Alta"})
        self.assertEqual(lido.missing["Q1"], [{"lo": 99.0, "hi": 99.0}])
        self.assertEqual(lido.missing["Q2"], [{"lo": 98.0, "hi": 99.0}])
        self.assertEqual(lido.medida["Q1"], "ordinal")
        self.assertEqual(lido.medida["Cidade"], "nominal")

    def test_exportacoes(self):
        import io
        b = _banco()
        xlsx = td.exportar_excel(b, com_rotulos=True)
        dados = pd.read_excel(io.BytesIO(xlsx), sheet_name="Dados")
        self.assertEqual(dados["Q1"].tolist()[:3], ["Baixa", "Média", "Alta"])
        variaveis = pd.read_excel(io.BytesIO(xlsx), sheet_name="Variáveis")
        self.assertEqual(variaveis["Variável"].tolist(), b.colunas)
        csv = td.exportar_csv(b).decode("utf-8-sig")
        self.assertTrue(csv.splitlines()[0].startswith("ID;Q1;Q2"))
        tab = td.tabela_variaveis(b)
        self.assertEqual(tab.loc[tab["Variável"] == "Q1", "Missing"].iloc[0], "99")


class TestNumberComFormatoDoSpss(unittest.TestCase):
    def test_formato_sem_aspas_e_ignorado(self):
        b = td.Banco(df=pd.DataFrame({"a": ["12", "7", "x"]}), nome="x.sav")
        for expr in ("NUMBER(a, F8)", "NUMBER(a, F8.2)", "NUMBER(a, 'F8')"):
            r = td.compute(b, "r", expr).banco.df["r"].tolist()
            self.assertEqual(r[:2], [12.0, 7.0], expr)
            self.assertTrue(np.isnan(r[2]), expr)


if __name__ == "__main__":
    unittest.main()
