"""Partes puras da tela no estilo SPSS: menus, grades -> regras/syntax, inserção na expressão."""
import unittest

import numpy as np
import pandas as pd

from core import tratamento_dados_dialogos as dlg
from core import tratamento_dados_math as td
from core.tratamento_dados_math import ErroTratamento


def _banco():
    df = pd.DataFrame({"Q1": [1., 2., 3., 99., np.nan], "Idade": [20., 35., 40., 52., 61.],
                       "Cidade": ["Rio", "Niterói", "Rio", "", "Rio"]})
    return td.Banco(df=df, rotulos_var={"Q1": "Satisfação"}, medida={"Q1": "ordinal", "Idade": "scale",
                                                                   "Cidade": "nominal"}, nome="x.sav")


class TestFaixas(unittest.TestCase):
    def test_ids_unicos_e_todos_mapeados(self):
        ids = dlg.ids_dos_comandos()
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual(set(ids), set(dlg.OPERACAO_DA_JANELA))

    def test_todas_as_operacoes_do_nucleo_tem_comando(self):
        usadas = {op for op in dlg.OPERACAO_DA_JANELA.values() if op}
        self.assertEqual(usadas, set(td.OPERACOES))

    def test_ajuda_para_cada_janela_que_altera_o_banco(self):
        for ident, op in dlg.OPERACAO_DA_JANELA.items():
            if op or ident in ("salvar_sav", "exportar", "salvar_sintaxe"):
                self.assertIn(ident, dlg.AJUDA, f"sem texto de ajuda: {ident}")

    def test_faixas_por_tarefa_na_ordem_do_fluxo(self):
        self.assertEqual([f for f, _ in dlg.FAIXAS],
                         ["Variáveis", "Casos", "Rótulos e ausentes", "Juntar bancos", "Arquivo", "Ajuda"])

    def test_cada_comando_tem_icone_rotulo_e_descricao_curta(self):
        for faixa, grupos in dlg.FAIXAS:
            self.assertTrue(grupos, faixa)
            for grupo, comandos in grupos:
                self.assertTrue(grupo and comandos, f"{faixa}/{grupo}")
                for ident, icone, rotulo, desc in comandos:
                    self.assertTrue(icone and rotulo and desc, ident)
                    self.assertRegex(icone, r"^[a-z_]+$")                  # nome de ícone Material
                    self.assertLessEqual(len(desc), 60, f"descrição longa demais: {ident}")
        rotulos = [c[2] for _, g in dlg.FAIXAS for _, cs in g for c in cs]
        self.assertEqual(len(rotulos), len(set(rotulos)), "rótulos repetidos")

    def test_acoes_diretas_nao_estao_nas_faixas(self):
        self.assertFalse(dlg.ACOES_DIRETAS & set(dlg.ids_dos_comandos()))


class TestInserir(unittest.TestCase):
    def test_espacamento(self):
        self.assertEqual(dlg.inserir_trecho("", "Idade"), "Idade")
        self.assertEqual(dlg.inserir_trecho("Idade", "*"), "Idade *")
        self.assertEqual(dlg.inserir_trecho("Idade *", "12"), "Idade * 12")
        self.assertEqual(dlg.inserir_trecho("MEAN(", "Q1"), "MEAN(Q1")
        self.assertEqual(dlg.inserir_trecho("MEAN(Q1", ")"), "MEAN(Q1)")
        self.assertEqual(dlg.inserir_trecho("ANY(x", ","), "ANY(x,")

    def test_todo_trecho_de_funcao_e_do_teclado_e_aceito_pelo_nucleo(self):
        df = pd.DataFrame({f"Q{i}": [1., 2., 3.] for i in range(1, 6)})
        df["x"], df["y"] = [1., 2., 3.], [4., 5., 6.]
        df["a"], df["b"] = ["12", "7", "x"], ["s", "t", "u"]
        banco = td.Banco(df=df, nome="x.sav")
        for grupo, funcs in dlg.FUNCOES.items():
            for f in funcs:
                modelo = f.replace("(x", "(a") if grupo == "Texto" else f      # funções de texto pedem variável de texto
                try:
                    td.compute(banco, "r", modelo)
                except ErroTratamento as exc:
                    self.fail(f"{grupo}: {f} -> {exc}")
        for linha in dlg.TECLADO:
            for _, trecho in linha:
                self.assertTrue(trecho)


class TestRegrasDeTabela(unittest.TestCase):
    def test_tipos_e_else_por_ultimo(self):
        linhas = [
            {"Tipo": "Todos os outros valores", "Novo valor": "3"},
            {"Tipo": "Do menor valor até …", "Até": "29", "Novo valor": "1"},
            {"Tipo": "Intervalo (de … até …)", "De": "30", "Até": "49", "Novo valor": "2"},
            {"Tipo": "De … até o maior valor", "De": "90", "Novo valor": "SYSMIS"},
            {"Tipo": "Ausente do sistema", "Novo valor": "0"},
            {"Tipo": "Ausente (sistema ou usuário)", "Novo valor": "9"},
            {"Tipo": "Valor", "De": "5", "Novo valor": "COPY"},
            {"Tipo": "Valor", "De": "", "Até": "", "Novo valor": ""},      # linha em branco: ignorada
        ]
        self.assertEqual(dlg.regras_de_tabela(linhas).split("\n"), [
            "LO THRU 29 = 1", "30 THRU 49 = 2", "90 THRU HI = SYSMIS", "SYSMIS = 0", "MISSING = 9", "5 = COPY",
            "ELSE = 3"])

    def test_erros_claros(self):
        with self.assertRaises(ErroTratamento) as e:
            dlg.regras_de_tabela([{"Tipo": "Valor", "De": "", "Novo valor": "1"}])
        self.assertIn("Linha 1", str(e.exception))
        with self.assertRaises(ErroTratamento):
            dlg.regras_de_tabela([{"Tipo": "Intervalo (de … até …)", "De": "1", "Novo valor": "1"}])
        with self.assertRaises(ErroTratamento):
            dlg.regras_de_tabela([{"Tipo": "Valor", "De": "1", "Novo valor": ""}])
        with self.assertRaises(ErroTratamento):
            dlg.regras_de_tabela([{"Tipo": "Todos os outros valores", "Novo valor": ""}])
        with self.assertRaises(ErroTratamento):
            dlg.regras_de_tabela([{"Tipo": "Valor", "De": "", "Novo valor": ""}])      # nenhuma regra
        with self.assertRaises(ErroTratamento):
            dlg.regras_de_tabela([{"Tipo": "Inventado", "De": "1", "Novo valor": "1"}])

    def test_celulas_vazias_nan_e_numeros_crus(self):
        linhas = [{"Tipo": "Valor", "De": "1", "Até": None, "Novo valor": "10"},
                  {"Tipo": "Valor", "De": "", "Até": "", "Novo valor": ""},                 # em branco: ignorada
                  {"Tipo": "Valor", "De": "2", "Até": np.nan, "Novo valor": 20}]            # NaN e número cru
        self.assertEqual(dlg.regras_de_tabela(linhas), "1 = 10\n2 = 20")

    def test_resultado_roda_no_recode_do_nucleo(self):
        b = _banco()
        regras = dlg.regras_de_tabela([
            {"Tipo": "Ausente (sistema ou usuário)", "Novo valor": "SYSMIS"},
            {"Tipo": "Valor", "De": "1", "Novo valor": "10"},
            {"Tipo": "Intervalo (de … até …)", "De": "2", "Até": "3", "Novo valor": "20"},
            {"Tipo": "Todos os outros valores", "Novo valor": "COPY"}])
        novo, _ = td.aplicar(b, "recode", origens=["Q1"], regras=regras, destinos=["R"])
        self.assertEqual(novo.df["R"].tolist()[:3], [10., 20., 20.])
        self.assertTrue(np.isnan(novo.df["R"].iloc[4]))


class TestMissing(unittest.TestCase):
    def test_especificacao_missing(self):
        self.assertEqual(dlg.especificacao_missing("nenhum"), "")
        self.assertEqual(dlg.especificacao_missing("discretos", ["98", "99", ""]), "98, 99")
        self.assertEqual(dlg.especificacao_missing("faixa", faixa_de="", faixa_ate="0", faixa_extra="99"),
                         "LO THRU 0, 99")
        self.assertEqual(dlg.especificacao_missing("faixa", faixa_de="90", faixa_ate="99"), "90 THRU 99")
        self.assertEqual(dlg.especificacao_missing("discretos", ["NS", "NR"], texto=True), "'NS', 'NR'")
        self.assertEqual(dlg.especificacao_missing("discretos", ["'NS'"], texto=True), "'NS'")
        for args in (("discretos", ["", "", ""]), ("discretos", ["1", "2", "3", "4"]), ("faixa",), ("xx",)):
            with self.assertRaises(ErroTratamento):
                dlg.especificacao_missing(*args)

    def test_especificacao_aceita_pelo_nucleo(self):
        b = _banco()
        for modo, kw in (("discretos", dict(discretos=["98", "99"])),
                         ("faixa", dict(faixa_ate="0", faixa_extra="99"))):
            spec = dlg.especificacao_missing(modo, **kw)
            novo, _ = td.aplicar(b, "definir_missing", nomes=["Q1"], especificacao=spec)
            self.assertTrue(novo.missing["Q1"])
        spec = dlg.especificacao_missing("discretos", ["NS"], texto=True)
        novo, _ = td.aplicar(b, "definir_missing", nomes=["Cidade"], especificacao=spec)
        self.assertTrue(novo.missing["Cidade"])


class TestGradesHtml(unittest.TestCase):
    def _banco_rico(self):
        df = pd.DataFrame({"Q1": [1., 2., np.nan], "Preco": [1.5, 2., 3.], "Nome": ["a<b", "", None]})
        return td.Banco(df=df, rotulos_var={"Q1": "Satisfação <geral>"}, rotulos_valor={"Q1": {1.0: "Baixa", 2.0: "Alta"}},
                        missing={"Q1": [{"lo": 99.0, "hi": 99.0}]}, medida={"Q1": "ordinal"}, nome="x.sav")

    def test_data_view_formata_como_o_spss(self):
        h = dlg.tabela_html_dados(self._banco_rico(), 10)
        self.assertIn('<th class="td-rn">1</th>', h)                    # nº do caso
        self.assertIn('<td class="td-n">1</td>', h)                     # inteiro sem casas
        self.assertIn('<td class="td-n">.</td>', h)                     # ausente do sistema = "."
        self.assertIn('<td class="td-n">1.50</td>', h)                  # coluna com decimais: 2 casas
        self.assertIn("<td>a&lt;b</td>", h)                             # texto escapado
        self.assertIn('title="Satisfação &lt;geral&gt;"', h)            # rótulo da variável na dica
        self.assertNotIn("<b>", h)

    def test_data_view_com_rotulos_dos_valores_e_limite_de_linhas(self):
        h = dlg.tabela_html_dados(self._banco_rico(), 10, com_rotulos=True)
        self.assertIn('<td class="td-n">Baixa</td>', h)
        self.assertIn('<td class="td-n">.</td>', h)                     # ausente continua "."
        h2 = dlg.tabela_html_dados(self._banco_rico(), 2)
        self.assertEqual(h2.count("<tr>"), 3)                           # cabeçalho + 2 linhas

    def test_variable_view_tem_as_colunas_reais(self):
        h = dlg.tabela_html_variaveis(self._banco_rico())
        for col in ("Nome", "Tipo", "Rótulo", "Valores", "Ausentes", "Medida", "Casos válidos"):
            self.assertIn(f"<th>{col}</th>", h)
        self.assertIn("Satisfação &lt;geral&gt;", h)
        self.assertIn("Ordinal", h)
        self.assertIn("String", h)
        self.assertIn("1=Baixa", h)
        self.assertEqual(h.count("<tr>"), 4)                            # cabeçalho + 3 variáveis

    def test_banco_vazio_nao_quebra(self):
        vazio = td.Banco(df=pd.DataFrame({"a": pd.Series([], dtype=float)}), nome="v.sav")
        self.assertIn("<table>", dlg.tabela_html_dados(vazio, 10))
        self.assertIn("<table>", dlg.tabela_html_variaveis(vazio))


class TestAcoesDaVariavel(unittest.TestCase):
    def test_acoes_da_variavel_apontam_para_janelas(self):
        for rotulo, ident in dlg.ACOES_DA_VARIAVEL:
            self.assertIn(ident, dlg.OPERACAO_DA_JANELA)


if __name__ == "__main__":
    unittest.main()
