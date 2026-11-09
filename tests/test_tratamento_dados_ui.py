"""Tela de Tratamento de Dados (estilo SPSS): fiação entre menus, janelas e núcleo, sem precisar do Streamlit real.

Usa um `streamlit` falso (só durante o teste) que devolve valores controlados, "clica" botões pelo rótulo
e registra mensagens, janelas abertas e downloads. Não valida o visual — só o comportamento.
"""
import ast
import importlib.util
import os
import sys
import types
import unittest
from unittest import mock

import numpy as np
import pandas as pd

from core import tratamento_dados_dialogos as dlg
from core import tratamento_dados_math as td

RAIZ = os.path.dirname(os.path.dirname(__file__))
CAMINHO_TELA = os.path.join(RAIZ, "analises", "tratamento_dados.py")


class _Rerun(Exception):
    pass


class _Ctx:
    """Contexto (coluna/aba/popover/expander...) que delega os widgets ao streamlit falso."""
    def __init__(self, fake):
        self._f = fake

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def __getattr__(self, nome):
        return getattr(self._f, nome)


class FakeSt(types.ModuleType):
    def __init__(self):
        super().__init__("streamlit")
        self.session_state = {}
        self.o = {}               # sobrescritas por rótulo do widget
        self.enviar = set()       # rótulos de botões "clicados" nesta execução
        self.msgs = {"error": [], "warning": [], "success": [], "info": []}
        self.downloads = []
        self.desabilitados = set()
        self.desabilitados_chaves = set()
        self.janelas = []         # (título, largura) das janelas abertas
        self.codigos = []
        self.rerun_escopos = []
        self.column_config = types.SimpleNamespace(
            Column=lambda *a, **k: ("col", a, k), TextColumn=lambda *a, **k: ("text", a, k),
            SelectboxColumn=lambda *a, **k: ("select", a, k))

    # entrada -----------------------------------------------------------------------------
    def _valor(self, label, padrao, key):
        if label in self.o:
            return self.o[label]
        if key is not None and key in self.session_state:
            return self.session_state[key]
        return padrao

    def text_input(self, label, value="", key=None, **k):
        return self._valor(label, value, key)

    text_area = text_input

    def selectbox(self, label, options, key=None, **k):
        return self._valor(label, options[0] if len(options) else None, key)

    def multiselect(self, label, options, default=None, key=None, **k):
        return self._valor(label, list(default or []), key)

    def radio(self, label, options, key=None, **k):
        return self._valor(label, options[0], key)

    def segmented_control(self, label, options, default=None, key=None, **k):
        return self._valor(label, default or options[0], key)

    def checkbox(self, label, value=False, key=None, **k):
        return self._valor(label, value, key)

    def number_input(self, label, min_value=None, max_value=None, value=0, key=None, **k):
        return self._valor(label, value, key)

    def file_uploader(self, label, **k):
        return self.o.get(label)

    def data_editor(self, df, **k):
        return self.o.get("editor", df)

    def button(self, label, on_click=None, args=(), disabled=False, **k):
        if disabled:
            self.desabilitados.add(label)
            self.desabilitados_chaves.add(k.get("key"))
            return False
        clicado = label in self.enviar or (k.get("key") in self.enviar)
        if clicado and on_click:
            on_click(*args)
        return clicado

    def download_button(self, label, data=None, **k):
        self.downloads.append((label, data, k))

    # saída -------------------------------------------------------------------------------
    def _msg(tipo):
        def f(self, texto, *a, **k):
            self.msgs[tipo].append(str(texto))
        return f

    error, warning, success, info = _msg("error"), _msg("warning"), _msg("success"), _msg("info")

    def code(self, texto, *a, **k):
        self.codigos.append(texto)

    def dialog(self, titulo, width="small", **k):
        def decorador(fn):
            def aberta(*a, **kw):
                self.janelas.append((titulo, width))
                return fn(*a, **kw)
            return aberta
        return decorador

    def __getattr__(self, nome):
        if nome.startswith("__"):
            raise AttributeError(nome)
        if nome in ("form", "expander", "spinner", "container", "popover"):
            return lambda *a, **k: _Ctx(self)
        return lambda *a, **k: None       # markdown, caption, header, dataframe, divider...

    def columns(self, spec, **k):
        return [_Ctx(self) for _ in range(spec if isinstance(spec, int) else len(spec))]

    def tabs(self, nomes):
        return [_Ctx(self) for _ in nomes]

    def rerun(self, scope=None):
        self.rerun_escopos.append(scope)
        raise _Rerun()


def _carregar_tela(fake):
    with mock.patch.dict(sys.modules, {"streamlit": fake}):
        spec = importlib.util.spec_from_file_location("tratamento_dados_ui_sob_teste", CAMINHO_TELA)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
    return mod


def _banco_exemplo():
    df = pd.DataFrame({
        "ID": [1., 2., 3., 4., 5., 6.],
        "Q1": [1., 2., 3., 99., 1., np.nan],
        "Q2": [5., 4., np.nan, 2., 1., 3.],
        "Idade": [20., 35., 40., 52., 61., 28.],
        "Cidade": ["Rio", "Niterói", "Rio", "", "Rio", "Niterói"],
    })
    return td.Banco(df=df, rotulos_var={"Q1": "Satisfação"}, rotulos_valor={"Q1": {1.0: "Baixa", 2.0: "Média", 3.0: "Alta"}},
                    missing={"Q1": [{"lo": 99.0, "hi": 99.0}]},
                    medida={"ID": "scale", "Q1": "ordinal", "Q2": "scale", "Idade": "scale", "Cidade": "nominal"},
                    nome="exemplo.sav")


class _Upload:
    def __init__(self, nome, conteudo):
        self.name, self._c = nome, conteudo
        self.size = len(conteudo)

    def getvalue(self):
        return self._c


class TestFiacaoEstatica(unittest.TestCase):
    def setUp(self):
        self.fonte = open(CAMINHO_TELA, encoding="utf-8").read()
        self.mod = _carregar_tela(FakeSt())

    def test_cada_comando_das_faixas_tem_janela_e_vice_versa(self):
        ids = set(dlg.ids_dos_comandos())
        self.assertEqual(ids, set(self.mod.JANELAS))

    def test_rodape_so_usa_operacoes_do_nucleo_e_todas_tem_janela(self):
        arvore = ast.parse(self.fonte)
        ops = {n.args[0].value for n in ast.walk(arvore)
               if isinstance(n, ast.Call) and getattr(n.func, "id", "") == "_rodape"
               and n.args and isinstance(n.args[0], ast.Constant)}
        self.assertTrue(ops <= set(td.OPERACOES), ops - set(td.OPERACOES))
        self.assertEqual(ops, set(td.OPERACOES), "toda operação do núcleo deve ter janela com OK")

    def test_operacao_declarada_no_mapa_e_a_usada_na_janela(self):
        # cada janela do menu chama _rodape com a operação que o mapa diz (ex.: recode_* -> recode)
        arvore = ast.parse(self.fonte)
        por_funcao = {}
        for f in [n for n in arvore.body if isinstance(n, ast.FunctionDef)]:
            usos = {c.args[0].value for c in ast.walk(f) if isinstance(c, ast.Call)
                    and getattr(c.func, "id", "") == "_rodape" and c.args and isinstance(c.args[0], ast.Constant)}
            por_funcao[f.name] = usos
        for ident, op in dlg.OPERACAO_DA_JANELA.items():
            if op is None or ident.startswith("recode"):
                continue
            funcao = self.mod.JANELAS[ident][2].__name__
            self.assertIn(op, por_funcao[funcao], f"{ident}: a janela {funcao} não usa a operação {op}")

    def test_app_e_tela_inicial(self):
        app = open(os.path.join(RAIZ, "app.py"), encoding="utf-8").read()
        self.assertIn('if analise == "Tratamento de Dados":', app)
        self.assertIn("modulo_tratamento_dados", app)


class Base(unittest.TestCase):
    def setUp(self):
        self.st = FakeSt()
        self.mod = _carregar_tela(self.st)
        self.b = _banco_exemplo()
        self.st.session_state.update({
            self.mod.K_ORIG: self.b, self.mod.K_ATUAL: self.b, self.mod.K_PASSOS: [], self.mod.K_VER: 0})

    def atual(self):
        return self.st.session_state[self.mod.K_ATUAL]

    def passos(self):
        return self.st.session_state[self.mod.K_PASSOS]

    def janela(self, ident, clicar=("OK",), **overrides):
        """Abre a janela `ident`; devolve True se terminou em st.rerun() (aplicou/fechou)."""
        self.st.o = overrides
        self.st.enviar = set(clicar)
        try:
            self.mod.JANELAS[ident][2](self.atual())
        except _Rerun:
            return True
        return False


class TestJanelasDeVariaveis(Base):
    def test_calcular_com_condicao_e_historico(self):
        ok = self.janela("compute", **{
            "Variável de destino (nova ou existente)": "Dobro", "Expressão numérica ou de texto": "Idade * 2",
            "Condição (opcional — equivale ao IF)": "Idade > 30"})
        self.assertTrue(ok, self.st.msgs)
        self.assertEqual(self.atual().df["Dobro"].dropna().tolist(), [70., 80., 104., 122.])
        self.assertEqual(len(self.passos()), 1)
        self.assertEqual(self.st.session_state[self.mod.K_VER], 1)

    def test_calcular_erro_amigavel_nao_muda_o_banco(self):
        ok = self.janela("compute", **{"Variável de destino (nova ou existente)": "X",
                                       "Expressão numérica ou de texto": "Inexistente + 1"})
        self.assertFalse(ok)
        self.assertTrue(any("Inexistente" in m for m in self.st.msgs["error"]))
        self.assertEqual(self.passos(), [])

    def test_botoes_de_inserir_montam_a_expressao(self):
        chave = self.mod._k("expr")
        self.janela("compute", clicar=("Inserir função ➜",), **{"Função": "MEAN(Q1 TO Q5)"})
        self.assertEqual(self.st.session_state[chave], "MEAN(Q1 TO Q5)")
        self.janela("compute", clicar=("×",))                         # um clique por execução, como no Streamlit
        self.janela("compute", clicar=("Inserir variável ➜",), **{"Variáveis do banco": "Idade"})
        self.assertEqual(self.st.session_state[chave], "MEAN(Q1 TO Q5) * Idade")
        # a expressão montada pelos botões é a que o OK usa
        self.st.session_state[chave] = "Idade * 3"
        self.assertTrue(self.janela("compute", **{"Variável de destino (nova ou existente)": "Tri"}))
        self.assertEqual(self.atual().df["Tri"].iloc[0], 60.)

    def test_ver_syntax_nao_altera_o_banco(self):
        ok = self.janela("compute", clicar=("Ver syntax",), **{
            "Variável de destino (nova ou existente)": "Dobro", "Expressão numérica ou de texto": "Idade * 2"})
        self.assertFalse(ok)
        self.assertEqual(self.passos(), [])
        self.assertNotIn("Dobro", self.atual().df.columns)
        self.assertTrue(any("COMPUTE" in c and "Dobro" in c for c in self.st.codigos), self.st.codigos)

    def test_reiniciar_recria_os_widgets_e_cancelar_fecha(self):
        self.assertTrue(self.janela("compute", clicar=("Reiniciar",)))
        self.assertEqual(self.st.session_state[self.mod.K_RESET], 1)
        self.assertEqual(self.st.rerun_escopos[-1], "fragment")
        self.assertNotEqual(self.mod._k("x"), "td_d_x_0_0")
        self.assertTrue(self.janela("compute", clicar=("Cancelar",)))
        self.assertEqual(self.st.rerun_escopos[-1], None)
        self.assertEqual(self.passos(), [])

    def test_recodificar_em_outra_variavel_pela_tabela(self):
        ok = self.janela("recode_outras", **{
            "Variável(is) de origem": ["Idade"],
            "Valor antigo 1": "Do menor valor até …", "Até 1": "29", "Novo valor 1": "1",
            "Valor antigo 2": "De … até o maior valor", "De 2": "30", "Novo valor 2": "2",
            "Nome(s) da(s) variável(is) de saída (um por origem, separados por espaço)": "Faixa",
            "Rótulos dos valores novos (opcional)": "1 = Jovem\n2 = Adulto"})
        self.assertTrue(ok, self.st.msgs)
        r = self.atual()
        self.assertEqual(r.df["Faixa"].tolist(), [1., 2., 2., 2., 2., 1.])
        self.assertEqual(r.rotulos_valor["Faixa"], {1.0: "Jovem", 2.0: "Adulto"})
        self.assertEqual(r.df["Idade"].tolist(), [20., 35., 40., 52., 61., 28.])      # a original fica intacta

    def test_recodificar_em_outra_variavel_pelo_texto(self):
        ok = self.janela("recode_outras", **{
            "Variável(is) de origem": ["Idade"], "Como informar as regras": "Texto (syntax)",
            "Regras (uma por linha)": "LO THRU 29 = 1\nELSE = COPY",
            "Nome(s) da(s) variável(is) de saída (um por origem, separados por espaço)": "F2"})
        self.assertTrue(ok, self.st.msgs)
        self.assertEqual(self.atual().df["F2"].tolist(), [1., 35., 40., 52., 61., 1.])

    def test_recodificar_nas_mesmas_variaveis(self):
        ok = self.janela("recode_mesmas", **{"Variável(is) de origem": ["Q2"], "De 1": "5", "Novo valor 1": "50"})
        self.assertTrue(ok, self.st.msgs)
        self.assertEqual(self.atual().df["Q2"].dropna().tolist(), [50., 4., 2., 1., 3.])
        self.assertEqual(self.passos()[0].operacao, "recode")
        self.assertNotIn("INTO", self.passos()[0].sintaxe)

    def test_recodificar_valida_origem_e_saida(self):
        regra = {"De 1": "5", "Novo valor 1": "50"}
        self.assertFalse(self.janela("recode_outras", **regra))
        self.assertTrue(any("origem" in m for m in self.st.msgs["error"]))
        self.assertFalse(self.janela("recode_outras", **regra, **{"Variável(is) de origem": ["Q2"]}))
        self.assertTrue(any("saída" in m for m in self.st.msgs["error"]))
        self.assertEqual(self.passos(), [])

    def test_adicionar_linha_nas_regras_e_nos_resumos(self):
        self.assertTrue(self.janela("recode_outras", clicar=(self.mod._k("n_regras_mais"),)))
        self.assertEqual(self.st.session_state[self.mod._k("n_regras")], 5)
        self.assertEqual(self.st.rerun_escopos[-1], "fragment")
        # a 5ª linha passa a existir e é lida pelo OK
        ok = self.janela("recode_outras", **{
            "Variável(is) de origem": ["Idade"], "Valor antigo 5": "Todos os outros valores", "Novo valor 5": "9",
            "Nome(s) da(s) variável(is) de saída (um por origem, separados por espaço)": "F"})
        self.assertTrue(ok, self.st.msgs)
        self.assertEqual(set(self.atual().df["F"]), {9.0})
        self.assertTrue(self.janela("agregar", clicar=(self.mod._k("n_agr_mais"),)))
        self.assertEqual(self.st.session_state[self.mod._k("n_agr")], 4)

    def test_renomear_e_excluir(self):
        self.assertTrue(self.janela("renomear_variavel", **{"Variável": "Q2", "Novo nome": "Q2b"}))
        self.assertIn("Q2b", self.atual().df.columns)
        self.assertTrue(self.janela("excluir_variaveis", **{"Variáveis a excluir": ["Q2b"]}))
        self.assertNotIn("Q2b", self.atual().df.columns)


class TestJanelasDeCasos(Base):
    def test_selecionar_manter_e_excluir(self):
        self.assertTrue(self.janela("selecionar_casos", **{"Condição": "Idade >= 35"}))
        self.assertEqual(len(self.atual().df), 4)
        self.assertTrue(self.janela("selecionar_casos", **{"Condição": "Idade > 50",
                                                           "Casos": "Excluir os casos que satisfazem a condição"}))
        self.assertEqual(self.atual().df["Idade"].tolist(), [35., 40.])      # saíram os maiores de 50

    def test_ordenar(self):
        self.assertTrue(self.janela("ordenar_casos", **{"1ª variável": "Idade", "Ordem": "D"}))
        self.assertEqual(self.atual().df["Idade"].tolist(), [61., 52., 40., 35., 28., 20.])

    def test_remover_duplicados(self):
        self.assertTrue(self.janela("remover_duplicados", **{
            "Variáveis que identificam o caso (ex.: telefone, ID)": ["Cidade"]}))
        self.assertEqual(len(self.atual().df), 3)

    def test_agregar_pela_tabela(self):
        ok = self.janela("agregar", **{"Variáveis de quebra (BREAK) — vazio = banco inteiro": ["Cidade"],
                                       "Nova variável 1": "Media", "Função 1": "MEAN", "Variável 1": "Idade"})
        self.assertTrue(ok, self.st.msgs)
        self.assertIn("Media", self.atual().df.columns)

    def test_segundo_banco_adicionar_casos_e_variaveis(self):
        dados, _ = td.gravar_sav(td.Banco(df=pd.DataFrame({"ID": [2., 4.], "Renda": [10., 20.]}), nome="renda.sav"))
        up = _Upload("renda.sav", dados)
        self.assertTrue(self.janela("adicionar_variaveis", **{
            "Segundo banco (.sav)": up, "Chave(s) de ligação (variáveis presentes nos dois bancos)": ["ID"]}))
        self.assertEqual(self.atual().df["Renda"].dropna().tolist(), [10., 20.])
        with self.assertRaises(_Rerun):
            self.mod._desfazer(1)                       # o histórico guarda o segundo banco: refaz do original
        self.assertNotIn("Renda", self.atual().df.columns)
        self.assertTrue(self.janela("adicionar_casos", **{"Segundo banco (.sav)": up}))
        self.assertEqual(len(self.atual().df), 8)

    def test_janela_de_segundo_banco_sem_arquivo_so_oferece_fechar(self):
        self.assertTrue(self.janela("adicionar_casos", clicar=("Fechar",)))
        self.assertEqual(self.passos(), [])


class TestJanelasDeRotulos(Base):
    def test_rotulo_da_variavel(self):
        self.assertTrue(self.janela("rotular_variavel", **{
            "Variável": "Q2", "Rótulo (deixe em branco para remover)": "Dois"}))
        self.assertEqual(self.atual().rotulos_var["Q2"], "Dois")

    def test_rotulos_dos_valores_pela_grade(self):
        self.assertTrue(self.janela("rotular_valores", **{
            "Variável de referência (mostra os rótulos atuais)": "Q2", "Aplicar às variáveis": ["Q2"],
            "Rótulos (um por linha: código = rótulo)": "1 = Um\n2 = Dois"}))
        self.assertEqual(self.atual().rotulos_valor["Q2"], {1.0: "Um", 2.0: "Dois"})
        self.assertTrue(self.janela("rotular_valores", **{
            "Variável de referência (mostra os rótulos atuais)": "Q2", "Aplicar às variáveis": ["Q2"],
            "Modo": "remover"}))
        self.assertFalse(self.atual().rotulos_valor.get("Q2"))

    def test_missing_nos_tres_modos(self):
        base = {"Variáveis (todas do mesmo tipo)": ["Q2"]}
        self.assertTrue(self.janela("definir_missing", **base, **{"Valores ausentes": "discretos", "Valor 1": "5"}))
        self.assertEqual(self.atual().missing["Q2"], [{"lo": 5.0, "hi": 5.0}])
        self.assertTrue(self.janela("definir_missing", **base, **{
            "Valores ausentes": "faixa", "Máximo (vazio = maior valor)": "0", "Valor discreto (opcional)": "99"}))
        self.assertEqual(len(self.atual().missing["Q2"]), 2)
        self.assertTrue(self.janela("definir_missing", **base, **{"Valores ausentes": "nenhum"}))
        self.assertFalse(self.atual().missing.get("Q2"))

    def test_missing_de_variavel_de_texto_ganha_aspas(self):
        self.assertTrue(self.janela("definir_missing", **{
            "Variáveis (todas do mesmo tipo)": ["Cidade"], "Valores ausentes": "discretos", "Valor 1": "NS"}))
        self.assertIn("'NS'", self.passos()[-1].sintaxe)

    def test_missing_sem_valor_mostra_erro(self):
        self.assertFalse(self.janela("definir_missing", **{
            "Variáveis (todas do mesmo tipo)": ["Q2"], "Valores ausentes": "discretos"}))
        self.assertTrue(self.st.msgs["error"])

    def test_nivel_de_medida(self):
        self.assertTrue(self.janela("definir_medida", **{"Variáveis": ["Q2"], "Nível de medida": "ordinal"}))
        self.assertEqual(self.atual().medida["Q2"], "ordinal")


class TestArquivoEAjuda(Base):
    def test_salvar_sav_gera_e_oferece_download(self):
        self.janela("salvar_sav", clicar=())
        self.assertNotIn("Baixar .sav", [d[0] for d in self.st.downloads])
        self.janela("salvar_sav", clicar=("Gerar arquivo .sav",))
        self.assertIn("Baixar .sav", [d[0] for d in self.st.downloads])
        self.assertEqual(self.st.msgs["error"], [])

    def test_exportar_excel_e_csv(self):
        self.janela("exportar", clicar=("Gerar Excel",))
        rotulos = [d[0] for d in self.st.downloads]
        self.assertIn("Baixar Excel (.xlsx)", rotulos)
        self.assertIn("Baixar CSV (;)", rotulos)
        self.assertEqual(self.st.msgs["error"], [])

    def test_salvar_syntax_precisa_de_passos(self):
        self.janela("salvar_sintaxe", clicar=())
        self.assertTrue(self.st.msgs["info"])
        self.assertNotIn("Baixar syntax (.sps)", [d[0] for d in self.st.downloads])
        self.janela("compute", **{"Variável de destino (nova ou existente)": "A", "Expressão numérica ou de texto": "1"})
        self.st.downloads.clear()
        self.janela("salvar_sintaxe", clicar=())
        self.assertIn("Baixar syntax (.sps)", [d[0] for d in self.st.downloads])

    def test_ajudas_abrem_sem_erro(self):
        for ident in ("ajuda_expressoes", "ajuda_geral"):
            self.assertFalse(self.janela(ident, clicar=()))
            self.assertTrue(self.janela(ident, clicar=("Fechar",)))


class TestMenuEModulo(Base):
    def test_faixas_devolvem_o_comando_clicado(self):
        for ident in dlg.ids_dos_comandos():
            self.st.enviar = {f"td_c_{ident}"}
            self.assertEqual(self.mod._faixas(), ident)
        self.st.enviar = set()
        self.assertIsNone(self.mod._faixas())

    def test_todos_os_comandos_aparecem_com_icone_e_descricao(self):
        rotulos, captions = [], []
        orig_button, orig_caption = self.st.button, self.st.caption
        self.st.button = lambda label, **k: rotulos.append((label, k.get("icon"), k.get("key"))) or False
        self.st.caption = lambda texto, *a, **k: captions.append(texto)
        self.mod._faixas()
        esperados = [(c[2], f":material/{c[1]}:", f"td_c_{c[0]}") for _, g in dlg.FAIXAS for _, cs in g for c in cs]
        self.assertEqual(rotulos, esperados)
        self.assertEqual(captions, [c[3] for _, g in dlg.FAIXAS for _, cs in g for c in cs])

    def test_desfazer_fica_a_vista_e_bloqueia_sem_passos(self):
        self.st.enviar = {"td_q_desfazer"}
        self.assertIsNone(self.mod._acesso_rapido(False))
        self.assertIn("td_q_desfazer", self.st.desabilitados_chaves)
        self.assertIn("td_q_desfazer_tudo", self.st.desabilitados_chaves)
        self.assertEqual(self.mod._acesso_rapido(True), "desfazer")
        self.st.enviar = {"td_q_desfazer_tudo"}
        self.assertEqual(self.mod._acesso_rapido(True), "desfazer_tudo")

    def test_todos_os_comandos_abrem_a_janela_certa(self):
        for ident in dlg.ids_dos_comandos():
            self.st.janelas.clear()
            self.st.o, self.st.enviar = {}, set()
            self.mod._abrir_janela(ident, self.atual())
            self.assertEqual(self.st.janelas[-1][0], self.mod.JANELAS[ident][0])

    def test_modulo_completo_menu_dialogo_ok(self):
        dados, _ = td.gravar_sav(self.b)
        self.st.session_state.clear()
        self.st.o = {"Carregue o banco SPSS (.sav)": _Upload("exemplo.sav", dados)}
        self.st.enviar = set()
        self.mod.modulo_tratamento_dados()
        self.assertEqual(self.st.msgs["error"], [])
        self.assertEqual(len(self.atual().df), 6)
        # clica no comando e a janela abre
        self.st.enviar = {"td_c_compute"}
        self.mod.modulo_tratamento_dados()
        self.assertEqual(self.st.janelas[-1][0], "Calcular variável")
        # na mesma execução o OK aplica (simula a janela aberta + clique)
        self.st.o.update({"Variável de destino (nova ou existente)": "Dobro",
                          "Expressão numérica ou de texto": "Idade * 2"})
        self.st.enviar = {"td_c_compute", "OK"}
        with self.assertRaises(_Rerun):
            self.mod.modulo_tratamento_dados()
        self.assertIn("Dobro", self.atual().df.columns)
        # na execução seguinte aparece a mensagem de sucesso e o passo no histórico
        self.st.o = {"Carregue o banco SPSS (.sav)": _Upload("exemplo.sav", dados)}
        self.st.enviar = set()
        self.mod.modulo_tratamento_dados()
        self.assertTrue(any("Dobro" in m for m in self.st.msgs["success"]))
        # desfazer pelo botão da barra
        self.st.enviar = {"td_q_desfazer"}
        with self.assertRaises(_Rerun):
            self.mod.modulo_tratamento_dados()
        self.assertNotIn("Dobro", self.atual().df.columns)

    def test_variable_view_abre_a_janela_ja_na_variavel_escolhida(self):
        self.st.o = {"Variável escolhida": "Q2"}
        self.st.enviar = {"td_vv_rotular_variavel"}
        self.assertEqual(self.mod._visao_variaveis(self.atual()), "rotular_variavel")
        self.assertEqual(self.st.session_state[self.mod.K_FOCO], "Q2")
        # a janela abre com Q2 pré-selecionada e o OK grava nela
        self.st.o, self.st.enviar = {"Rótulo (deixe em branco para remover)": "Segunda"}, {"OK"}
        self.st.session_state[self.mod._k("var")] = "Q2"
        with self.assertRaises(_Rerun):
            self.mod.JANELAS["rotular_variavel"][2](self.atual())
        self.assertEqual(self.atual().rotulos_var["Q2"], "Segunda")
        self.assertNotIn(self.mod.K_FOCO, self.st.session_state)          # aplicar limpa a variável escolhida

    def test_foco_pre_seleciona_nas_janelas(self):
        self.st.session_state[self.mod.K_FOCO] = "Idade"
        self.assertEqual(self.mod._foco(self.atual()), "Idade")
        self.st.session_state[self.mod.K_FOCO] = "Inexistente"
        self.assertIsNone(self.mod._foco(self.atual()))
        self.st.session_state[self.mod.K_FOCO] = "Idade"
        grades = []
        orig = self.st.multiselect
        self.st.multiselect = lambda label, options, default=None, key=None, **k: grades.append((label, default)) or []
        self.mod.JANELAS["definir_medida"][2](self.atual())
        self.st.multiselect = orig
        self.assertIn(("Variáveis", ["Idade"]), grades)

    def test_variable_view_sem_clique_nao_pede_janela(self):
        self.assertIsNone(self.mod._visao_variaveis(self.atual()))
        self.assertNotIn(self.mod.K_FOCO, self.st.session_state)

    def test_visualizacao_de_syntax_e_dados_renderizam(self):
        self.mod._visao_dados(self.atual())
        self.mod._visao_syntax(self.b, [])
        self.assertTrue(self.st.msgs["info"])
        self.janela("compute", **{"Variável de destino (nova ou existente)": "A", "Expressão numérica ou de texto": "1"})
        self.st.enviar = {"Desfazer último passo"}
        with self.assertRaises(_Rerun):
            self.mod._visao_syntax(self.b, self.passos())
        self.assertEqual(self.passos(), [])

    def test_desfazer_refaz_a_partir_do_original(self):
        for nome, expr in (("A", "1"), ("B", "2")):
            self.janela("compute", **{"Variável de destino (nova ou existente)": nome,
                                      "Expressão numérica ou de texto": expr})
        with self.assertRaises(_Rerun):
            self.mod._desfazer(1)
        self.assertIn("A", self.atual().df.columns)
        self.assertNotIn("B", self.atual().df.columns)
        with self.assertRaises(_Rerun):
            self.mod._desfazer(None)
        self.assertEqual(list(self.atual().df.columns), list(self.b.df.columns))

    def test_sem_arquivo_mostra_aviso_e_limpa_estado(self):
        self.st.o = {}
        self.mod.modulo_tratamento_dados()
        self.assertTrue(self.st.msgs["info"])
        self.assertNotIn(self.mod.K_ATUAL, self.st.session_state)


if __name__ == "__main__":
    unittest.main()
