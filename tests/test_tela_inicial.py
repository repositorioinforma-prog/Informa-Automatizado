"""Tela inicial: todo card com destino precisa abrir uma opção que existe no menu."""
import ast
import os
import sys
import types
import unittest

if "streamlit" not in sys.modules:
    _st = types.ModuleType("streamlit")
    _st.session_state = {}
    _comp = types.ModuleType("streamlit.components")
    _v1 = types.ModuleType("streamlit.components.v1")
    _v1.declare_component = lambda *a, **k: (lambda **kw: None)
    _comp.v1 = _v1
    _st.components = _comp
    sys.modules.update({"streamlit": _st, "streamlit.components": _comp, "streamlit.components.v1": _v1})

from analises import tela_inicial as ti  # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(__file__))


def _menu_do_app():
    arvore = ast.parse(open(os.path.join(RAIZ, "app.py"), encoding="utf-8").read())
    for no in arvore.body:
        if isinstance(no, ast.Assign) and getattr(no.targets[0], "id", "") == "MENU":
            return [ti.INICIO if isinstance(e, ast.Name) else e.value for e in no.value.elts]
    raise AssertionError("MENU não encontrado em app.py")


class TestTelaInicial(unittest.TestCase):
    def test_destinos_existem_no_menu(self):
        menu = _menu_do_app()
        self.assertEqual(menu[0], ti.INICIO)
        for sec in ti.SECOES:
            for card in sec["cards"]:
                if card["alvo"]:
                    self.assertIn(card["alvo"], menu, f"card {card['titulo']!r} aponta para opção inexistente")

    def test_toda_ferramenta_do_menu_tem_card(self):
        """Sem seletor lateral, uma ferramenta sem card ficaria inacessível."""
        alvos = {c["alvo"] for sec in ti.SECOES for c in sec["cards"] if c["alvo"]}
        faltando = [m for m in _menu_do_app() if m != ti.INICIO and m not in alvos]
        self.assertEqual(faltando, [], f"ferramentas sem card na tela inicial: {faltando}")

    def test_cards_tem_finalidade_e_cores_validas(self):
        """O painel «finalidade» da central lê `desc`; a cor precisa existir na paleta do componente."""
        cores = {"azul", "roxo", "verde", "laranja", "violeta", "ciano"}
        html = open(os.path.join(RAIZ, "componentes", "tela_inicial", "index.html"), encoding="utf-8").read()
        for sec in ti.SECOES:
            self.assertIn(sec["cor"], cores)
            for card in sec["cards"]:
                self.assertTrue(card.get("desc"), f"card {card['titulo']!r} sem descrição (finalidade)")
                self.assertIn(f"{card['icone']}:", html.replace('"aa"', "aa:"), f"ícone {card['icone']!r} inexistente") \
                    if card["icone"] != "aa" else None
                if card.get("legado"):
                    self.assertTrue(card["alvo"])

    def test_nomes_unicos_na_busca(self):
        titulos = [c["titulo"] for sec in ti.SECOES for c in sec["cards"]]
        self.assertEqual(len(titulos), len(set(titulos)))

    def test_voltar_ao_inicio(self):
        st = sys.modules["streamlit"]
        st.session_state[ti.CHAVE_MENU] = "Ponderação"
        ti.voltar_ao_inicio()
        self.assertEqual(st.session_state[ti.CHAVE_MENU], ti.INICIO)
        st.session_state.clear()

    def test_clique_troca_o_menu_uma_vez(self):
        st = sys.modules["streamlit"]
        menu = _menu_do_app()
        st.session_state.clear()
        st.session_state["inicio_cards"] = {"alvo": "Ponderação", "nonce": "1"}
        ti.aplicar_clique_do_inicio(menu)
        self.assertEqual(st.session_state[ti.CHAVE_MENU], "Ponderação")
        st.session_state[ti.CHAVE_MENU] = ti.INICIO   # usuário volta ao Início
        ti.aplicar_clique_do_inicio(menu)              # o mesmo clique não age de novo
        self.assertEqual(st.session_state[ti.CHAVE_MENU], ti.INICIO)
        st.session_state["inicio_cards"] = {"alvo": "Não existe", "nonce": "2"}
        ti.aplicar_clique_do_inicio(menu)
        self.assertEqual(st.session_state[ti.CHAVE_MENU], ti.INICIO)
        st.session_state.clear()

    def test_ferramenta_guardada_na_url_e_restaurada(self):
        st = sys.modules["streamlit"]
        menu = _menu_do_app()
        st.session_state.clear()
        st.query_params = {}
        ti.guardar_ferramenta("Total Automático")
        self.assertEqual(st.query_params["ferramenta"], "Total Automático")
        # estado perdido (rerun/F5): a ferramenta volta da URL, não cai no Início
        st.session_state.clear()
        ti.restaurar_ferramenta(menu)
        self.assertEqual(st.session_state[ti.CHAVE_MENU], "Total Automático")
        # valor inválido na URL é ignorado
        st.session_state.clear()
        st.query_params = {"ferramenta": "Não existe"}
        ti.restaurar_ferramenta(menu)
        self.assertNotIn(ti.CHAVE_MENU, st.session_state)
        # voltar ao Início limpa o parâmetro
        st.query_params = {"ferramenta": "Ponderação"}
        ti.guardar_ferramenta(ti.INICIO)
        self.assertNotIn("ferramenta", st.query_params)
        st.session_state.clear()

    def test_reiniciar_do_zero_nao_troca_de_ferramenta(self):
        st = sys.modules["streamlit"]
        st.rerun = lambda: None
        for mod, pre, nome in (("analises.total_automatico", "ta_", "Total Automático"),
                               ("analises.relatorio_automatizado", "ra_", "Relatório Automatizado")):
            try:
                m = __import__(mod, fromlist=["_reiniciar"])
            except Exception as exc:  # dependência pesada ausente no ambiente de teste
                self.skipTest(f"{mod}: {exc}")
            st.session_state.clear()
            st.session_state[ti.CHAVE_MENU] = nome
            st.session_state[pre + "step"] = "x"
            st.session_state[pre + "log"] = ["a"]
            m._reiniciar()
            self.assertEqual(st.session_state[ti.CHAVE_MENU], nome)
            self.assertNotIn(pre + "step", st.session_state)
            self.assertNotIn(pre + "log", st.session_state)
        st.session_state.clear()


if __name__ == "__main__":
    unittest.main()
