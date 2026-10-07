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

    def test_clique_troca_o_menu_uma_vez(self):
        st = sys.modules["streamlit"]
        menu = _menu_do_app()
        st.session_state.clear()
        st.session_state["inicio_cards"] = {"alvo": "Ponderação", "nonce": "1"}
        ti.aplicar_clique_do_inicio(menu)
        self.assertEqual(st.session_state[ti.CHAVE_MENU], "Ponderação")
        st.session_state[ti.CHAVE_MENU] = ti.INICIO   # usuário volta ao Início pelo menu
        ti.aplicar_clique_do_inicio(menu)              # o mesmo clique não age de novo
        self.assertEqual(st.session_state[ti.CHAVE_MENU], ti.INICIO)
        st.session_state["inicio_cards"] = {"alvo": "Não existe", "nonce": "2"}
        ti.aplicar_clique_do_inicio(menu)
        self.assertEqual(st.session_state[ti.CHAVE_MENU], ti.INICIO)
        st.session_state.clear()


if __name__ == "__main__":
    unittest.main()
