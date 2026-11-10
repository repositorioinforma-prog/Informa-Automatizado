"""Padronização de Títulos: catálogo dos títulos padrão por tipo de pergunta (núcleo, sem Streamlit)."""
import io
import os
import sys
import unittest

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import padronizacao_titulos_math as pt  # noqa: E402


def _t(pid, **valores):
    return pt.montar_titulo(pt.por_id(pid), valores)


class TestCatalogo(unittest.TestCase):
    def test_ids_unicos_e_grupos_validos(self):
        ids = [p.id for p in pt.PADROES]
        self.assertEqual(len(ids), len(set(ids)))
        for p in pt.PADROES:
            self.assertIn(p.grupo, pt.GRUPOS)
        self.assertEqual({p.grupo for p in pt.PADROES}, set(pt.GRUPOS))   # nenhum grupo vazio

    def test_campos_declarados_batem_com_os_modelos(self):
        for p in pt.PADROES:
            usados = set(pt._campos_do_modelo(p.modelo_titulo)) | set(pt._campos_do_modelo(p.modelo_pergunta))
            self.assertEqual(usados, set(p.nomes_campos), p.id)
            self.assertEqual(len(p.nomes_campos), len(set(p.nomes_campos)), p.id)

    def test_exemplo_monta_sem_campo_faltando(self):
        for p in pt.PADROES:
            titulo = pt.titulo_exemplo(p)
            self.assertNotIn("{", titulo)
            self.assertNotIn("‹", titulo)
            self.assertEqual(titulo, " ".join(titulo.split()))     # sem espaço duplo

    def test_todo_titulo_padrao_segue_as_formas_escolhidas(self):
        for p in pt.PADROES:
            titulo = pt.titulo_exemplo(p)
            self.assertNotRegex(titulo, r"/26\b", p.id)               # /2026, nunca /26
            self.assertNotIn("Cenário único", titulo, p.id)           # Cenário Único
            self.assertNotIn("Índice de probabilidade", titulo, p.id)
            self.assertNotIn("Rejeição de voto", titulo, p.id)
            self.assertNotIn("Rejeição aos candidatos a ", titulo, p.id)
            self.assertNotIn("Presidente/", titulo, p.id)             # Presidente da República

    def test_sufixo_do_tipo_de_resposta(self):
        self.assertTrue(_t("voto_espontanea", cargo="Governador").endswith("(Espontânea e Única)"))
        self.assertTrue(_t("rejeicao", cargo="Senador").endswith("(Estimulada e Múltipla)"))
        self.assertTrue(pt.titulo_exemplo(pt.por_id("ipv_candidato")).endswith("(Estimulada e Única)"))
        # tabelas informativas não levam sufixo
        self.assertFalse(pt.por_id("partidos_deputados").sufixo)
        self.assertEqual(_t("partidos_deputados", cargo_plural="Deputados Federais"), "Partido/Federação dos Deputados Federais")
        self.assertEqual(_t("conhece_quem_e"), "Para quem conhece o candidato em algum nível")

    def test_campos_de_lista_aceitam_outro_valor(self):
        self.assertEqual(_t("rejeicao", cargo="Vice-Governador"), "Rejeição aos candidatos para Vice-Governador (Estimulada e Múltipla)")

    def test_cargos_usados_como_opcao_existem(self):
        for p in pt.PADROES:
            for c in p.campos:
                if c.nome == "cargo":
                    self.assertIn(c.exemplo, c.opcoes)


class TestTitulosReais(unittest.TestCase):
    """Títulos dos relatórios de verdade que já estão no padrão escolhido devem sair idênticos."""

    CASOS = [
        ("voto_espontanea", dict(cargo="Deputado Estadual"), "Intenção de voto para Deputado Estadual/2026 (Espontânea e Única)"),
        ("voto_espontanea", dict(cargo="Deputado Federal"), "Intenção de voto para Deputado Federal/2026 (Espontânea e Única)"),
        ("voto_cenario_unico", dict(cargo="Presidente da República"), "Intenção de voto para Presidente da República/2026 - Cenário Único (Estimulada e Única)"),
        ("voto_cenario_unico", dict(cargo="Governador"), "Intenção de voto para Governador/2026 - Cenário Único (Estimulada e Única)"),
        ("voto_cenario_unico", dict(cargo="Deputado Federal"), "Intenção de voto para Deputado Federal/2026 - Cenário Único (Estimulada e Única)"),
        ("voto_cenario_numerado", dict(cargo="Deputado Estadual", numero="2"), "Intenção de voto para Deputado Estadual/2026 - Cenário 2 (Estimulada e Única)"),
        ("voto_transferencia", dict(cargo="Deputado Estadual", candidato="Greg Duarte"), "Intenção de voto para Deputado Estadual/2026 - Transferência de voto sem Greg Duarte (Estimulada e Única)"),
        ("voto_senador_primeiro", {}, "Intenção de voto para Senador/2026 - Primeiro Candidato - Cenário Único (Estimulada e Única)"),
        ("voto_senador_segundo", {}, "Intenção de voto para Senador/2026 - Segundo Candidato - Cenário Único (Estimulada e Única)"),
        ("voto_senador_base100", {}, "Intenção de voto para Senador/2026 - Base 100% (1º + 2º Voto) (Estimulada e Única)"),
        ("voto_segundo_turno", dict(cargo="Governador"), "Intenção de voto para Governador/2026 no 2º Turno (Estimulada e Única)"),
        ("rejeicao", dict(cargo="Presidente da República"), "Rejeição aos candidatos para Presidente da República (Estimulada e Múltipla)"),
        ("rejeicao", dict(cargo="Senador"), "Rejeição aos candidatos para Senador (Estimulada e Múltipla)"),
        ("ipv_candidato", dict(cargo="Deputado Estadual", candidato="Adilson Pires"), "Índice de Probabilidade de Voto (IPV) para Deputado Estadual/2026 - Adilson Pires (Estimulada e Única)"),
        ("ipv_candidato", dict(cargo="Senadora", candidato="Benedita da Silva"), "Índice de Probabilidade de Voto (IPV) para Senadora/2026 - Benedita da Silva (Estimulada e Única)"),
        ("ipv_apoiado_por", dict(cargo="Deputado Estadual", apoiador="Bolsonaro"), "Índice de Probabilidade de Voto (IPV) para candidato a Deputado Estadual/2026 apoiado por Bolsonaro (Estimulada e Única)"),
        ("ipv_apos_informacoes", dict(cargo="Deputado Estadual", candidato="Rogério Lins"), "Índice de Probabilidade de Voto (IPV) para Deputado Estadual/2026 - Rogério Lins, após acréscimo de informações (Estimulada e Única)"),
        ("chance_chapa", dict(chapa="Lula para Presidente e Benedita da Silva para Senadora"), "Chance de voto (Chapa) Lula para Presidente e Benedita da Silva para Senadora/2026 (Estimulada e Única)"),
        ("perfil_candidato", dict(cargo="Deputado Estadual"), "Perfil de candidato de interesse para Deputado Estadual/2026 (Estimulada e Única)"),
        ("conhecimento", dict(candidato="Rogério Lins"), "Nível de conhecimento de Rogério Lins (Estimulada e Única)"),
        ("conhecimento_apos_info", dict(candidato="Rogério Lins"), "Nível de conhecimento de Rogério Lins, após acréscimo de informações (Estimulada e Única)"),
        ("imagem_candidato", dict(candidato="Wesley Cezar", atributo="Preparo para exercer o cargo", afirmacao="x"), "Imagem de Wesley Cezar - Preparo para exercer o cargo (Estimulada e Única)"),
        ("aval_gov_estadual", dict(governador="Tarcísio de Freitas"), "Avaliação do Governo Estadual de Tarcísio de Freitas (Estimulada e Única)"),
        ("aval_gov_federal", dict(presidente="Lula"), "Avaliação do Governo Federal de Lula (Estimulada e Única)"),
        ("aval_gov_municipal", dict(cidade="Santana de Parnaíba"), "Avaliação do Governo Municipal de Santana de Parnaíba (Estimulada e Única)"),
        ("aval_gestao_prefeito", dict(prefeito="Gerson Pessoa"), "Avaliação da Gestão de Gerson Pessoa (Estimulada e Única)"),
        ("aval_governo_prefeito", dict(prefeito="Ferreti"), "Avaliação do Governo do Prefeito Ferreti (Estimulada e Única)"),
        ("aprovacao_prefeito", dict(prefeito="Ferreti"), "Aprovação do Prefeito Ferreti (Estimulada e Única)"),
        ("melhor_acao", dict(prefeito="Ferreti"), "Melhor ação do Governo do Prefeito Ferreti (Estimulada e Única)"),
        ("pior_acao", dict(prefeito="Ferreti"), "Pior ação do Governo do Prefeito Ferreti (Estimulada e Única)"),
        ("aval_camara", dict(cidade="Angra dos Reis"), "Avaliação da Câmara de Vereadores de Angra dos Reis (Estimulada e Única)"),
        ("aval_autoridade", dict(autoridade="Governador Interino", nome="Desembargador Ricardo Couto"), "Avaliação do Governador Interino - Desembargador Ricardo Couto (Estimulada e Única)"),
        ("posicionamento", {}, "Posicionamento político (Estimulada e Única)"),
        ("meios_comunicacao", dict(local="Angra dos Reis"), "Meios de comunicação para saber as notícias de Angra dos Reis (Estimulada e Múltipla)"),
        ("interesse_informar", {}, "Grau de interesse em se informar sobre as ações da Prefeitura (Estimulada e Única)"),
        ("prioridade_divulgacao", dict(cidade="Angra dos Reis"), "Prioridades de divulgação da Prefeitura de Angra dos Reis (Estimulada e Única)"),
        ("lembranca_acao", {}, "Lembrança de ação ou projeto da Prefeitura neste ano (Estimulada e Única)"),
        ("tv_aberta", {}, "Emissora de TV aberta mais assistida (Estimulada e Única)"),
        ("radio", {}, "Emissora de rádio mais ouvida (Estimulada e Única)"),
        ("jornal_diario", {}, "Jornal diário mais lido (Estimulada e Única)"),
        ("jornal_periodico", {}, "Jornal periódico mais lido (Estimulada e Única)"),
        ("redes_participacao", {}, "Participação nas redes sociais (Estimulada e Múltipla)"),
        ("redes_confianca", {}, "Confiança nas notícias veiculadas nas redes sociais (Estimulada e Única)"),
        ("importancia_aliado", dict(cargo="Deputado Federal", cidade="x"), "Percepção sobre a importância de um Deputado Federal aliado ao Prefeito (Estimulada e Única)"),
    ]

    def test_casos_reais(self):
        for pid, valores, esperado in self.CASOS:
            self.assertEqual(_t(pid, **valores), esperado, pid)

    def test_todo_padrao_tem_caso_real_ou_e_informativo(self):
        cobertos = {c[0] for c in self.CASOS}
        for p in pt.PADROES:
            if p.id not in cobertos:
                self.assertIn(p.id, ("conhece_quem_e", "partidos_deputados"), p.id)


class TestPerguntaEPreenchimento(unittest.TestCase):
    def test_pergunta_modelo(self):
        p = pt.por_id("rejeicao")
        self.assertEqual(
            pt.montar_pergunta(p, {"cargo": "Governador"}),
            "E se os candidatos a Governador fossem esses em quem você NÃO votaria de jeito nenhum?",
        )
        self.assertEqual(
            pt.montar_pergunta(pt.por_id("voto_senador_segundo")),
            "E qual seria o segundo candidato que você votaria para senador?",
        )

    def test_campo_faltando_e_erro_ou_marcador(self):
        p = pt.por_id("ipv_candidato")
        with self.assertRaises(ValueError) as ctx:
            pt.montar_titulo(p, {"cargo": "Governador"})
        self.assertIn("candidato", str(ctx.exception))
        parcial = pt.montar_titulo(p, {"cargo": "Governador"}, parcial=True)
        self.assertIn("‹candidato›", parcial)
        self.assertTrue(parcial.endswith("(Estimulada e Única)"))

    def test_espacos_sobrando_sao_limpos(self):
        self.assertEqual(_t("rejeicao", cargo="  Deputado   Federal "), "Rejeição aos candidatos para Deputado Federal (Estimulada e Múltipla)")

    def test_linhas_do_sps_dobram_apostrofo(self):
        self.assertEqual(pt.linha_title_sps("Avaliação do d'Ávila"), "/TITLE 'Titulo: Avaliação do d''Ávila'")
        self.assertEqual(pt.linha_caption_sps("Quem é?"), "/CAPTION 'Pergunta: Quem é?'")


class TestBusca(unittest.TestCase):
    def test_sem_consulta_devolve_todos_e_filtra_por_grupo(self):
        self.assertEqual(len(pt.buscar("")), len(pt.PADROES))
        so_midia = pt.buscar("", pt.G_MIDIA)
        self.assertTrue(so_midia and all(p.grupo == pt.G_MIDIA for p in so_midia))

    def test_busca_sem_acento_nem_maiuscula(self):
        ids = [p.id for p in pt.buscar("REJEICAO")]
        self.assertEqual(ids[0], "rejeicao")
        self.assertIn("rejeicao", [p.id for p in pt.buscar("rejeição")])

    def test_todas_as_palavras_precisam_casar(self):
        ids = [p.id for p in pt.buscar("ipv apoiado")]
        self.assertEqual(ids[0], "ipv_apoiado_por")
        self.assertEqual(pt.buscar("palavra que nao existe em nenhum titulo"), [])

    def test_acha_pela_forma_antiga_do_relatorio(self):
        # quem lembra só do «Conhecimento inicial sobre» (Alphaville) também chega no padrão
        self.assertIn("conhecimento", [p.id for p in pt.buscar("conhecimento inicial")])
        self.assertIn("rejeicao", [p.id for p in pt.buscar("rejeição de voto")])

    def test_nome_pesa_mais_que_o_resto(self):
        self.assertEqual(pt.buscar("segundo turno")[0].id, "voto_segundo_turno")


class TestExcel(unittest.TestCase):
    def test_catalogo_em_excel(self):
        wb = openpyxl.load_workbook(io.BytesIO(pt.catalogo_excel()))
        ws = wb.active
        self.assertEqual([c.value for c in ws[1]], list(pt.COLUNAS_CATALOGO))
        self.assertEqual(ws.max_row, len(pt.PADROES) + 1)
        titulos = [ws.cell(r, 6).value for r in range(2, ws.max_row + 1)]
        self.assertIn("Rejeição aos candidatos para Governador (Estimulada e Múltipla)", titulos)

    def test_placeholder_da_proxima_etapa_nao_altera_nada(self):
        self.assertEqual(pt.aplicar_padronizacao_titulos(object()), 0)


if __name__ == "__main__":
    unittest.main()
