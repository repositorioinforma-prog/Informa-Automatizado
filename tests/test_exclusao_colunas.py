"""Exclusão de colunas por Base / Base reduzida (regras configuráveis)."""
import io
import unittest

import openpyxl
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from core.exclusao_colunas_math import (
    TIPO_BASE,
    TIPO_BASE_REDUZIDA,
    Regra,
    aplicar_exclusoes,
    escanear_planilha,
    excluir_colunas_por_regras,
    eh_nota_base_pequena,
    linhas_do_plano,
    listar_grupos,
    numero_da_base,
    planejar_exclusoes,
    revisar_notas_base_pequena,
    tipo_de_base,
)
from openpyxl.worksheet.pagebreak import Break

from core.exclusao_colunas_math import _remover_linhas_em_lote
from core.planilha_utils import excluir_colunas_base_reduzida, remover_linhas_seguro
from core.relatorio_automatizado_math import aplicar_codigo_08

FINA, MEDIA = Side(style="thin"), Side(style="medium")
AMARELO = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
AZUL = PatternFill("solid", start_color="0000FF", end_color="0000FF")

BR = (TIPO_BASE_REDUZIDA,)
BASE = (TIPO_BASE,)


def tabela(ws, topo=1, grupos=(("Região", ("R1", "R2", "R3", "R4")),), bases=(80, 60, 40, 10),
           tipo="Base reduzida", total=None, titulo="Titulo: Intenção de voto"):
    """
    Réplica do formato do relatório. Linhas (a partir de `topo`):
      +0 título (A..última mesclado) | +1 grupos | +2 segmentos |
      +3,+4 dados | +5 Base/Base reduzida | +6 Pergunta | +7 (em branco)
    Coluna B = Total; segmentos a partir da C.
    """
    segmentos = [s for _, segs in grupos for s in segs]
    assert len(segmentos) == len(bases)
    ult = 2 + len(segmentos)
    t, h1, h2, d1, d2, b, p = (topo + i for i in range(7))

    ws.cell(row=t, column=1, value=titulo)
    ws.cell(row=h1, column=2, value="Total")
    c = 3
    for nome, segs in grupos:
        ws.cell(row=h1, column=c, value=nome)
        for s in segs:
            ws.cell(row=h2, column=c, value=s)
            c += 1
    for r, rot in ((d1, "Candidato A"), (d2, "Candidato B")):
        ws.cell(row=r, column=1, value=rot)
        for col in range(2, ult + 1):
            ws.cell(row=r, column=col, value=50.0 + col).number_format = "0.0" if col != 5 else "0.00"
    ws.cell(row=b, column=1, value=tipo)
    ws.cell(row=b, column=2, value=total if total is not None else sum(v for v in bases if isinstance(v, (int, float))))
    for i, v in enumerate(bases):
        ws.cell(row=b, column=3 + i, value=v)
    ws.cell(row=p, column=1, value="Pergunta: Em quem você votaria?")

    for r in range(h1, b + 1):
        for col in range(1, ult + 1):
            ws.cell(row=r, column=col).border = Border(
                left=MEDIA if col == 1 else FINA, right=MEDIA if col == ult else FINA, top=FINA, bottom=FINA)
    for col in range(1, ult + 1):
        ws.cell(row=h2, column=col).font = Font(bold=True, name="DIN", size=10)

    ws.merge_cells(start_row=t, start_column=1, end_row=t, end_column=ult)
    ws.merge_cells(start_row=h1, start_column=1, end_row=h2, end_column=1)
    ws.merge_cells(start_row=h1, start_column=2, end_row=h2, end_column=2)
    c = 3
    for nome, segs in grupos:
        fim = c + len(segs) - 1
        ancora = ws.cell(row=h1, column=c)
        ancora.fill = AMARELO
        ancora.font = Font(bold=True, italic=True)
        ancora.alignment = Alignment(horizontal="center", wrap_text=True)
        # como no Excel/openpyxl: a borda externa da mesclagem fica na âncora
        ancora.border = Border(left=FINA, right=MEDIA if fim == ult else FINA, top=FINA, bottom=FINA)
        if fim > c:
            ws.merge_cells(start_row=h1, start_column=c, end_row=h1, end_column=fim)
        c = fim + 1
    return {"t": t, "h1": h1, "h2": h2, "d1": d1, "d2": d2, "b": b, "p": p, "ult": ult}


def recarregar(ws, vezes=1):
    for _ in range(vezes):
        buf = io.BytesIO()
        ws.parent.save(buf)
        ws = openpyxl.load_workbook(io.BytesIO(buf.getvalue()), rich_text=True).active
    return ws


def linha(ws, r, c_ini=1, c_fim=None):
    c_fim = c_fim or ws.max_column
    return [ws.cell(row=r, column=c).value for c in range(c_ini, c_fim + 1)]


def lado(ws, r, c, nome):
    s = getattr(ws.cell(row=r, column=c).border, nome)
    return s.style if s is not None else None


def mesclagens(ws):
    return sorted(str(r) for r in ws.merged_cells.ranges)


class TestNumeroDaBase(unittest.TestCase):
    def test_formatos(self):
        self.assertEqual(numero_da_base(604), 604.0)
        self.assertEqual(numero_da_base("25*"), 25.0)
        self.assertEqual(numero_da_base(" 25 * "), 25.0)
        self.assertEqual(numero_da_base("12,5"), 12.5)
        self.assertEqual(numero_da_base("1.234"), 1234.0)
        self.assertEqual(numero_da_base("1.234,5"), 1234.5)
        self.assertIsNone(numero_da_base("-"))
        self.assertIsNone(numero_da_base(None))
        self.assertIsNone(numero_da_base(True))
        self.assertIsNone(numero_da_base("Base"))

    def test_tipo_de_base(self):
        self.assertEqual(tipo_de_base("Base"), TIPO_BASE)
        self.assertEqual(tipo_de_base("  BASE REDUZIDA "), TIPO_BASE_REDUZIDA)
        self.assertEqual(tipo_de_base("Base:"), TIPO_BASE)
        self.assertIsNone(tipo_de_base("Base ponderada"))
        self.assertIsNone(tipo_de_base("Candidato"))
        self.assertIsNone(tipo_de_base(None))


class TestRegra(unittest.TestCase):
    def test_validacao_e_descricao(self):
        with self.assertRaises(ValueError):
            Regra(10, operador=">")
        with self.assertRaises(ValueError):
            Regra(10, linhas=("OUTRA",))
        self.assertEqual(Regra(25, BR).descricao(), "Base reduzida < 25")
        self.assertEqual(Regra(30, operador="<=").descricao(), "Base ou Base reduzida ≤ 30")
        self.assertIn("Região", Regra(25, BASE, grupos=("Região",)).descricao())

    def test_comparacao(self):
        self.assertFalse(Regra(25).casa_valor(25))
        self.assertTrue(Regra(25, operador="<=").casa_valor(25))
        self.assertTrue(Regra(25).casa_valor(24.9))


class TestEscaneamento(unittest.TestCase):
    def test_le_tabela_cabecalhos_e_bases(self):
        ws = Workbook().active
        tabela(ws, grupos=(("Sexo", ("M", "F")), ("Região", ("R1", "R2"))), bases=(50, 50, 30, 70))
        scan = escanear_planilha(ws)
        self.assertEqual(len(scan["tabelas"]), 1)
        t = scan["tabelas"][0]
        self.assertEqual(t["titulo"], "Intenção de voto")
        self.assertEqual((t["linha_ini"], t["linha_fim"], t["largura"]), (1, 7, 6))
        self.assertEqual(t["colunas"][2], {"grupo": "", "segmento": "Total"})
        self.assertEqual(t["colunas"][3], {"grupo": "Sexo", "segmento": "M"})
        self.assertEqual(t["colunas"][6], {"grupo": "Região", "segmento": "R2"})
        self.assertEqual(t["bases"][0]["tipo"], TIPO_BASE_REDUZIDA)
        self.assertEqual(t["bases"][0]["valores"][5], 30.0)
        self.assertEqual(scan["n_base_reduzida"], 1)
        self.assertEqual(listar_grupos([scan]), ["Sexo", "Região"])

    def test_tabelas_empilhadas_viram_tabelas_separadas(self):
        ws = Workbook().active
        tabela(ws, topo=1, tipo="Base")
        tabela(ws, topo=9, tipo="Base reduzida")
        scan = escanear_planilha(ws)
        self.assertEqual([(t["linha_ini"], t["linha_fim"]) for t in scan["tabelas"]], [(1, 7), (9, 15)])
        self.assertEqual((scan["n_base"], scan["n_base_reduzida"]), (1, 1))

    def test_tabelas_coladas_sem_linha_em_branco_sao_separadas(self):
        ws = Workbook().active
        tabela(ws, topo=1)
        tabela(ws, topo=8)   # título logo abaixo da "Pergunta:" da primeira
        scan = escanear_planilha(ws)
        self.assertEqual([(t["linha_ini"], t["linha_fim"]) for t in scan["tabelas"]], [(1, 7), (8, 14)])

    def test_aba_sem_base_e_rotulo_desconhecido(self):
        ws = Workbook().active
        ws.append(["Título"])
        ws.append(["Candidato A", 10, 20])
        ws.append(["Base ponderada", 100, 200])
        scan = escanear_planilha(ws)
        self.assertEqual(scan["tabelas"], [])
        self.assertEqual(scan["rotulos_base_desconhecidos"], ["Base ponderada"])
        self.assertEqual(scan["linhas_fora"], 3)


class TestExclusaoBasica(unittest.TestCase):
    def test_ultima_coluna_sai_e_contorno_fecha(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 60, 40, 10))
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual((resumo["colunas_excluidas"], resumo["tabelas_afetadas"]), (1, 1))
        ws = recarregar(ws, 2)
        nova = L["ult"] - 1
        self.assertEqual(linha(ws, L["h2"], 3, 6), ["R1", "R2", "R3", None])
        self.assertEqual(linha(ws, L["b"], 2, 6), [190, 80, 60, 40, None])
        self.assertEqual(mesclagens(ws), sorted(["A1:E1", "A2:A3", "B2:B3", "C2:E2"]))
        for r in range(L["h1"], L["b"] + 1):
            self.assertEqual(lado(ws, r, nova, "right"), "medium", f"linha {r}")
        self.assertEqual(lado(ws, L["d1"], nova - 1, "right"), "thin")   # divisória interna intocada
        self.assertIsNone(ws.cell(row=L["d1"], column=L["ult"]).fill.fgColor.rgb
                          if ws.cell(row=L["d1"], column=L["ult"]).fill.fill_type else None)

    def test_coluna_do_meio_desloca_valor_e_estilo(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 10, 60, 40))   # R2 (coluna D) sai; "0.00" estava na coluna E
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        ws = recarregar(ws)
        self.assertEqual(linha(ws, L["h2"], 3, 5), ["R1", "R3", "R4"])
        self.assertEqual(linha(ws, L["b"], 3, 5), [80, 60, 40])
        self.assertEqual(ws.cell(row=L["d1"], column=4).number_format, "0.00")   # veio da E
        self.assertEqual(ws.cell(row=L["d1"], column=5).number_format, "0.0")
        self.assertTrue(ws.cell(row=L["h2"], column=4).font.bold)
        self.assertEqual(ws.cell(row=L["h2"], column=4).font.name, "DIN")
        self.assertEqual(mesclagens(ws), sorted(["A1:E1", "A2:A3", "B2:B3", "C2:E2"]))
        self.assertEqual(lado(ws, L["d1"], 5, "right"), "medium")   # a última continua sendo a última

    def test_ancora_do_grupo_apagada_passa_texto_e_estilo_para_a_nova_ancora(self):
        ws = Workbook().active
        L = tabela(ws, grupos=(("Sexo", ("M", "F")), ("Região", ("R1", "R2", "R3"))),
                   bases=(100, 100, 5, 50, 50))   # R1 = coluna E = âncora de "Região"
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        ws = recarregar(ws, 2)
        self.assertEqual(linha(ws, L["h2"], 3, 6), ["M", "F", "R2", "R3"])
        self.assertIn("E2:F2", mesclagens(ws))
        anc = ws.cell(row=L["h1"], column=5)
        self.assertEqual(anc.value, "Região")
        self.assertEqual(anc.fill.fgColor.rgb[-6:], "FFFF00")
        self.assertTrue(anc.font.bold and anc.font.italic)
        self.assertEqual(anc.alignment.horizontal, "center")
        self.assertTrue(anc.alignment.wrap_text)
        self.assertEqual(lado(ws, L["h1"], 6, "right"), "medium")   # fim do grupo = fim da tabela

    def test_grupo_inteiro_apagado_some_com_o_titulo_do_grupo(self):
        ws = Workbook().active
        L = tabela(ws, grupos=(("Sexo", ("M", "F")), ("Região", ("R1", "R2", "R3"))),
                   bases=(100, 100, 5, 6, 7))
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        ws = recarregar(ws, 2)
        textos = [c.value for row in ws.iter_rows() for c in row if c.value]
        self.assertNotIn("Região", textos)
        self.assertIn("Sexo", textos)
        self.assertEqual(mesclagens(ws), sorted(["A1:D1", "A2:A3", "B2:B3", "C2:D2"]))
        self.assertEqual(lado(ws, L["h2"], 4, "right"), "medium")

    def test_celulas_vazias_a_direita_nao_ficam_com_preenchimento(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 60, 40, 10))
        ws.cell(row=L["d1"], column=L["ult"]).fill = AZUL
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        ws = recarregar(ws)
        cel = ws.cell(row=L["d1"], column=L["ult"])
        self.assertIsNone(cel.value)
        self.assertIsNone(cel.fill.fill_type)

    def test_linhas_fora_da_tabela_ficam_intactas(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 10, 60, 40))
        ws.cell(row=L["p"] + 1, column=1, value=None)
        ws.cell(row=L["p"] + 2, column=5, value="nota solta")      # depois da linha em branco
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(ws.cell(row=L["p"] + 2, column=5).value, "nota solta")
        self.assertEqual(ws.cell(row=L["p"], column=1).value, "Pergunta: Em quem você votaria?")

    def test_nada_a_excluir_nao_altera_a_planilha(self):
        ws = Workbook().active
        tabela(ws, bases=(80, 60, 40, 30))
        antes = (mesclagens(ws), [linha(ws, r) for r in range(1, 8)])
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(resumo["colunas_excluidas"], 0)
        self.assertEqual((mesclagens(ws), [linha(ws, r) for r in range(1, 8)]), antes)


class TestRegras(unittest.TestCase):
    def test_so_a_linha_escolhida_vale(self):
        ws = Workbook().active
        a = tabela(ws, topo=1, tipo="Base", bases=(80, 10, 60, 40))
        b = tabela(ws, topo=9, tipo="Base reduzida", bases=(80, 10, 60, 40))
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(linha(ws, a["b"], 3, 6), [80, 10, 60, 40])        # "Base": intacta
        self.assertEqual(linha(ws, b["b"], 3, 6), [80, 60, 40, None])      # "Base reduzida": mexeu

    def test_regra_para_base_e_base_reduzida_juntas(self):
        ws = Workbook().active
        a = tabela(ws, topo=1, tipo="Base", bases=(80, 10, 60, 40))
        b = tabela(ws, topo=9, tipo="Base reduzida", bases=(80, 60, 20, 40))
        resumo = excluir_colunas_por_regras(ws, [Regra(25)])
        self.assertEqual(resumo["tabelas_afetadas"], 2)
        self.assertEqual(linha(ws, a["b"], 3, 6), [80, 60, 40, None])
        self.assertEqual(linha(ws, b["b"], 3, 6), [80, 60, 40, None])

    def test_operadores(self):
        for operador, esperado in (("<", [80, 25, 40, None]), ("<=", [80, 40, None, None])):
            ws = Workbook().active
            L = tabela(ws, bases=(80, 25, 10, 40))
            excluir_colunas_por_regras(ws, [Regra(25, BR, operador=operador)])
            self.assertEqual(linha(ws, L["b"], 3, 6), esperado if operador == "<" else [80, 40, None, None])

    def test_filtro_por_grupo(self):
        ws = Workbook().active
        L = tabela(ws, grupos=(("Sexo", ("M", "F")), ("Região", ("R1", "R2"))), bases=(100, 5, 8, 90))
        excluir_colunas_por_regras(ws, [Regra(25, BR, grupos=("região",))])   # ignora caixa/acento
        ws = recarregar(ws)
        self.assertEqual(linha(ws, L["h2"], 3, 6), ["M", "F", "R2", None])
        self.assertEqual(linha(ws, L["b"], 3, 5), [100, 5, 90])   # "F" (5) ficou: não é Região

    def test_duas_regras_somam_exclusoes(self):
        ws = Workbook().active
        L = tabela(ws, grupos=(("Sexo", ("M", "F")), ("Região", ("R1", "R2"))), bases=(100, 5, 8, 90))
        excluir_colunas_por_regras(ws, [Regra(25, BR, grupos=("Região",)), Regra(10, BR, grupos=("Sexo",))])
        self.assertEqual(linha(ws, L["b"], 3, 4), [100, 90])

    def test_coluna_total_protegida_por_padrao(self):
        ws = Workbook().active
        L = tabela(ws, bases=(5, 6, 7, 8), total=20)    # Total (20) < 25
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(ws.cell(row=L["h1"], column=2).value, "Total")
        self.assertEqual(ws.cell(row=L["b"], column=2).value, 20)
        self.assertEqual(escanear_planilha(ws)["tabelas"][0]["largura"], 2)   # só A e B sobraram

    def test_total_sem_protecao(self):
        ws = Workbook().active
        L = tabela(ws, bases=(50, 60, 70, 80), total=20)
        excluir_colunas_por_regras(ws, [Regra(25, BR)], proteger_total=False)
        self.assertEqual(ws.cell(row=L["h1"], column=2).value, "Região")   # o grupo andou para a B

    def test_coluna_b_que_nao_e_total_nao_e_protegida(self):
        ws = Workbook().active
        L = tabela(ws, bases=(50, 60, 70, 80), total=20)
        ws.cell(row=L["h1"], column=2).value = "Masculino"   # B é um segmento, não o Total
        plano = planejar_exclusoes(escanear_planilha(ws), [Regra(25, BR)])
        self.assertEqual(plano["total_colunas"], 1)

    def test_celula_de_base_vazia(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 60, 40, 10))
        ws.cell(row=L["b"], column=4).value = None
        self.assertEqual(excluir_colunas_por_regras(ws, [Regra(25, BR)])["colunas_excluidas"], 1)
        ws = Workbook().active
        L = tabela(ws, bases=(80, 60, 40, 10))
        ws.cell(row=L["b"], column=4).value = None
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)], vazias_como_zero=True)
        self.assertEqual(resumo["colunas_excluidas"], 2)

    def test_base_pequena_ja_marcada_com_asterisco(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, "18*", 60, 40))
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(linha(ws, L["b"], 3, 5), [80, 60, 40])

    def test_tabela_que_fica_so_com_total_e_sinalizada(self):
        ws = Workbook().active
        tabela(ws, bases=(5, 6, 7, 8), total=100)
        plano = planejar_exclusoes(escanear_planilha(ws), [Regra(25, BR)])
        self.assertEqual(len(plano["tabelas_sem_segmentos"]), 1)


class TestVariasTabelas(unittest.TestCase):
    def test_cada_tabela_perde_as_proprias_colunas(self):
        ws = Workbook().active
        a = tabela(ws, topo=1, bases=(80, 10, 60, 40))
        b = tabela(ws, topo=9, bases=(5, 60, 40, 50))
        c = tabela(ws, topo=17, bases=(80, 60, 40, 50))
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual((resumo["colunas_excluidas"], resumo["tabelas_afetadas"]), (2, 2))
        ws = recarregar(ws, 2)
        self.assertEqual(linha(ws, a["h2"], 3, 6), ["R1", "R3", "R4", None])
        self.assertEqual(linha(ws, b["h2"], 3, 6), ["R2", "R3", "R4", None])
        self.assertEqual(linha(ws, c["h2"], 3, 6), ["R1", "R2", "R3", "R4"])
        self.assertEqual(lado(ws, b["h2"], 5, "right"), "medium")
        self.assertEqual(lado(ws, c["h2"], 6, "right"), "medium")
        self.assertEqual(mesclagens(ws).count("A9:E9"), 1)   # título da 2ª tabela encolheu
        self.assertIn("A17:F17", mesclagens(ws))             # a 3ª não mudou

    def test_tabelas_coladas_sem_linha_em_branco(self):
        ws = Workbook().active
        a = tabela(ws, topo=1, bases=(80, 10, 60, 40))
        b = tabela(ws, topo=8, bases=(5, 60, 40, 50))
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(linha(ws, a["h2"], 3, 5), ["R1", "R3", "R4"])
        self.assertEqual(linha(ws, b["h2"], 3, 5), ["R2", "R3", "R4"])

    def test_duas_linhas_de_base_na_mesma_tabela_apagam_uma_vez_so(self):
        ws = Workbook().active
        L = tabela(ws, tipo="Base", bases=(80, 10, 60, 40))
        ws.insert_rows(L["b"] + 1)
        ws.cell(row=L["b"] + 1, column=1, value="Base reduzida")
        for i, v in enumerate((30, 30, 12, 30)):
            ws.cell(row=L["b"] + 1, column=3 + i, value=v)
        excluir_colunas_por_regras(ws, [Regra(25)])   # B: R2; BR: R3  -> 2 colunas, uma passada
        self.assertEqual(linha(ws, L["h2"], 3, 4), ["R1", "R4"])
        self.assertEqual(linha(ws, L["b"], 3, 4), [80, 40])
        self.assertEqual(linha(ws, L["b"] + 1, 3, 4), [30, 30])


class TestLarguras(unittest.TestCase):
    def _larguras(self, ws, n=8):
        return [ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width for c in range(1, n + 1)]

    def test_larguras_acompanham_quando_todas_as_tabelas_apagam_o_mesmo(self):
        ws = Workbook().active
        tabela(ws, topo=1, bases=(80, 10, 60, 40))
        tabela(ws, topo=9, bases=(70, 12, 50, 45))
        for c, w in zip(range(1, 7), (30, 9, 11, 12, 13, 14)):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertTrue(resumo["larguras_ajustadas"])
        ws = recarregar(ws)
        self.assertEqual(self._larguras(ws, 5), [30, 9, 11, 13, 14])   # coluna D (12) saiu

    def test_larguras_agrupadas_no_arquivo(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 10, 60, 40))
        for c in range(3, 7):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 9
        ws.column_dimensions["F"].width = 20
        ws = recarregar(ws)   # o openpyxl agrupa C..E (mesma largura) em um só <col>
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        ws = recarregar(ws)
        self.assertEqual(self._larguras(ws, 6)[2:5], [9, 9, 20])

    def test_larguras_ficam_quando_as_tabelas_apagam_colunas_diferentes(self):
        ws = Workbook().active
        tabela(ws, topo=1, bases=(80, 10, 60, 40))
        tabela(ws, topo=9, bases=(5, 60, 40, 50))
        for c, w in zip(range(1, 7), (30, 9, 11, 12, 13, 14)):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertFalse(resumo["larguras_ajustadas"])
        self.assertEqual(self._larguras(ws, 6), [30, 9, 11, 12, 13, 14])

    def test_larguras_ficam_se_ha_outro_conteudo_na_aba(self):
        ws = Workbook().active
        tabela(ws, topo=1, bases=(80, 10, 60, 40))
        ws.cell(row=20, column=1, value="Observações")
        ws.cell(row=20, column=4, value="solta")
        for c, w in zip(range(1, 7), (30, 9, 11, 12, 13, 14)):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertFalse(resumo["larguras_ajustadas"])

    def test_opcao_de_nao_ajustar(self):
        ws = Workbook().active
        tabela(ws, bases=(80, 10, 60, 40))
        for c, w in zip(range(1, 7), (30, 9, 11, 12, 13, 14)):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)], ajustar_larguras=False)
        self.assertFalse(resumo["larguras_ajustadas"])
        self.assertEqual(self._larguras(ws, 6), [30, 9, 11, 12, 13, 14])


class TestAvisosEConteudos(unittest.TestCase):
    def test_formula_deslocada_gera_aviso(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 10, 60, 40))
        ws.cell(row=L["d1"], column=5, value="=C4+1")
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertTrue(any("fórmula" in a for a in resumo["avisos"]))

    def test_texto_rico_sobrevive(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 10, 60, 40))
        ws.cell(row=L["h2"], column=5).value = CellRichText(
            "Re", TextBlock(InlineFont(b=True, rFont="DIN", sz=9), "gião 3"))
        ws = recarregar(ws)
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        ws = recarregar(ws)
        v = ws.cell(row=L["h2"], column=4).value
        self.assertIsInstance(v, CellRichText)
        self.assertEqual(str(v), "Região 3")

    def test_aba_sem_base_nao_muda(self):
        ws = Workbook().active
        ws.append(["Título"])
        ws.append(["x", 1, 2])
        resumo = excluir_colunas_por_regras(ws, [Regra(25)])
        self.assertEqual(resumo["colunas_excluidas"], 0)
        self.assertEqual(linha(ws, 2), ["x", 1, 2])


class TestPlano(unittest.TestCase):
    def test_linhas_do_plano(self):
        ws = Workbook().active
        tabela(ws, grupos=(("Sexo", ("M", "F")), ("Região", ("R1", "R2"))), bases=(100, 5, 8, 90))
        plano = planejar_exclusoes(escanear_planilha(ws), [Regra(25, BR, grupos=("Região",))])
        linhas = linhas_do_plano(plano)
        self.assertEqual(len(linhas), 1)
        self.assertEqual(linhas[0]["Segmento"], "R1")
        self.assertEqual(linhas[0]["Grupo"], "Região")
        self.assertEqual(linhas[0]["Coluna"], "E")
        self.assertEqual(linhas[0]["Valor da base"], 8)
        self.assertEqual(linhas[0]["Linha de base"], "Base reduzida")
        self.assertTrue(linhas[0]["Regra"].startswith("1: Base reduzida < 25"))

    def test_planejar_nao_mexe_na_planilha(self):
        ws = Workbook().active
        tabela(ws, bases=(80, 10, 60, 40))
        antes = [linha(ws, r) for r in range(1, 8)]
        planejar_exclusoes(escanear_planilha(ws), [Regra(25, BR)])
        self.assertEqual([linha(ws, r) for r in range(1, 8)], antes)

    def test_plano_pode_ser_reaplicado_em_outra_copia_do_arquivo(self):
        """A tela planeja com uma carga do arquivo e aplica em outra."""
        def novo():
            ws = Workbook().active
            tabela(ws, topo=1, bases=(80, 10, 60, 40))
            tabela(ws, topo=9, bases=(5, 60, 40, 50))
            return ws
        plano = planejar_exclusoes(escanear_planilha(novo()), [Regra(25, BR)])
        ws2 = novo()
        resumo = aplicar_exclusoes(ws2, plano)
        self.assertEqual(resumo["colunas_excluidas"], 2)


class TestEquivalenciaComMotorAntigo(unittest.TestCase):
    """Para 'Base reduzida < N' o resultado (valores e mesclagens) tem que bater
    com o motor já usado no código 05 / Processador de Base Reduzida."""

    def _montar(self):
        ws = Workbook().active
        tabela(ws, topo=1, grupos=(("Sexo", ("M", "F")), ("Região", ("R1", "R2", "R3"))),
               bases=(100, 100, 5, 50, 20))
        tabela(ws, topo=9, bases=(5, 60, 40, 50))
        tabela(ws, topo=17, bases=(80, 60, 40, 50))
        tabela(ws, topo=25, grupos=(("Região", ("R1", "R2", "R3")),), bases=(30, 12, 9))
        return ws

    def test_mesmo_resultado(self):
        antigo, novo = self._montar(), self._montar()
        n_antigo = excluir_colunas_base_reduzida(antigo, 25)
        resumo = excluir_colunas_por_regras(novo, [Regra(25, BR)], proteger_total=False)
        self.assertEqual(resumo["colunas_excluidas"], n_antigo)
        self.assertEqual(mesclagens(novo), mesclagens(antigo))
        for r in range(1, antigo.max_row + 1):
            self.assertEqual(linha(novo, r, 1, 8), linha(antigo, r, 1, 8), f"linha {r}")
        for r in range(1, antigo.max_row + 1):          # contorno direito igual onde há tabela
            for c in range(1, 8):
                if novo.cell(row=r, column=c).value is None and antigo.cell(row=r, column=c).value is None:
                    # célula vazia: o motor antigo deixa aqui uma borda direita "órfã" quando a tabela é
                    # mais estreita que a aba (ele limpa só a última coluna da ABA); o novo apaga a célula
                    continue
                self.assertEqual(lado(novo, r, c, "right"), lado(antigo, r, c, "right"), f"({r},{c})")


NOTA = "*Base pequena para análise estatística"


def _linhas_com_texto(ws, texto):
    return [r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == texto]


def _titulos(ws):
    return [r for r in range(1, ws.max_row + 1) if str(ws.cell(row=r, column=1).value or "").startswith("Titulo")]


class TestRevisaoNotaBasePequena(unittest.TestCase):
    """Depois de apagar as colunas: tabela com base '*' mantém a nota; sem base '*', a nota sai."""

    def _relatorio(self):
        ws = Workbook().active
        tabela(ws, topo=1, bases=(80, 10, 60, 28), titulo="Titulo: T1 (sobra 28*)")
        tabela(ws, topo=9, bases=(80, 12, 60, 40), titulo="Titulo: T2 (só 12*)")
        tabela(ws, topo=17, bases=(80, 70, 60, 40), titulo="Titulo: T3 (sem base pequena)")
        tabela(ws, topo=25, bases=(80, 15, 14, 29), titulo="Titulo: T4 (sobra 29*)")
        tabela(ws, topo=33, bases=(5, 70, 60, 40), titulo="Titulo: T5 (só 5*)")
        aplicar_codigo_08(ws)   # marca <30 com '*' e insere uma nota por tabela afetada
        return ws

    def test_pre_condicao_codigo_08(self):
        ws = self._relatorio()
        self.assertEqual(len(_linhas_com_texto(ws, NOTA)), 4)   # T1, T2, T4, T5 (T3 não tem base < 30)

    def test_mantem_com_asterisco_e_remove_sem_asterisco(self):
        ws = self._relatorio()
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        # limite 25: sai 10*, 12*, 15*, 14*, 5* ; sobram 28* (T1) e 29* (T4)
        self.assertEqual((resumo["notas_removidas"], resumo["notas_mantidas"]), (2, 2))
        ws = recarregar(ws, 2)
        self.assertEqual(len(_linhas_com_texto(ws, NOTA)), 2)
        # cada nota que ficou está colada na tabela certa: logo depois da "Pergunta:" da T1 e da T4
        for r in _linhas_com_texto(ws, NOTA):
            self.assertTrue(str(ws.cell(row=r - 1, column=1).value).startswith("Pergunta"))
        titulos = {ws.cell(row=r, column=1).value: r for r in _titulos(ws)}
        nota_t1, nota_t4 = _linhas_com_texto(ws, NOTA)
        self.assertTrue(titulos["Titulo: T1 (sobra 28*)"] < nota_t1 < titulos["Titulo: T2 (só 12*)"])
        self.assertTrue(titulos["Titulo: T4 (sobra 29*)"] < nota_t4 < titulos["Titulo: T5 (só 5*)"])

    def test_tabelas_de_baixo_sobem_inteiras(self):
        ws = self._relatorio()
        antes = {ws.cell(row=r, column=1).value: r for r in _titulos(ws)}
        # altura marcada no cabeçalho da T3 e uma mesclagem (título) — devem acompanhar a subida
        t3 = antes["Titulo: T3 (sem base pequena)"]
        ws.row_dimensions[t3 + 2].height = 33
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        depois = {ws.cell(row=r, column=1).value: r for r in _titulos(ws)}
        self.assertEqual(depois["Titulo: T3 (sem base pequena)"], t3 - 1)       # 1 nota (T2) removida acima
        self.assertEqual(depois["Titulo: T5 (só 5*)"], antes["Titulo: T5 (só 5*)"] - 1)
        self.assertEqual(ws.row_dimensions[t3 - 1 + 2].height, 33)
        self.assertIsNone(ws.row_dimensions[t3 + 2].height)   # a altura antiga não ficou para trás
        self.assertIn(f"A{t3 - 1}:F{t3 - 1}", mesclagens(ws))   # título da T3 (nenhuma coluna saiu dela) subiu junto

    def test_sem_exclusoes_ainda_revisa_notas_orfas(self):
        ws = Workbook().active
        tabela(ws, bases=(80, 70, 60, 40))          # nenhuma base < 30, mas alguém deixou a nota
        ws.insert_rows(9)
        ws.cell(row=9, column=1, value=NOTA)
        ws.cell(row=8, column=1, value="Pergunta: Em quem você votaria?")
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual((resumo["colunas_excluidas"], resumo["notas_removidas"]), (0, 1))
        self.assertEqual(_linhas_com_texto(ws, NOTA), [])

    def test_opcao_desligada_nao_mexe_nas_notas(self):
        ws = self._relatorio()
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)], revisar_notas=False)
        self.assertEqual(resumo["notas_removidas"], 0)
        self.assertEqual(len(_linhas_com_texto(ws, NOTA)), 4)

    def test_nota_separada_por_linha_em_branco_pertence_a_tabela_anterior(self):
        ws = Workbook().active
        t1 = tabela(ws, topo=1, bases=(80, 12, 60, 40))      # 12 sai; nenhum '*' (número puro)
        ws.cell(row=t1["p"] + 2, column=1, value=NOTA)       # nota depois de uma linha em branco
        t2 = tabela(ws, topo=11, bases=(80, 70, 20, "29*"))   # 29* sobra
        ws.cell(row=t2["p"] + 2, column=1, value=NOTA)
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual((resumo["notas_removidas"], resumo["notas_mantidas"]), (1, 1))
        self.assertEqual(len(_linhas_com_texto(ws, NOTA)), 1)
        self.assertGreater(_linhas_com_texto(ws, NOTA)[0], 11)   # a que ficou é a da 2ª tabela

    def test_nota_antes_da_primeira_tabela_nao_e_tocada(self):
        ws = Workbook().active
        ws.cell(row=1, column=1, value=NOTA)
        tabela(ws, topo=3, bases=(80, 70, 60, 40))
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(resumo["notas_removidas"], 0)
        self.assertEqual(ws.cell(row=1, column=1).value, NOTA)

    def test_asterisco_sem_nota_gera_aviso_e_nao_insere_nada(self):
        ws = Workbook().active
        L = tabela(ws, bases=(80, 70, 60, "28*"), titulo="Titulo: Sem nota")
        n_linhas = ws.max_row
        resumo = excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(resumo["tabelas_com_asterisco_sem_nota"], ["Sem nota"])
        self.assertEqual(ws.max_row, n_linhas)

    def test_idempotente(self):
        ws = self._relatorio()
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        foto = [linha(ws, r, 1, 7) for r in range(1, ws.max_row + 1)]
        resumo = revisar_notas_base_pequena(ws)
        self.assertEqual(resumo["notas_removidas"], 0)
        self.assertEqual([linha(ws, r, 1, 7) for r in range(1, ws.max_row + 1)], foto)

    def test_previa_do_plano_bate_com_o_que_a_revisao_faz(self):
        ws = self._relatorio()
        plano = planejar_exclusoes(escanear_planilha(ws), [Regra(25, BR)])
        previa = plano["revisao_notas"]
        self.assertEqual(len(previa["remover"]), 2)
        self.assertEqual(previa["mantidas"], 2)
        self.assertEqual(sorted(i["titulo"] for i in previa["remover"]), ["T2 (só 12*)", "T5 (só 5*)"])
        resumo = aplicar_exclusoes(ws, plano)
        self.assertEqual((resumo["notas_removidas"], resumo["notas_mantidas"]), (2, 2))

    def test_nota_mesclada_tambem_e_removida(self):
        ws = self._relatorio()
        nota_t2 = _linhas_com_texto(ws, NOTA)[1]
        ws.merge_cells(start_row=nota_t2, start_column=1, end_row=nota_t2, end_column=4)
        excluir_colunas_por_regras(ws, [Regra(25, BR)])
        self.assertEqual(len(_linhas_com_texto(ws, NOTA)), 2)
        self.assertNotIn(f"A{nota_t2}:D{nota_t2}", mesclagens(ws))   # a mesclagem não ficou órfã

    def test_variacao_do_texto_da_nota(self):
        self.assertTrue(eh_nota_base_pequena("  *BASE PEQUENA para analise estatistica"))
        self.assertFalse(eh_nota_base_pequena("Base pequena"))      # sem o '*'
        self.assertFalse(eh_nota_base_pequena(None))


class TestRemocaoDeLinhasEmLote(unittest.TestCase):
    def _montar(self):
        ws = Workbook().active
        for topo in (1, 9, 17, 25):
            tabela(ws, topo=topo)
        for r in range(1, ws.max_row + 1):
            ws.row_dimensions[r].height = 12 + r      # altura diferente em cada linha
        for r in (8, 16, 24):
            ws.row_breaks.append(Break(id=r))
        return ws

    def _foto(self, ws):
        return (
            mesclagens(ws),
            [linha(ws, r, 1, 7) for r in range(1, ws.max_row + 1)],
            {r: d.height for r, d in ws.row_dimensions.items() if d.height},
            [b.id for b in ws.row_breaks.brk],
            {(r, c): ws.cell(row=r, column=c).border.right.style for r in range(1, 20) for c in range(1, 8)},
        )

    def test_equivale_a_remover_uma_a_uma(self):
        linhas = [4, 5, 12, 20]       # 4 e 5 vizinhas; as outras soltas (todas dentro de tabelas)
        lote, uma = self._montar(), self._montar()
        _remover_linhas_em_lote(lote, linhas)
        for r in sorted(linhas, reverse=True):
            remover_linhas_seguro(uma, r, 1)
        f_lote, f_uma = self._foto(lote), self._foto(uma)
        self.assertEqual(f_lote[0], f_uma[0], "mesclagens")
        self.assertEqual(f_lote[1], f_uma[1], "valores")
        self.assertEqual(f_lote[2], f_uma[2], "alturas")
        self.assertEqual(f_lote[3], f_uma[3], "quebras de página")
        self.assertEqual(f_lote[4], f_uma[4], "bordas")

    def test_quebra_de_pagina_logo_depois_de_linha_removida_nao_se_perde(self):
        ws = self._montar()
        ws.row_breaks.brk = [Break(id=8)]            # a página acaba na linha 8 (em branco)
        _remover_linhas_em_lote(ws, [8])             # ... e a linha 8 sai
        self.assertEqual([b.id for b in ws.row_breaks.brk], [7])   # agora acaba na linha de cima

    def test_quebras_continuam_em_ordem_e_sem_repeticao(self):
        ws = self._montar()
        ws.row_breaks.brk = [Break(id=7), Break(id=8), Break(id=16)]
        _remover_linhas_em_lote(ws, [8])             # 7 e 8 viram a mesma quebra
        ids = [b.id for b in ws.row_breaks.brk]
        self.assertEqual(ids, sorted(set(ids)))
        self.assertEqual(ids, [7, 15])

    def test_lista_vazia_nao_faz_nada(self):
        ws = self._montar()
        antes = self._foto(ws)
        _remover_linhas_em_lote(ws, [])
        self.assertEqual(self._foto(ws), antes)

if __name__ == "__main__":
    unittest.main()
