"""
Motor de "InserirCapasResultados" (código 15) — o mais complexo dos 16
códigos originais do VBA.

Insere duas "capas" de separação no relatório:
    - "RESULTADOS PELOS SEGMENTOS", logo antes da primeira tabela que
      tem alguma segmentação (Sexo, Idade, Renda, Escolaridade,
      Religião, Regiões etc. — identificado pela própria linha de
      cabeçalho conter "Total" + uma palavra de segmento conhecida).
    - "RESULTADOS PELO TOTAL", logo antes da primeira tabela do
      relatório (a primeira ocorrência de "Total" em qualquer célula),
      precedida por 6 páginas em branco (espaço reservado pra
      sumário/capa geral do documento impresso).

Cada capa é um bloco de 7 linhas (título grande centralizado na 4ª) +
2 linhas em branco depois, com uma célula-marcador invisível (texto
branco, tamanho 1) na primeira linha do bloco — usada só pra
reidentificar o bloco depois (aplicar quebras de página) e pra permitir
rodar o código de novo sem duplicar capas (remove as antigas antes de
inserir as novas).

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

from core.planilha_utils import inserir_linhas_seguro, remover_linhas_seguro
from core.relatorio_automatizado_math import _ultima_linha_com_conteudo

COL_INICIO = 1  # A
COL_FIM = 13  # M
PAGINAS_BRANCAS_ANTES_TOTAL = 6
LINHAS_BLOCO_TITULO = 7
LINHA_DO_TITULO_NO_BLOCO = 4
LINHAS_EM_BRANCO_APOS_CAPA = 2
MARCADOR_TOTAL = "__CAPA_TOTAL__"
MARCADOR_SEGMENTOS = "__CAPA_SEGMENTOS__"

PALAVRAS_SEGMENTO = [
    "SEXO", "IDADE", "FAIXA ETARIA", "FAIXA ETÁRIA", "RENDA", "ESCOLARIDADE",
    "RELIGIAO", "RELIGIÃO", "REGIAO", "REGIÃO", "CATOLICAS", "CATÓLICAS",
    "EVANGELICAS", "EVANGÉLICAS", "GOVL", "GOVM", "GOVO",
    "APROVACAO", "APROVAÇÃO", "AVALIACAO", "AVALIAÇÃO",
]


# ---------------------------------------------------------------------------
# Leitura de linhas (equivalentes às funções auxiliares "AM" do VBA)
# ---------------------------------------------------------------------------
def _texto_linha(ws, linha):
    partes = []
    for c in range(COL_INICIO, COL_FIM + 1):
        v = ws.cell(row=linha, column=c).value
        if v is not None and str(v).strip() != "":
            partes.append(str(v).strip().upper())
    return " ".join(partes)


def _linha_tem_total(ws, linha):
    for c in range(COL_INICIO, COL_FIM + 1):
        v = ws.cell(row=linha, column=c).value
        if v is not None and str(v).strip().upper() == "TOTAL":
            return True
    return False


def _linha_tem_conteudo(ws, linha):
    for c in range(COL_INICIO, COL_FIM + 1):
        v = ws.cell(row=linha, column=c).value
        if v is not None and str(v).strip() != "":
            return True
    return False


def _contem_palavra_segmento(texto):
    return any(p in texto for p in PALAVRAS_SEGMENTO)


def _encontrar_linha_cabecalho_total(ws, max_row):
    for r in range(1, max_row + 1):
        if _linha_tem_total(ws, r):
            return r
    return 0


def _encontrar_linha_cabecalho_segmentado(ws, linha_inicial, max_row):
    for r in range(linha_inicial, max_row + 1):
        texto = _texto_linha(ws, r)
        if "TOTAL" in texto and _contem_palavra_segmento(texto):
            return r
    return 0


def _eh_linha_titulo(ws, linha):
    """Uma linha 'Titulo: ...' é o marcador semântico mais confiável do
    início de um bloco de pergunta/tabela — muito mais preciso do que
    só "subir enquanto tiver conteúdo" (que nas versões anteriores
    acabava chegando até a linha 1, incluindo cabeçalhos/títulos
    institucionais que não fazem parte de bloco nenhum)."""
    valor = ws.cell(row=linha, column=1).value
    return isinstance(valor, str) and valor.strip().lower().startswith("titulo")


def _encontrar_inicio_bloco(ws, linha_ref):
    """
    Acha o início "de verdade" do bloco de pergunta/tabela que contém
    `linha_ref` — primeiro tenta achar a linha 'Titulo: ...' mais
    próxima IGUAL ou ACIMA de linha_ref (identificação semântica: o
    título é sempre a primeira linha de um bloco, nesse formato de
    relatório). Só cai pro comportamento conservador antigo ("sobe
    enquanto a linha anterior tiver conteúdo") se não achar nenhum
    'Titulo:' no caminho — o que não deveria acontecer num relatório
    bem formado, mas evita quebrar em arquivos fora do padrão.
    """
    r = linha_ref
    while r >= 1:
        if _eh_linha_titulo(ws, r):
            return r
        r -= 1

    # Fallback conservador (comportamento antigo) — só chega aqui se não
    # existir NENHUMA linha "Titulo:" entre linha_ref e o topo do arquivo
    r = linha_ref
    while r > 1 and _linha_tem_conteudo(ws, r - 1):
        r -= 1
    return r


def _encontrar_linha_marcador(ws, marcador, max_row):
    for r in range(1, max_row + 1):
        if str(ws.cell(row=r, column=1).value or "").strip() == marcador:
            return r
    return 0


# ---------------------------------------------------------------------------
# Configuração de página / quebras
# ---------------------------------------------------------------------------
def _configurar_pagina_am(ws):
    ultima_linha = _ultima_linha_com_conteudo(ws)
    ws.print_area = f"A1:{get_column_letter(COL_FIM)}{ultima_linha}"
    # Escala 100% (não "ajustar largura") — ver mesma explicação no
    # código 07 (core/relatorio_automatizado_math.py), que configura a
    # página do mesmo jeito.
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.fitToWidth = None
    ws.page_setup.fitToHeight = None
    ws.page_setup.scale = 100


def _obter_linhas_por_pagina(ws):
    """Usa a 1ª quebra de página manual já existente (normalmente inserida
    pelo código 07) pra descobrir quantas linhas cabem numa página."""
    ids = sorted(b.id for b in ws.row_breaks.brk if b.id is not None)
    if ids:
        return ids[0]
    return 45


def _adicionar_quebra(ws, linha_id):
    if linha_id >= 1 and not any(b.id == linha_id for b in ws.row_breaks.brk):
        ws.row_breaks.append(Break(id=linha_id))


def _limpar_quebras_no_intervalo(ws, linha_inicial, linha_final):
    ini = max(1, linha_inicial)
    if linha_final < 1:
        return
    ws.row_breaks.brk = [b for b in ws.row_breaks.brk if not (ini <= b.id <= linha_final)]


def _aplicar_quebras_bloco_total(ws, linha_inicio_titulo, linhas_por_pagina):
    inicio_paginas_brancas = linha_inicio_titulo - (PAGINAS_BRANCAS_ANTES_TOTAL * linhas_por_pagina)
    linha_quebra_depois_titulo = linha_inicio_titulo + LINHAS_BLOCO_TITULO

    _limpar_quebras_no_intervalo(ws, inicio_paginas_brancas - 2, linha_quebra_depois_titulo + 2)

    for i in range(0, PAGINAS_BRANCAS_ANTES_TOTAL + 1):
        linha_quebra = inicio_paginas_brancas + i * linhas_por_pagina
        if linha_quebra > 1:
            _adicionar_quebra(ws, linha_quebra - 1)

    _adicionar_quebra(ws, linha_quebra_depois_titulo - 1)


def _aplicar_quebras_bloco_normal(ws, linha_inicio_titulo):
    linha_quebra_depois_titulo = linha_inicio_titulo + LINHAS_BLOCO_TITULO
    _limpar_quebras_no_intervalo(ws, linha_inicio_titulo - 2, linha_quebra_depois_titulo + 2)
    if linha_inicio_titulo > 1:
        _adicionar_quebra(ws, linha_inicio_titulo - 1)
    _adicionar_quebra(ws, linha_quebra_depois_titulo - 1)


def _aplicar_quebras_resultados(ws, linhas_por_pagina):
    max_row = _ultima_linha_com_conteudo(ws)
    linha_total = _encontrar_linha_marcador(ws, MARCADOR_TOTAL, max_row)
    if linha_total:
        _aplicar_quebras_bloco_total(ws, linha_total, linhas_por_pagina)

    linha_segmentos = _encontrar_linha_marcador(ws, MARCADOR_SEGMENTOS, max_row)
    if linha_segmentos:
        _aplicar_quebras_bloco_normal(ws, linha_segmentos)


# ---------------------------------------------------------------------------
# Inserção / limpeza das capas
# ---------------------------------------------------------------------------
def _inserir_bloco_resultado(ws, linha_base, titulo, marcador, paginas_brancas_antes, linhas_por_pagina):
    linhas_paginas_brancas = paginas_brancas_antes * linhas_por_pagina
    total_linhas_inserir = linhas_paginas_brancas + LINHAS_BLOCO_TITULO + LINHAS_EM_BRANCO_APOS_CAPA

    inserir_linhas_seguro(ws, linha_base, total_linhas_inserir)

    linha_inicio_titulo = linha_base + linhas_paginas_brancas
    altura_padrao = ws.sheet_format.defaultRowHeight or 15

    if linhas_paginas_brancas > 0:
        for i in range(linha_base, linha_inicio_titulo):
            ws.row_dimensions[i].height = altura_padrao

    for i in range(linha_inicio_titulo, linha_inicio_titulo + LINHAS_BLOCO_TITULO):
        ws.row_dimensions[i].height = 18

    linha_titulo = linha_inicio_titulo + LINHA_DO_TITULO_NO_BLOCO - 1
    ws.row_dimensions[linha_titulo].height = 42

    ws.row_dimensions[linha_inicio_titulo + LINHAS_BLOCO_TITULO].height = 18
    ws.row_dimensions[linha_inicio_titulo + LINHAS_BLOCO_TITULO + 1].height = 18

    # marcador invisível (texto branco, tamanho 1) — só pra reidentificar
    # o bloco depois (quebras de página, limpeza numa próxima execução)
    cel_marcador = ws.cell(row=linha_inicio_titulo, column=1, value=marcador)
    cel_marcador.font = Font(size=1, color="FFFFFFFF")

    for rng in list(ws.merged_cells.ranges):
        if rng.min_row == linha_titulo == rng.max_row:
            ws.unmerge_cells(str(rng))
    for c in range(COL_INICIO, COL_FIM + 1):
        ws.cell(row=linha_titulo, column=c, value=None)
    ws.merge_cells(start_row=linha_titulo, start_column=COL_INICIO, end_row=linha_titulo, end_column=COL_FIM)

    cel_titulo = ws.cell(row=linha_titulo, column=COL_INICIO)
    cel_titulo.value = titulo
    cel_titulo.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=False)
    cel_titulo.font = Font(name="DIN", size=28, bold=True)

    return linha_inicio_titulo


def _limpar_capas_antigas(ws, linhas_por_pagina):
    """Remove capas inseridas numa execução anterior deste código (busca
    pelos marcadores invisíveis) — permite rodar de novo sem duplicar."""
    while True:
        max_row = ws.max_row
        encontrou = False
        for i in range(max_row, 0, -1):
            valor = str(ws.cell(row=i, column=1).value or "").strip()
            if valor == MARCADOR_TOTAL:
                linha_inicio = max(1, i - PAGINAS_BRANCAS_ANTES_TOTAL * linhas_por_pagina)
                linha_fim = i + LINHAS_BLOCO_TITULO + LINHAS_EM_BRANCO_APOS_CAPA - 1
                remover_linhas_seguro(ws, linha_inicio, linha_fim - linha_inicio + 1)
                encontrou = True
                break
            if valor == MARCADOR_SEGMENTOS:
                linha_inicio = i
                linha_fim = i + LINHAS_BLOCO_TITULO + LINHAS_EM_BRANCO_APOS_CAPA - 1
                remover_linhas_seguro(ws, linha_inicio, linha_fim - linha_inicio + 1)
                encontrou = True
                break
        if not encontrou:
            break


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------
def aplicar_codigo_15(ws):
    """
    Insere as capas "RESULTADOS PELO TOTAL" e "RESULTADOS PELOS
    SEGMENTOS" no relatório (ver docstring do módulo). Pode ser
    reaplicado sem duplicar — remove as capas de uma execução anterior
    antes de inserir as novas.

    Depende de quebras de página manuais já existirem no relatório
    (inseridas pelo código 07), usadas pra calcular quantas linhas cabem
    numa página impressa.

    Returns:
        dict com "status" ("ok" ou "erro"), e em caso de sucesso as
        linhas onde cada capa foi inserida.
    """
    _configurar_pagina_am(ws)
    linhas_por_pagina = _obter_linhas_por_pagina(ws)
    if linhas_por_pagina <= 0:
        linhas_por_pagina = 45

    _limpar_capas_antigas(ws, linhas_por_pagina)
    _configurar_pagina_am(ws)

    max_row = _ultima_linha_com_conteudo(ws)
    linha_cab_total = _encontrar_linha_cabecalho_total(ws, max_row)
    if linha_cab_total == 0:
        return {"status": "erro", "mensagem": "Não encontrei a primeira tabela do total."}

    linha_inicio_total = _encontrar_inicio_bloco(ws, linha_cab_total)
    linha_cab_seg = _encontrar_linha_cabecalho_segmentado(ws, linha_cab_total + 1, max_row)

    linha_capa_segmentos = None
    if linha_cab_seg:
        linha_inicio_seg = _encontrar_inicio_bloco(ws, linha_cab_seg)
        # insere primeiro o bloco de baixo (Segmentos) — inserir linhas
        # ali não afeta a posição do bloco do Total, que fica acima
        linha_capa_segmentos = _inserir_bloco_resultado(
            ws, linha_inicio_seg, "RESULTADOS PELOS SEGMENTOS", MARCADOR_SEGMENTOS,
            0, linhas_por_pagina,
        )

    linha_capa_total = _inserir_bloco_resultado(
        ws, linha_inicio_total, "RESULTADOS PELO TOTAL", MARCADOR_TOTAL,
        PAGINAS_BRANCAS_ANTES_TOTAL, linhas_por_pagina,
    )

    _configurar_pagina_am(ws)
    _aplicar_quebras_resultados(ws, linhas_por_pagina)
    _configurar_pagina_am(ws)

    # a posição da capa de Segmentos pode ter sido deslocada pela
    # inserção da capa do Total (que fica acima dela no arquivo) —
    # busca a posição final de verdade em vez de devolver a de antes
    # desse deslocamento
    if linha_capa_segmentos is not None:
        max_row_final = _ultima_linha_com_conteudo(ws)
        linha_capa_segmentos = _encontrar_linha_marcador(ws, MARCADOR_SEGMENTOS, max_row_final)

    return {
        "status": "ok",
        "linha_capa_total": linha_capa_total,
        "linha_capa_segmentos": linha_capa_segmentos,
    }


def validar_capas(ws, resultado):
    """
    Confere programaticamente que a estrutura das capas ficou correta —
    pra ser chamada DEPOIS de salvar e reabrir o arquivo (não em cima do
    `ws` ainda em memória antes de salvar), pra pegar qualquer coisa que
    o ciclo salvar/reabrir do Excel possa ter derrubado. Levanta
    ValueError com uma mensagem clara se algo estiver errado, em vez de
    devolver silenciosamente um arquivo incorreto.

    Args:
        ws: a planilha já reaberta (via openpyxl.load_workbook, depois
            de salvar) do resultado de `aplicar_codigo_15`.
        resultado: o dict devolvido por `aplicar_codigo_15`.

    Returns:
        True se tudo estiver certo (levanta exceção caso contrário).
    """
    erros = []
    max_row = ws.max_row

    linha_total = _encontrar_linha_marcador(ws, MARCADOR_TOTAL, max_row)
    if not linha_total:
        erros.append(f"Marcador {MARCADOR_TOTAL} não encontrado após salvar/reabrir.")
    else:
        linha_titulo_total = linha_total + LINHA_DO_TITULO_NO_BLOCO - 1
        valor = ws.cell(row=linha_titulo_total, column=1).value
        if valor != "RESULTADOS PELO TOTAL":
            erros.append(
                f"Esperava 'RESULTADOS PELO TOTAL' na linha {linha_titulo_total}, achei {valor!r}."
            )
        tem_mesclagem = any(
            r.min_row == linha_titulo_total == r.max_row and r.min_col == COL_INICIO and r.max_col == COL_FIM
            for r in ws.merged_cells.ranges
        )
        if not tem_mesclagem:
            erros.append(f"Mesclagem A:M não encontrada na linha do título ({linha_titulo_total}).")

        # confere as 6 páginas em branco: precisa haver pelo menos 6
        # quebras manuais entre o início do bloco de páginas brancas e a capa
        ids_quebras = sorted(b.id for b in ws.row_breaks.brk if b.id is not None)
        quebras_antes_do_total = [i for i in ids_quebras if i < linha_total]
        if len(quebras_antes_do_total) < PAGINAS_BRANCAS_ANTES_TOTAL:
            erros.append(
                f"Esperava pelo menos {PAGINAS_BRANCAS_ANTES_TOTAL} quebras de página "
                f"antes da capa do Total, achei {len(quebras_antes_do_total)}."
            )
        linha_depois_bloco_total = linha_total + LINHAS_BLOCO_TITULO
        if (linha_depois_bloco_total - 1) not in ids_quebras:
            erros.append(
                f"Esperava uma quebra de página logo depois do bloco da capa do "
                f"Total (id={linha_depois_bloco_total - 1})."
            )

    linha_seg = _encontrar_linha_marcador(ws, MARCADOR_SEGMENTOS, max_row)
    if resultado.get("linha_capa_segmentos") and not linha_seg:
        erros.append(f"Marcador {MARCADOR_SEGMENTOS} não encontrado após salvar/reabrir.")
    elif linha_seg:
        linha_titulo_seg = linha_seg + LINHA_DO_TITULO_NO_BLOCO - 1
        valor = ws.cell(row=linha_titulo_seg, column=1).value
        if valor != "RESULTADOS PELOS SEGMENTOS":
            erros.append(
                f"Esperava 'RESULTADOS PELOS SEGMENTOS' na linha {linha_titulo_seg}, achei {valor!r}."
            )
        ids_quebras = sorted(b.id for b in ws.row_breaks.brk if b.id is not None)
        if (linha_seg - 1) not in ids_quebras:
            erros.append(f"Esperava uma quebra de página logo antes da capa dos Segmentos (id={linha_seg - 1}).")
        linha_depois_bloco_seg = linha_seg + LINHAS_BLOCO_TITULO
        if (linha_depois_bloco_seg - 1) not in ids_quebras:
            erros.append(
                f"Esperava uma quebra de página logo depois do bloco da capa dos "
                f"Segmentos (id={linha_depois_bloco_seg - 1})."
            )

    area_esperada = f"A1:{get_column_letter(COL_FIM)}{_ultima_linha_com_conteudo(ws)}"
    # openpyxl devolve o print_area formatado como "'NomeDaAba'!$A$1:$M$136"
    # — compara só a parte da faixa de células, sem se importar com o
    # nome da aba nem com "$"
    area_atual_normalizada = str(ws.print_area or "").replace("$", "")
    if area_esperada not in area_atual_normalizada:
        erros.append(f"Área de impressão esperada terminando em '{area_esperada}', achei {ws.print_area!r}.")

    # Nota: o documento de referência original pedia fitToWidth=1 (modo
    # "ajustar largura"), mas isso foi substituído pelo modo "escala
    # 100%" (fitToPage=False + scale=100) num pedido posterior — os dois
    # modos são mutuamente exclusivos no Excel, então validamos o que
    # realmente configuramos agora, não os dois ao mesmo tempo.
    if ws.sheet_properties.pageSetUpPr.fitToPage:
        erros.append("fitToPage deveria estar desligado (modo escala 100%), veio ligado.")
    if ws.page_setup.scale != 100:
        erros.append(f"scale deveria ser 100, veio {ws.page_setup.scale!r}.")

    if erros:
        raise ValueError("Validação das capas falhou depois de salvar/reabrir:\n" + "\n".join(f"- {e}" for e in erros))

    return True
