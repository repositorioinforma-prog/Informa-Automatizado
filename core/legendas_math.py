"""
Motor de "Legendas": localiza, num relatório de tabelas, os blocos cuja
segmentação é por região/território (Regiões, Bairros, Municípios etc.)
e insere o texto de legenda correspondente (vindo de um segundo arquivo)
logo depois da linha "Pergunta:" de cada bloco.

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
import openpyxl

from core.base_multiplas_math import normalizar_texto

# Nomes de categoria (grupo da tabela) reconhecidos como "segmentação
# territorial" — qualquer bloco cujo grupo bata com um destes (comparação
# tolerante a acento/caixa/hífen, via normalizar_texto) é candidato a
# receber legenda.
CATEGORIAS_LEGENDA = [
    "Regiões",
    "Mesorregiões",
    "Microrregiões",
    "Regiões Intermediárias",
    "Regiões Imediatas",
    "Municípios",
    "Cidades",
    "Bairros",
    "Regiões Administrativas",
    "RPA",
    "Regiões da Capital",
    "Capital",
    "Sub-Regiões",
    "Sub-regiões",
    "Subregiões",
]
CATEGORIAS_LEGENDA_NORM = {normalizar_texto(c) for c in CATEGORIAS_LEGENDA}


def _eh_linha_vazia(linha):
    return linha is None or all(v is None for v in linha)


def _forward_fill(linha):
    resultado = list(linha)
    atual = None
    for i, v in enumerate(resultado):
        if v is not None and str(v).strip() != "":
            atual = v
        resultado[i] = atual
    return resultado


def parsear_blocos_tabela(caminho_ou_arquivo, aba=None):
    """
    Parseia um relatório de tabelas SEM o marcador 'Titulo:' (o título é
    só o texto puro na primeira linha não-vazia de cada bloco) — formato
    usado neste tipo de exportação. Cada bloco é delimitado por linhas em
    branco e termina numa linha 'Pergunta:'.

    Retorna uma lista de dicts:
      {"titulo_idx": linha do título (1-indexada), "titulo": texto,
       "grupos": [...], "categorias": [...], "linha_pergunta": nº da
       linha 'Pergunta:' (1-indexada) ou None}
    """
    wb = openpyxl.load_workbook(caminho_ou_arquivo, data_only=True, rich_text=True)
    ws = wb[aba] if aba else wb[wb.sheetnames[0]]
    linhas = list(ws.iter_rows(values_only=True))
    n = len(linhas)
    blocos = []
    i = 0

    while i < n:
        if _eh_linha_vazia(linhas[i]):
            i += 1
            continue

        titulo_idx = i + 1
        titulo = linhas[i][0]
        i += 1
        if i >= n:
            break

        linha_grupo = linhas[i]
        i += 1
        grupos = _forward_fill(linha_grupo)

        tem_segmentacao = len(linha_grupo) > 2 and any(v is not None for v in linha_grupo[2:])
        if tem_segmentacao and i < n:
            categorias = list(linhas[i])
            i += 1
        else:
            categorias = [None] * len(linha_grupo)

        linha_pergunta_idx = None
        while i < n:
            linha = linhas[i]
            primeiro = linha[0] if linha else None
            if isinstance(primeiro, str) and primeiro.strip().startswith("Pergunta:"):
                linha_pergunta_idx = i + 1
                i += 1
                break
            i += 1

        blocos.append({
            "titulo_idx": titulo_idx,
            "titulo": titulo,
            "grupos": grupos,
            "categorias": categorias,
            "linha_pergunta": linha_pergunta_idx,
        })

    return blocos


def bloco_precisa_legenda(bloco):
    """Bloco é candidato a legenda se algum de seus grupos bater com CATEGORIAS_LEGENDA."""
    textos = {normalizar_texto(g) for g in bloco["grupos"] if g is not None}
    return bool(textos & CATEGORIAS_LEGENDA_NORM)


def parsear_blocos_legenda(caminho_ou_arquivo, aba=None):
    """
    Parseia o arquivo de legendas: blocos que começam com uma linha cujo
    texto COMEÇA com 'LEGENDA' — pode ser só 'LEGENDA', ou com um rótulo
    depois (ex.: 'LEGENDA CAPITAL', 'LEGENDA REGIÕES') — seguidos de
    linhas de texto (uma por item), até uma linha em branco ou o próximo
    cabeçalho 'LEGENDA...'.

    Carrega com rich_text=True para preservar formatação parcial dentro
    da célula (ex.: "Região 1 [9,09%]:" em negrito, o resto normal) — sem
    isso, o openpyxl achata a célula toda para texto simples e o negrito
    se perde.

    Retorna uma lista de dicts: {"itens": [...], "rotulo": str|None} —
    'rotulo' é o texto normalizado depois de 'LEGENDA' (None se o
    cabeçalho for só 'LEGENDA', sem nada depois). Cada item pode ser uma
    string simples ou um CellRichText (preserva os trechos em negrito).
    """
    wb = openpyxl.load_workbook(caminho_ou_arquivo, rich_text=True)
    ws = wb[aba] if aba else wb[wb.sheetnames[0]]
    linhas = list(ws.iter_rows(values_only=True))
    n = len(linhas)
    blocos = []
    i = 0

    def _rotulo_legenda(texto):
        """Se 'texto' começa com 'LEGENDA', devolve o rótulo normalizado depois dela (ou None)."""
        t = texto.strip().upper()
        if not t.startswith("LEGENDA"):
            return None, False
        resto = texto.strip()[len("LEGENDA"):].strip()
        return (normalizar_texto(resto) if resto else None), True

    while i < n:
        primeiro = linhas[i][0] if linhas[i] else None
        rotulo, eh_legenda = _rotulo_legenda(primeiro) if isinstance(primeiro, str) else (None, False)
        if eh_legenda:
            i += 1
            itens = []
            while i < n and not _eh_linha_vazia(linhas[i]):
                v = linhas[i][0]
                _, v_eh_legenda = _rotulo_legenda(v) if isinstance(v, str) else (None, False)
                if v_eh_legenda:
                    break
                if v is not None:
                    itens.append(v)
                i += 1
            blocos.append({"itens": itens, "rotulo": rotulo})
        else:
            i += 1

    return blocos


def _assinatura_bloco(bloco):
    """
    'Assinatura' de um bloco elegível: (categoria territorial reconhecida,
    valores exatos das categorias). Blocos com a MESMA assinatura usam a
    mesma legenda — é assim que uma segmentação como "Regiões" (que se
    repete idêntica numa tabela por pergunta/candidato) recebe a legenda
    certa em TODAS as ocorrências, não só na primeira.

    Importante usar os valores exatos (não só a quantidade): uma mesma
    categoria territorial pode aparecer dividida em mais de uma tabela
    quando há categorias demais pra caber numa só (ex.: "Regiões 1-7" e
    "Regiões 8-14" — mesma categoria territorial, mesma quantidade de
    colunas, mas valores diferentes e legendas diferentes). Usar só a
    quantidade juntaria essas duas por engano.
    """
    textos = {normalizar_texto(g) for g in bloco["grupos"] if g is not None}
    reconhecidas = sorted(textos & CATEGORIAS_LEGENDA_NORM)
    categoria_territorial = reconhecidas[0] if reconhecidas else None
    valores_categoria = tuple(
        normalizar_texto(c) for c in bloco["categorias"] if c is not None
    )
    return (categoria_territorial, valores_categoria)


def parear_blocos(blocos_tabela_elegiveis, blocos_legenda):
    """
    Agrupa os blocos elegíveis por "tipo" de segmentação territorial
    (assinatura: categoria + valores exatos das categorias) e associa
    cada tipo distinto a UM bloco de legenda — essa mesma legenda é
    reaplicada a TODAS as tabelas daquele tipo ao longo do relatório, não
    só à primeira ocorrência (a mesma segmentação, ex.: "Regiões",
    costuma se repetir em dezenas de tabelas — uma por pergunta/candidato
    — mas o arquivo de legendas normalmente só define a legenda uma vez
    por tipo).

    Se um bloco de legenda tiver um rótulo explícito que bata com uma das
    categorias territoriais reconhecidas (ex.: 'LEGENDA CAPITAL', 'LEGENDA
    REGIÕES'), esse rótulo é usado para casar direto com o tipo certo —
    mais confiável do que depender só da ordem em que os blocos aparecem
    no arquivo. Blocos de legenda sem rótulo (só 'LEGENDA') continuam
    sendo usados na ordem em que aparecem, como reserva para os tipos que
    não tiverem um bloco rotulado correspondente.

    Retorna (pares, avisos), onde pares é uma lista de (bloco_tabela, bloco_legenda) —
    um item por TABELA (podendo repetir o mesmo bloco_legenda várias vezes).
    """
    avisos = []

    # separa os blocos de legenda com rótulo reconhecido dos genéricos
    fila_por_rotulo = {}
    fila_generica = []
    for bl in blocos_legenda:
        rotulo = bl.get("rotulo")
        if rotulo and rotulo in CATEGORIAS_LEGENDA_NORM:
            fila_por_rotulo.setdefault(rotulo, []).append(bl)
        else:
            fila_generica.append(bl)

    ordem_assinaturas = []
    vistos = set()
    for bt in blocos_tabela_elegiveis:
        sig = _assinatura_bloco(bt)
        if sig not in vistos:
            vistos.add(sig)
            ordem_assinaturas.append(sig)

    assinatura_para_legenda = {}
    indice_generico = 0
    for sig in ordem_assinaturas:
        categoria_territorial = sig[0]
        bl = None

        # 1) tenta um bloco de legenda com rótulo explícito batendo com o tipo
        if categoria_territorial and fila_por_rotulo.get(categoria_territorial):
            bl = fila_por_rotulo[categoria_territorial].pop(0)
        # 2) senão, cai pro próximo bloco genérico (sem rótulo), na ordem do arquivo
        elif indice_generico < len(fila_generica):
            bl = fila_generica[indice_generico]
            indice_generico += 1

        if bl is not None:
            assinatura_para_legenda[sig] = bl
        else:
            avisos.append(
                f"Não sobrou nenhum bloco de legenda disponível para o tipo "
                f"'{categoria_territorial or '?'}'."
            )

    sobras_rotulo = sum(len(fila) for fila in fila_por_rotulo.values())
    sobras_generica = len(fila_generica) - indice_generico
    if sobras_rotulo or sobras_generica:
        avisos.append(
            f"{sobras_rotulo + sobras_generica} bloco(s) de legenda sobraram sem "
            "nenhum tipo de segmentação territorial correspondente."
        )

    pares = []
    avisados_tipo = set()
    for bt in blocos_tabela_elegiveis:
        sig = _assinatura_bloco(bt)
        bl = assinatura_para_legenda.get(sig)
        if bl is None:
            continue

        n_categorias = len(sig[1])
        if n_categorias and len(bl["itens"]) != n_categorias and sig not in avisados_tipo:
            avisados_tipo.add(sig)
            avisos.append(
                f"Tipo '{sig[0] or '?'}' ({n_categorias} categoria(s) na tabela): a "
                f"legenda correspondente tem {len(bl['itens'])} item(ns) — confira se "
                "a ordem/quantidade bate."
            )
        pares.append((bt, bl))

    return pares, avisos


def gerar_workbook_com_legendas(caminho_tabela, pares, aba=None):
    """
    Insere, no PRÓPRIO workbook original (carregado a partir de
    'caminho_tabela'), o texto de legenda de cada par (bloco_tabela,
    bloco_legenda) logo depois da linha 'Pergunta:' daquele bloco — mais
    precisamente: mantém a primeira linha em branco após 'Pergunta:' como
    está, insere "LEGENDA" + os itens, e mais uma linha em branco de
    espaçamento antes do restante do arquivo continuar normalmente.

    Diferente de uma versão anterior, NÃO reconstrói a planilha do zero —
    reconstruir descartava configurações do arquivo original que não
    são copiadas por padrão pelo openpyxl (orientação da página, quebras
    de página manuais, área de impressão etc.). Em vez disso, usa
    `ws.insert_rows()` diretamente no arquivo carregado, o que preserva
    tudo isso automaticamente, já que é o mesmo objeto de planilha.

    O único cuidado necessário com `insert_rows()` é com células
    mescladas: o openpyxl não desloca corretamente faixas mescladas que
    estejam na região afetada, então cada faixa mesclada em/abaixo do
    ponto de inserção é desfeita antes de inserir e refeita depois, já
    deslocada.
    """
    from copy import copy as _copy_style

    from openpyxl.styles import Alignment, Font

    wb = openpyxl.load_workbook(caminho_tabela, rich_text=True)
    nome_aba = aba or wb.sheetnames[0]
    ws = wb[nome_aba]

    font_item = Font(name="DIN Book", size=9, bold=False)
    alinhamento_esquerda = Alignment(horizontal="left")

    # (linha a manter como está, itens de legenda, linha do título do
    # bloco — usada como referência de estilo do cabeçalho "LEGENDA").
    # Processa do bloco mais abaixo para o mais acima: assim, inserir
    # linhas num bloco não bagunça os índices de linha dos blocos que
    # ainda faltam processar (todos mais acima na planilha).
    insercoes = []
    for bt, bl in pares:
        if bt["linha_pergunta"]:
            insercoes.append((bt["linha_pergunta"] + 1, bl["itens"], bt["titulo_idx"]))
    insercoes.sort(key=lambda x: x[0], reverse=True)

    for linha_manter, itens, titulo_idx in insercoes:
        linha_insercao = linha_manter + 1
        n_novas_linhas = len(itens) + 2  # "LEGENDA" + itens + 1 linha em branco no final

        # desfaz temporariamente as mesclagens em/abaixo do ponto de
        # inserção (só essas — as de cima não são afetadas e continuam
        # como estavam)
        faixas_a_remesclar = []
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row >= linha_insercao:
                faixas_a_remesclar.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
                ws.unmerge_cells(str(rng))

        # `insert_rows` também não desloca a ALTURA customizada de linha
        # (row_dimensions) — ela fica "grudada" no número antigo da linha
        # em vez de acompanhar o conteúdo que se moveu, o que faz uma
        # linha alta e vazia sobrar bem no meio do relatório, longe de
        # onde a formatação original fazia sentido. Guarda e remove antes
        # de inserir, para recolocar já deslocada depois.
        alturas_a_deslocar = {}
        for idx in list(ws.row_dimensions.keys()):
            if idx >= linha_insercao:
                alturas_a_deslocar[idx] = ws.row_dimensions.pop(idx)

        ws.insert_rows(linha_insercao, amount=n_novas_linhas)

        # refaz as mesclagens desfeitas, já na posição deslocada
        for min_row, min_col, max_row, max_col in faixas_a_remesclar:
            ws.merge_cells(
                start_row=min_row + n_novas_linhas, start_column=min_col,
                end_row=max_row + n_novas_linhas, end_column=max_col
            )

        # recoloca as alturas de linha guardadas, já na posição deslocada
        for idx, dim in alturas_a_deslocar.items():
            nova_idx = idx + n_novas_linhas
            dim.index = nova_idx
            ws.row_dimensions[nova_idx] = dim

        # `insert_rows` não desloca quebras de página manuais sozinho —
        # faz isso à mão para toda quebra que estava em/abaixo do ponto
        # de inserção
        for quebra in ws.row_breaks.brk:
            if quebra.id >= linha_insercao:
                quebra.id += n_novas_linhas

        # preenche o conteúdo nas linhas recém-inseridas (title_idx é
        # sempre anterior ao ponto de inserção deste próprio bloco, então
        # continua válido e sem deslocamento neste momento do laço)
        cel_titulo_ref = ws.cell(row=titulo_idx, column=1)

        r = linha_insercao
        cel_legenda = ws.cell(row=r, column=1, value="LEGENDA")
        cel_legenda.font = _copy_style(cel_titulo_ref.font)
        cel_legenda.alignment = alinhamento_esquerda
        r += 1

        for item in itens:
            cel_item = ws.cell(row=r, column=1, value=item)
            cel_item.font = font_item
            cel_item.alignment = alinhamento_esquerda
            r += 1
        # a última linha inserida (r) fica em branco de propósito —
        # espaçamento antes do restante do arquivo continuar

    _expandir_area_impressao(ws)

    return wb


def _expandir_area_impressao(ws):
    """
    `insert_rows` não expande a área de impressão sozinho — se o arquivo
    tinha uma área de impressão definida, garante que ela cubra até a
    última linha atual da planilha, pra as novas linhas de legenda não
    ficarem de fora na hora de imprimir/exportar em PDF.
    """
    if not ws.print_area:
        return

    from openpyxl.utils.cell import range_boundaries, get_column_letter

    areas = ws.print_area if isinstance(ws.print_area, list) else [ws.print_area]
    novas_areas = []
    for area in areas:
        ref = area.split("!")[-1]
        try:
            min_col, min_row, max_col, max_row = range_boundaries(ref)
        except ValueError:
            novas_areas.append(area)
            continue
        novo_max_row = max(max_row, ws.max_row)
        novas_areas.append(
            f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{novo_max_row}"
        )
    ws.print_area = novas_areas
