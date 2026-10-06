"""
Utilidades genéricas para manipular planilhas .xlsx, reaproveitadas por
vários dos códigos portados do "Relatório Automatizado".

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
import math

import pandas as pd
from openpyxl.utils import get_column_letter


def _desfazer_mesclagem_rapido(ws, rng):
    """
    Equivalente a `ws.unmerge_cells(str(rng))`, mas sem a checagem de
    pertencimento que o método oficial faz — `ws.unmerge_cells()`
    confere se a faixa existe varrendo TODAS as mesclagens do arquivo
    (`coord in self.merged_cells`, uma comparação de sobreposição, não
    um hash simples), o que é caro em arquivos com centenas/milhares de
    mesclagens e sempre redundante aqui: `rng` já veio de
    `ws.merged_cells.ranges`, então sabidamente já existe.

    Faz o mesmo trabalho por baixo — remove a faixa da lista de
    mesclagens e apaga as células "filhas" que ela criou (preservando a
    célula-âncora, a primeira, que carrega o valor de verdade).
    """
    ws.merged_cells.ranges.discard(rng)
    cells = rng.cells
    next(cells)  # pula a célula-âncora (primeira), igual ao unmerge_cells oficial
    for row, col in cells:
        ws._cells.pop((row, col), None)


def _remesclar_rapido(ws, min_row, min_col, max_row, max_col):
    """
    Equivalente a `ws.merge_cells(...)`, mas sem a checagem de
    sobreposição contra TODAS as mesclagens já existentes que
    `ws.merge_cells()` faz por padrão (comparação de sobreposição, não
    um hash simples — o mesmo custo caro do lado do
    `_desfazer_mesclagem_rapido`, só que na hora de mesclar de novo).
    Redundante aqui: estamos recriando uma mesclagem que a gente mesmo
    desfez há pouco (só deslocada por causa de uma inserção/remoção de
    linha), então sabidamente não sobrepõe nada.

    Preserva a parte legítima do `ws.merge_cells()` oficial — a
    formatação de borda nas bordas da mesclagem —, só pula a checagem
    de sobreposição em si.
    """
    from openpyxl.worksheet.cell_range import CellRange
    from openpyxl.worksheet.merge import MergedCellRange

    cr = CellRange(min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col)
    mcr = MergedCellRange(ws, cr.coord)
    ws.merged_cells.ranges.add(mcr)
    ws._clean_merge_range(mcr)


def inserir_linhas_seguro(ws, linha_insercao, quantidade):
    """
    Insere 'quantidade' linhas em branco em 'linha_insercao', deslocando
    corretamente mesclagens, alturas de linha customizadas e quebras de
    página manuais que estejam em/abaixo do ponto de inserção — coisas
    que o `ws.insert_rows()` sozinho não desloca (comportamento já
    mapeado e corrigido no módulo de Legendas).
    """
    if quantidade <= 0:
        return

    faixas_a_remesclar = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= linha_insercao:
            faixas_a_remesclar.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
            _desfazer_mesclagem_rapido(ws, rng)

    alturas_a_deslocar = {}
    for idx in list(ws.row_dimensions.keys()):
        if idx >= linha_insercao:
            alturas_a_deslocar[idx] = ws.row_dimensions.pop(idx)

    ws.insert_rows(linha_insercao, amount=quantidade)

    for min_row, min_col, max_row, max_col in faixas_a_remesclar:
        _remesclar_rapido(
            ws, min_row=min_row + quantidade, min_col=min_col,
            max_row=max_row + quantidade, max_col=max_col,
        )

    for idx, dim in alturas_a_deslocar.items():
        nova_idx = idx + quantidade
        dim.index = nova_idx
        ws.row_dimensions[nova_idx] = dim

    for quebra in ws.row_breaks.brk:
        if quebra.id >= linha_insercao:
            quebra.id += quantidade


def remover_linhas_seguro(ws, linha_remocao, quantidade):
    """
    Remove 'quantidade' linhas a partir de 'linha_remocao' (inclusive),
    deslocando corretamente mesclagens, alturas de linha customizadas e
    quebras de página manuais que estejam abaixo da faixa removida —
    espelha `inserir_linhas_seguro`, na direção contrária. Só deve ser
    usada em linhas que já se sabe estarem vazias (o chamador é
    responsável por confirmar isso antes) — não existe checagem de
    conteúdo aqui.
    """
    if quantidade <= 0:
        return

    linha_fim_removida = linha_remocao + quantidade - 1

    # Desfaz mesclagens que tocam a faixa removida (não deveria acontecer
    # se o chamador só remove linhas vazias, mas por segurança) e as que
    # estão abaixo dela (serão remescladas na posição nova)
    faixas_a_remesclar = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row > linha_fim_removida:
            faixas_a_remesclar.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
            _desfazer_mesclagem_rapido(ws, rng)

    alturas_a_deslocar = {}
    for idx in list(ws.row_dimensions.keys()):
        if idx >= linha_remocao:
            dim = ws.row_dimensions.pop(idx)
            if idx > linha_fim_removida:
                alturas_a_deslocar[idx] = dim
            # alturas dentro da própria faixa removida são descartadas

    ws.delete_rows(linha_remocao, amount=quantidade)

    for min_row, min_col, max_row, max_col in faixas_a_remesclar:
        _remesclar_rapido(
            ws, min_row=min_row - quantidade, min_col=min_col,
            max_row=max_row - quantidade, max_col=max_col,
        )

    for idx, dim in alturas_a_deslocar.items():
        nova_idx = idx - quantidade
        dim.index = nova_idx
        ws.row_dimensions[nova_idx] = dim

    quebras_restantes = []
    for quebra in ws.row_breaks.brk:
        if quebra.id >= linha_remocao and quebra.id <= linha_fim_removida:
            continue  # quebra estava dentro da faixa removida, descarta
        if quebra.id > linha_fim_removida:
            quebra.id -= quantidade
        quebras_restantes.append(quebra)
    ws.row_breaks.brk = quebras_restantes


def estimar_numero_linhas(texto, largura_total_colunas, tamanho_fonte):
    """
    Estima quantas linhas de texto (por quebra automática) uma célula
    com quebra de linha vai precisar, dada a largura total das colunas
    e o tamanho da fonte — mesma aproximação usada em
    `estimar_altura_calculada` (ver docstring lá pras ressalvas/
    calibração), só que devolve o número de linhas cru, sem multiplicar
    por nenhuma altura-por-linha. Usado onde a altura final precisa ser
    uma fórmula direta em cima do número de linhas (ex.: código 12 —
    altura da Pergunta = nº de linhas × 15).

    Texto com quebra MANUAL (\\n) é tratado trecho por trecho — cada
    trecho entre quebras conta as próprias linhas por quebra automática
    (por largura), e o total é a SOMA delas — nunca o texto inteiro
    como uma sequência só (isso subestimava a altura sempre que havia
    quebra manual: um texto como "Renda Média Mensal Domiciliar\\n(em
    Salário Mínimo)" tem 2 trechos que, cada um, já pode precisar de
    mais de uma linha por quebra automática).

    Dentro de cada trecho, a quebra automática é SIMULADA palavra por
    palavra (empacotando palavras numa linha até estourar a largura,
    igual o Excel faz) — não apenas `caracteres_do_trecho /
    caracteres_por_linha`. Essa divisão simples ignora que o Excel não
    quebra no meio de uma palavra: uma palavra longa que não cabe no
    resto da linha força a quebra ali, sobrando espaço vazio no fim da
    linha anterior — então o número de linhas reais costuma ser maior
    do que a divisão simples prevê, principalmente com palavras longas
    (nomes com barra, hífen, números). Confirmado com renderização real
    (LibreOffice) em dois casos: um texto com quebra
    manual («Renda Média...») e um título sem quebra manual nenhuma,
    só automática («Intenção de voto para Presidente da República/2026
    - Cenário Único (Estimulada e Única)») — os dois precisavam de mais
    linhas do que a divisão simples calculava.
    """
    if not texto:
        return 1
    largura_pt = (largura_total_colunas + 0.62) * 7 * 0.90
    # o 1.15 é margem de segurança adicional calibrada com os dois casos
    # reais acima — a largura de caractere "crua" (tamanho*0.62) por si
    # só já não bastava, mesmo simulando quebra por palavra.
    largura_char = tamanho_fonte * 0.62 * 1.15
    caracteres_por_linha = max(1, int(largura_pt / largura_char))
    trechos = str(texto).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    total = sum(_simular_linhas_quebra_palavra(trecho, caracteres_por_linha) for trecho in trechos)
    return max(1, total)


def _simular_linhas_quebra_palavra(trecho, caracteres_por_linha):
    """Empacota as palavras de `trecho` em linhas de até
    `caracteres_por_linha` caracteres, uma palavra de cada vez — a
    mesma lógica de quebra automática do Excel (nunca quebra no meio de
    uma palavra). Uma única palavra maior que `caracteres_por_linha`
    ocupa a linha sozinha (não força mais de uma linha por si só nesta
    aproximação)."""
    palavras = trecho.split(" ")
    linhas = 1
    ocupado_na_linha = 0
    for palavra in palavras:
        tamanho = len(palavra)
        if ocupado_na_linha == 0:
            ocupado_na_linha = tamanho
        elif ocupado_na_linha + 1 + tamanho <= caracteres_por_linha:
            ocupado_na_linha += 1 + tamanho
        else:
            linhas += 1
            ocupado_na_linha = tamanho
    return max(1, linhas)


# ---------------------------------------------------------------------------
# Altura das labels da coluna A — regra fixa por nº de caracteres
# ---------------------------------------------------------------------------
# Regra definida pelo Lucas (coluna A com largura 21): até 23 caracteres
# (contando espaços) -> altura 15; até 46 -> 30; até 69 -> 45; e assim
# por diante (15 por "faixa" de 23 caracteres). Primeira versão usava 21;
# aumentada para 23 a pedido do Lucas. Substitui, só para as
# labels das tabelas, a estimativa por largura/fonte (que dava alturas
# "quebradas" tipo 16 ou 32 e deixava o relatório irregular).
CARACTERES_POR_LINHA_LABEL = 23
ALTURA_POR_LINHA_LABEL = 15

_PREFIXOS_NAO_LABEL = ("pergunta:", "titulo:", "título:")


def altura_label_por_caracteres(
    texto,
    caracteres_por_linha=CARACTERES_POR_LINHA_LABEL,
    altura_por_linha=ALTURA_POR_LINHA_LABEL,
):
    """Altura da label pela regra fixa de caracteres (23 -> 15, 46 -> 30...).

    Conta todos os caracteres, inclusive espaços. Se a label tiver quebra
    manual (\\n, ex.: rótulos formatados pelo código 04), cada trecho conta
    as próprias faixas e o total é somado — "Aluno em domicílio\\n(Educação
    Pública)" = 1 + 1 faixa = 30. Espaços sobrando nas pontas de cada
    trecho (comuns em exportação do SPSS) são ignorados.
    """
    if texto is None:
        return altura_por_linha
    trechos = str(texto).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    linhas = 0
    for trecho in trechos:
        n = len(trecho.strip())
        linhas += max(1, -(-n // caracteres_por_linha))  # teto da divisão
    return max(1, linhas) * altura_por_linha


# --- Medição por largura real (fonte proporcional) ------------------------
# A contagem de caracteres sozinha falha com textos de letras largas: a
# DIN é proporcional ("W", "R", "m" ocupam 3-4x um "i"/"l"), e o Excel
# não quebra no meio da palavra. Caso real: "Waguinho - Republicanos" tem
# 23 caracteres (regra -> 15), mas na coluna A (21) o Excel jogava
# "Republicanos" pra linha de baixo e o texto ficava cortado.
#
# Por isso a altura da label é o MAIOR entre:
#   1) a regra do Lucas por caracteres (23 -> 15, 46 -> 30...) — piso;
#   2) a quebra simulada palavra por palavra usando a largura real de
#      cada caractere (métricas de Helvetica, proporção bem próxima da
#      DIN; unidades de 1/1000 do tamanho da fonte).
#
# LARGURA_UTIL_LABEL_COLUNA_A: largura útil (nessas unidades, fonte 10)
# de uma célula da coluna A (largura 21 + ajuste). Calibrada entre as
# labels de 22-23 caracteres que cabem numa linha (~10.500-10.700) e o
# caso real que estourou ("Waguinho - Republicanos" = 11.389). Se algum
# caso ainda cortar, é só baixar este número; se aparecer label alta
# demais, subir.
LARGURA_UTIL_LABEL_COLUNA_A = 11000
LARGURA_REFERENCIA_COLUNA_A = 21.78  # 21 + AJUSTE_LARGURA_COLUNA_DIN

_LARGURAS_CARACTERES = {}
for _largura, _chars in (
    (222, "'`ijl"), (260, "|"), (278, " !,./:;I[\\]ft"), (333, "()-r"),
    (334, "{}"), (355, '"'), (389, "*"), (469, "^"), (500, "Jcksvxyz"),
    (556, "#$0123456789?L_abdeghnopqu"), (584, "+<=>~"), (611, "FTZ"),
    (667, "&ABEKPSVXY"), (722, "CDHNRUw"), (778, "GOQ"), (833, "Mm"),
    (889, "%"), (944, "W"), (1015, "@"),
):
    for _ch in _chars:
        _LARGURAS_CARACTERES[_ch] = _largura
_LARGURA_CARACTERE_PADRAO = 556


def largura_texto_unidades(texto):
    """Largura do texto em 1/1000 do tamanho da fonte (acentos contam
    como a letra base: 'ã' = 'a')."""
    import unicodedata
    total = 0
    for ch in str(texto):
        base = unicodedata.normalize("NFKD", ch)[:1] or ch
        total += _LARGURAS_CARACTERES.get(base, _LARGURA_CARACTERE_PADRAO)
    return total


def linhas_por_largura(texto, largura_util):
    """Quantas linhas o Excel usa para `texto` numa célula com
    `largura_util` (unidades de `largura_texto_unidades`), simulando a
    quebra palavra por palavra. Quebras manuais (\\n) somam trechos.
    Palavra maior que a linha inteira é partida (como o Excel faz)."""
    largura_util = max(1, largura_util)
    espaco = _LARGURAS_CARACTERES[" "]
    trechos = str(texto).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    total = 0
    for trecho in trechos:
        palavras = trecho.strip().split()
        linhas = 1
        ocupado = 0
        for palavra in palavras:
            w = largura_texto_unidades(palavra)
            if ocupado == 0:
                if w > largura_util:
                    extras = -(-w // largura_util) - 1
                    linhas += extras
                    w -= extras * largura_util
                ocupado = w
            elif ocupado + espaco + w <= largura_util:
                ocupado += espaco + w
            else:
                linhas += 1
                if w > largura_util:
                    extras = -(-w // largura_util) - 1
                    linhas += extras
                    w -= extras * largura_util
                ocupado = w
        total += linhas
    return max(1, total)


def altura_label_celula(ws, cel):
    """Altura da label em `cel` (coluna A): maior entre a regra por
    caracteres (piso) e a quebra simulada por largura real."""
    altura_regra = altura_label_por_caracteres(cel.value)
    tamanho_fonte = float(cel.font.size or 10) if cel.font else 10.0
    largura_coluna = largura_total_mesclagem(ws, cel)
    # (o recuo das labels já está embutido na calibração, feita com
    # relatório real)
    largura_util = (
        LARGURA_UTIL_LABEL_COLUNA_A * (largura_coluna / LARGURA_REFERENCIA_COLUNA_A)
    ) * (10.0 / tamanho_fonte)
    altura_largura = linhas_por_largura(cel.value, largura_util) * ALTURA_POR_LINHA_LABEL
    return max(altura_regra, altura_largura)


# Textos que, onde aparecerem (qualquer coluna), fixam a altura da linha
# inteira — com prioridade sobre qualquer outra regra de altura do
# código 09. Comparação sem acento, sem diferença de maiúsculas e com
# espaços/quebras de linha normalizados ("Denominações\nCatólicas" e
# "DENOMINACOES CATOLICAS" também valem). Pedido do Lucas: 15.
ALTURAS_FIXAS_POR_TEXTO = {
    "denominacoes catolicas": 15,
    "denominacoes evangelicas": 15,
}


def _normalizar_para_comparacao(texto):
    import unicodedata
    sem_acento = "".join(
        ch for ch in unicodedata.normalize("NFKD", str(texto))
        if not unicodedata.combining(ch)
    )
    return " ".join(sem_acento.lower().split())


def altura_fixa_por_texto(ws, r, max_col):
    """Se a linha `r` contiver algum texto de `ALTURAS_FIXAS_POR_TEXTO`,
    retorna a altura fixa correspondente; senão, None."""
    for c in range(1, max_col + 1):
        valor = ws.cell(row=r, column=c).value
        if not isinstance(valor, str):
            continue
        normalizado = _normalizar_para_comparacao(valor)
        for termo, altura in ALTURAS_FIXAS_POR_TEXTO.items():
            if termo in normalizado:
                return altura
    return None


def linhas_com_mesclagem_na_coluna_a(ws):
    """Linhas em que a coluna A está numa mesclagem que NÃO é de label.

    Label de tabela pode vir mesclada em A:B numa linha só (comum em
    tabelas estreitas, como as do Total Automático) — essa continua
    valendo como label. Qualquer outra mesclagem que pegue a coluna A
    (mais de uma linha, ou indo além da coluna B — títulos, perguntas,
    legendas e notas, mescladas na largura da tabela) tira a linha da
    regra de label.

    Calculado uma vez só por planilha (varrer `merged_cells` célula a
    célula é caro — ver learnings de performance)."""
    linhas = set()
    for rng in ws.merged_cells.ranges:
        if not (rng.min_col <= 1 <= rng.max_col):
            continue
        mesclagem_de_label = rng.min_row == rng.max_row and rng.max_col <= 2
        if not mesclagem_de_label:
            linhas.update(range(rng.min_row, rng.max_row + 1))
    return linhas


def eh_label_de_tabela_coluna_a(ws, r, max_col, linhas_mescladas_a):
    """True se a célula A{r} é uma label de tabela (não título/pergunta).

    Critérios:
    - texto não vazio na coluna A;
    - não começa com 'Pergunta:' / 'Titulo:' / 'Título:';
    - A{r} não está numa mesclagem "de título" (várias linhas, ou além da
      coluna B) — ver `linhas_com_mesclagem_na_coluna_a`. Mesclagem A:B
      numa linha só é aceita como label;
    - a linha tem algum valor a partir da coluna B — ou seja, é uma linha
      DE TABELA (inclui tabelas só com Total, onde o dado pode estar na
      coluna B). Isso deixa de fora legendas/notas soltas na coluna A
      (a etapa de Legendas roda antes do código 09).
    """
    valor = ws.cell(row=r, column=1).value
    if not isinstance(valor, str) or not valor.strip():
        return False
    if valor.strip().lower().startswith(_PREFIXOS_NAO_LABEL):
        return False
    if r in linhas_mescladas_a:
        return False
    for c in range(2, max_col + 1):
        v = ws.cell(row=r, column=c).value
        if v is not None and v != "":
            return True
    return False


def estimar_altura_calculada(texto, largura_total_colunas, tamanho_fonte):
    """
    Estimativa (aproximada) da altura que uma célula com quebra de linha
    precisaria para caber todo o texto, dada a largura total das colunas
    envolvidas e o tamanho da fonte.

    IMPORTANTE: isto é uma APROXIMAÇÃO por contagem de caracteres — não é
    o cálculo real do Excel (que depende da métrica exata da fonte, algo
    que só o próprio Excel consegue calcular ao renderizar). Serve como
    um bom ponto de partida; ajustes finos de altura continuam sendo
    melhor feitos no Excel (duplo clique na borda da linha).

    Calibrado com margem de segurança pra fontes largas como a "DIN" (a
    fonte usada nos rótulos do relatório) — a estimativa original (fator
    de caractere 0.52, sem margem) subestimava o número de linhas em
    textos como "Não aprova nenhuma área" numa coluna de 20.67 de
    largura: previa 1 linha quando na real precisava de 2, cortando o
    texto visualmente. O fator de caractere maior (0.62) e a margem de
    10% a menos na largura útil (0.90) tornam a estimativa mais
    conservadora — prefere estourar um pouco a altura a arriscar cortar
    texto, o que é bem pior visualmente.
    """
    if not texto:
        return tamanho_fonte * 1.6
    n_linhas = estimar_numero_linhas(texto, largura_total_colunas, tamanho_fonte)
    altura_por_linha = tamanho_fonte * 1.6
    return n_linhas * altura_por_linha


def largura_total_mesclagem(ws, celula):
    """Largura total (em unidades de coluna do Excel) da célula ou da sua faixa mesclada."""
    faixa = None
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= celula.row <= rng.max_row and rng.min_col <= celula.column <= rng.max_col:
            faixa = rng
            break

    if faixa is None:
        colunas = [celula.column]
    else:
        colunas = range(faixa.min_col, faixa.max_col + 1)

    total = 0.0
    largura_padrao = ws.sheet_format.defaultColWidth or 8.43
    for c in colunas:
        letra = get_column_letter(c)
        dim = ws.column_dimensions.get(letra)
        total += dim.width if (dim and dim.width) else largura_padrao
    return total


def worksheet_para_dataframe(ws, max_linhas=300, max_colunas=40):
    """
    Converte uma planilha num pandas.DataFrame só de VALORES (sem
    formatação — fonte, cor, borda etc. não têm como aparecer numa
    pré-visualização em tabela HTML/Streamlit), pra uso em pré-
    visualização dentro do app antes do download.

    Células dentro de uma faixa mesclada (fora a célula-âncora, que já
    tem o valor) recebem o mesmo valor da âncora — sem isso, a pré-
    visualização mostraria um monte de células vazias onde na verdade
    o Excel mostra um bloco mesclado com texto, o que confundiria mais
    do que ajudaria a pessoa a validar o resultado.

    Trunca em `max_linhas`/`max_colunas` por performance — planilhas de
    relatório real passam de 3000 linhas, e renderizar tudo isso como
    HTML deixaria a pré-visualização lenta sem ganho real (a pessoa
    consegue ver o começo e decidir se baixa o arquivo completo).

    Returns:
        (df, truncado_linhas, truncado_colunas) — o DataFrame já com as
        mesclagens preenchidas, e dois booleanos indicando se cada eixo
        foi cortado (pra o chamador avisar a pessoa, se for o caso).
    """
    total_linhas = ws.max_row
    total_colunas = ws.max_column

    limite_linhas = min(total_linhas, max_linhas)
    limite_colunas = min(total_colunas, max_colunas)

    linhas = []
    for r in range(1, limite_linhas + 1):
        linha = [ws.cell(row=r, column=c).value for c in range(1, limite_colunas + 1)]
        linhas.append(linha)

    colunas_nomes = [get_column_letter(c) for c in range(1, limite_colunas + 1)]
    df = pd.DataFrame(linhas, columns=colunas_nomes)

    # Preenche as mesclagens que caem dentro da área exibida
    for rng in ws.merged_cells.ranges:
        if rng.min_row > limite_linhas or rng.min_col > limite_colunas:
            continue
        valor_ancora = ws.cell(row=rng.min_row, column=rng.min_col).value
        if valor_ancora is None:
            continue
        max_r = min(rng.max_row, limite_linhas)
        max_c = min(rng.max_col, limite_colunas)
        for r in range(rng.min_row, max_r + 1):
            for c in range(rng.min_col, max_c + 1):
                df.iat[r - 1, c - 1] = valor_ancora

    truncado_linhas = total_linhas > max_linhas
    truncado_colunas = total_colunas > max_colunas
    return df, truncado_linhas, truncado_colunas


# ---------------------------------------------------------------------------
# Pré-visualização "de verdade" (HTML que imita o Excel: mesclagem, cor de
# fundo, negrito/itálico, bordas, alinhamento) — worksheet_para_dataframe()
# acima vira uma tabela genérica sem nada disso, o que fica muito diferente
# do arquivo real. Esta função gera HTML pra renderizar num componente
# (st.components.v1.html), não um DataFrame.
# ---------------------------------------------------------------------------
_MAPA_BORDA_PX = {
    "thin": 1, "hair": 1, "dotted": 1, "dashed": 1,
    "medium": 2, "mediumDashed": 2, "slantDashDot": 2,
    "thick": 3, "double": 3,
}


def _cor_hex(cor_openpyxl, padrao=None):
    """Extrai um #RRGGBB de um objeto Color do openpyxl, se der pra saber
    a cor real (RGB explícito) — cores indexadas/de tema não têm mapeamento
    simples fora do Excel, então caem no padrão."""
    if cor_openpyxl is None:
        return padrao
    try:
        if cor_openpyxl.type == "rgb" and cor_openpyxl.rgb and len(cor_openpyxl.rgb) == 8:
            rgb = cor_openpyxl.rgb[2:]  # descarta o alpha (AARRGGBB -> RRGGBB)
            if rgb != "000000" or padrao is None:
                return f"#{rgb}"
    except AttributeError:
        pass
    return padrao


def _estilo_borda_css(borda):
    """Monta o CSS de borda (um lado por vez) a partir de um Border do openpyxl."""
    partes = {}
    for lado_nome, lado in (
        ("top", borda.top), ("bottom", borda.bottom),
        ("left", borda.left), ("right", borda.right),
    ):
        if lado is not None and lado.style:
            espessura = _MAPA_BORDA_PX.get(lado.style, 1)
            cor = _cor_hex(lado.color, "#000000")
            estilo = "double" if lado.style == "double" else "solid"
            partes[lado_nome] = f"border-{lado_nome}: {espessura}px {estilo} {cor};"
    return "".join(partes.values())


def _estilo_celula_css(cel):
    """Monta o CSS inline de uma célula (fonte, preenchimento, alinhamento, borda)."""
    partes = []

    fonte = cel.font
    if fonte:
        if fonte.bold:
            partes.append("font-weight:bold;")
        if fonte.italic:
            partes.append("font-style:italic;")
        if fonte.size:
            partes.append(f"font-size:{fonte.size * 1.05:.0f}px;")
        cor_fonte = _cor_hex(fonte.color)
        if cor_fonte:
            partes.append(f"color:{cor_fonte};")

    fill = cel.fill
    if fill and fill.patternType == "solid":
        cor_fundo = _cor_hex(fill.fgColor)
        if cor_fundo:
            partes.append(f"background-color:{cor_fundo};")

    alinh = cel.alignment
    if alinh:
        mapa_h = {"left": "left", "center": "center", "right": "right", "general": None}
        h = mapa_h.get(alinh.horizontal) if alinh.horizontal else None
        if h:
            partes.append(f"text-align:{h};")
        mapa_v = {"top": "top", "center": "middle", "bottom": "bottom"}
        v = mapa_v.get(alinh.vertical) if alinh.vertical else None
        if v:
            partes.append(f"vertical-align:{v};")
        if alinh.wrap_text:
            partes.append("white-space:normal;word-break:break-word;")
        else:
            partes.append("white-space:nowrap;")

    partes.append(_estilo_borda_css(cel.border))
    return "".join(partes)


def _formatar_valor_preview(valor):
    """
    Formata o valor de uma célula pra exibição na pré-visualização: número
    com casa decimal (float) mostra só 1 casa — a planilha real guarda a
    precisão completa, isso é só pra não poluir a leitura na tela (ex.:
    "38.58568887059666" vira "38.6"). Inteiros e texto ficam como estão.
    """
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return str(valor)
    if isinstance(valor, float):
        return f"{valor:.1f}"
    return str(valor)


def worksheet_para_html_paginado(ws):
    """
    Como `worksheet_para_html`, mas sem truncar linhas/colunas (o
    relatório inteiro) e dividido em UMA TABELA HTML POR PÁGINA (nas
    mesmas linhas onde a planilha tem uma quebra de página manual),
    cada uma dentro de uma <div> com quebra de página — pensado pra
    virar um PDF paginado de verdade (ver core/pdf_preview.py).

    Importante: a quebra de página é aplicada na <div> que envolve cada
    tabela, NÃO num <tr> dentro de uma tabela única — "page-break-before"
    em linha de tabela não é respeitado de forma confiável por praticamente
    nenhum motor de renderização HTML->PDF (é uma limitação conhecida de
    CSS, não só do wkhtmltopdf), por isso cada página vira sua própria
    tabela HTML independente.

    Returns:
        str de HTML completo (com <html>/<head>/<body>, já pronto pra
        passar direto pro conversor de PDF).
    """
    total_linhas = ws.max_row
    total_colunas = ws.max_column

    celulas_cobertas = set()
    spans = {}
    for rng in ws.merged_cells.ranges:
        spans[(rng.min_row, rng.min_col)] = (
            rng.max_row - rng.min_row + 1, rng.max_col - rng.min_col + 1
        )
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    celulas_cobertas.add((r, c))

    larguras_px = []
    largura_padrao = ws.sheet_format.defaultColWidth or 8.43
    for c in range(1, total_colunas + 1):
        letra = get_column_letter(c)
        dim = ws.column_dimensions.get(letra)
        largura = dim.width if (dim and dim.width) else largura_padrao
        larguras_px.append(max(20, round(largura * 6.2)))
    colgroup = "".join(f'<col style="width:{w}px;">' for w in larguras_px)

    # pontos onde uma nova página começa (linha 1-based) — a partir das
    # quebras de página manuais já presentes na planilha
    inicios_pagina = sorted({b.id + 1 for b in ws.row_breaks.brk if b.id is not None})
    limites = [1] + [i for i in inicios_pagina if i > 1] + [total_linhas + 1]
    limites = sorted(set(limites))

    def _renderizar_linha(r):
        celulas_html = []
        for c in range(1, total_colunas + 1):
            if (r, c) in celulas_cobertas:
                continue
            cel = ws.cell(row=r, column=c)
            texto = _formatar_valor_preview(cel.value)
            texto = (
                texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\n", "<br>")
            )
            attrs = ""
            span_r, span_c = spans.get((r, c), (1, 1))
            if span_r > 1:
                attrs += f' rowspan="{span_r}"'
            if span_c > 1:
                attrs += f' colspan="{span_c}"'
            estilo = _estilo_celula_css(cel)
            celulas_html.append(f'<td style="{estilo}"{attrs}>{texto}</td>')
        dim_linha = ws.row_dimensions.get(r)
        altura_px = round((dim_linha.height or 15) * 1.1) if dim_linha else 17
        return f'<tr style="height:{altura_px}px;">{"".join(celulas_html)}</tr>'

    paginas_html = []
    for i in range(len(limites) - 1):
        inicio_pag, fim_pag = limites[i], limites[i + 1] - 1
        linhas_da_pagina = [
            _renderizar_linha(r) for r in range(inicio_pag, fim_pag + 1)
            if any((r, c) not in celulas_cobertas for c in range(1, total_colunas + 1))
            or True  # inclui mesmo linhas "vazias" (mantém espaçamento do layout original)
        ]
        if not any(ws.cell(row=r, column=c).value not in (None, "") for r in range(inicio_pag, fim_pag + 1) for c in range(1, total_colunas + 1)):
            # página inteiramente vazia (ex.: páginas em branco antes de uma capa) —
            # ainda assim gera uma <div> própria, só que sem tabela dentro,
            # pra manter a contagem/posição de páginas do PDF fiel ao original
            paginas_html.append('<div class="pagina"></div>')
            continue
        tabela = (
            f'<table><colgroup>{colgroup}</colgroup>'
            f'<tbody>{"".join(linhas_da_pagina)}</tbody></table>'
        )
        paginas_html.append(f'<div class="pagina">{tabela}</div>')

    orientacao = (ws.page_setup.orientation or "portrait") if ws.page_setup else "portrait"

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page {{ size: A4 {orientacao}; margin: 10mm; }}
  body {{ margin: 0; }}
  .pagina {{ page-break-after: always; }}
  .pagina:last-child {{ page-break-after: avoid; }}
  table {{ border-collapse: collapse; font-family: Calibri, Arial, sans-serif;
           font-size: 10px; table-layout: fixed; width: 100%; }}
  td {{ overflow: hidden; }}
</style>
</head>
<body>
  {"".join(paginas_html)}
</body>
</html>"""


def worksheet_para_html(ws, max_linhas=150, max_colunas=30):
    """
    Gera uma tabela HTML que imita a aparência real da planilha no Excel:
    mesclagem de células (colspan/rowspan), cor de fundo, negrito/itálico,
    tamanho e cor de fonte, alinhamento, quebra de texto e bordas — tudo
    lido direto dos estilos do openpyxl.

    Diferente de `worksheet_para_dataframe`, isso é HTML pronto pra
    renderizar (ex.: com `st.components.v1.html`), não um DataFrame — não
    dá pra ordenar/filtrar como uma tabela de dados, é uma pré-visualização
    visual mesmo.

    Trunca em `max_linhas`/`max_colunas` por performance (ver docstring de
    `worksheet_para_dataframe` — mesmo motivo).

    Returns:
        (html, truncado_linhas, truncado_colunas)
    """
    total_linhas = ws.max_row
    total_colunas = ws.max_column
    limite_linhas = min(total_linhas, max_linhas)
    limite_colunas = min(total_colunas, max_colunas)

    # Mapeia cada célula coberta por uma mesclagem (exceto a âncora) pra
    # não renderizá-la — e guarda o rowspan/colspan da âncora.
    celulas_cobertas = set()
    spans = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row > limite_linhas or rng.min_col > limite_colunas:
            continue
        max_r = min(rng.max_row, limite_linhas)
        max_c = min(rng.max_col, limite_colunas)
        spans[(rng.min_row, rng.min_col)] = (max_r - rng.min_row + 1, max_c - rng.min_col + 1)
        for r in range(rng.min_row, max_r + 1):
            for c in range(rng.min_col, max_c + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    celulas_cobertas.add((r, c))

    # Larguras de coluna (aproximação: unidade de largura do Excel -> px)
    larguras_px = []
    largura_padrao = ws.sheet_format.defaultColWidth or 8.43
    for c in range(1, limite_colunas + 1):
        letra = get_column_letter(c)
        dim = ws.column_dimensions.get(letra)
        largura = dim.width if (dim and dim.width) else largura_padrao
        larguras_px.append(max(24, round(largura * 7 + 5)))

    linhas_html = []
    for r in range(1, limite_linhas + 1):
        dim_linha = ws.row_dimensions.get(r)
        altura_px = round((dim_linha.height or 15) * 1.33) if dim_linha else 20
        celulas_html = []
        for c in range(1, limite_colunas + 1):
            if (r, c) in celulas_cobertas:
                continue
            cel = ws.cell(row=r, column=c)
            texto = _formatar_valor_preview(cel.value)
            texto = (
                texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\n", "<br>")
            )
            attrs = ""
            span_r, span_c = spans.get((r, c), (1, 1))
            if span_r > 1:
                attrs += f' rowspan="{span_r}"'
            if span_c > 1:
                attrs += f' colspan="{span_c}"'
            estilo = _estilo_celula_css(cel)
            celulas_html.append(f'<td style="{estilo}"{attrs}>{texto}</td>')
        linhas_html.append(f'<tr style="height:{altura_px}px;">{"".join(celulas_html)}</tr>')

    colgroup = "".join(f'<col style="width:{w}px;">' for w in larguras_px)

    html = f"""
    <div style="overflow:auto; max-height:600px; border:1px solid #d1d5db;">
      <table style="border-collapse:collapse; font-family:Calibri,Arial,sans-serif;
                     font-size:13px; table-layout:fixed;">
        <colgroup>{colgroup}</colgroup>
        <tbody>{"".join(linhas_html)}</tbody>
      </table>
    </div>
    """

    truncado_linhas = total_linhas > max_linhas
    truncado_colunas = total_colunas > max_colunas
    return html, truncado_linhas, truncado_colunas


# ---------------------------------------------------------------------------
# Detecção de blocos de tabela e exclusão de colunas por "Base reduzida"
# abaixo de um limite — usado tanto pelo código 05 do Relatório
# Automatizado quanto pelo Processador de Base Reduzida standalone (as
# duas telas compartilham este mesmo motor, testado uma vez só).
# ---------------------------------------------------------------------------
def normalizar_texto_maiusculo(valor):
    return str(valor).strip().upper() if valor is not None else ""


def linha_vazia_ate_coluna(ws, r, max_col):
    return all(ws.cell(row=r, column=c).value is None for c in range(1, max_col + 1))


def encontrar_inicio_bloco(ws, linha_ref, max_col):
    """Sobe a partir de `linha_ref` enquanto a linha anterior não estiver
    totalmente vazia — acha o início do bloco/tabela contíguo."""
    r = linha_ref
    while r > 1 and not linha_vazia_ate_coluna(ws, r - 1, max_col):
        r -= 1
    return r


def encontrar_fim_bloco(ws, linha_ref, max_col):
    """Desce a partir de `linha_ref` enquanto a linha seguinte não estiver
    totalmente vazia — acha o fim do bloco/tabela contíguo."""
    r = linha_ref
    max_row = ws.max_row
    while r < max_row and not linha_vazia_ate_coluna(ws, r + 1, max_col):
        r += 1
    return r


def _excluir_coluna_no_bloco(ws, linha_inicio, linha_fim, coluna, max_col):
    """Remove uma coluna dentro de [linha_inicio, linha_fim], deslocando
    tudo à direita dela uma posição pra esquerda (valor + estilo), e
    limpa a última coluna do bloco (que sobra "duplicada" depois do
    deslocamento)."""
    from copy import copy as _copy_style
    from openpyxl.styles import Border

    for r in range(linha_inicio, linha_fim + 1):
        for c in range(coluna, max_col):
            origem = ws.cell(row=r, column=c + 1)
            destino = ws.cell(row=r, column=c)
            destino.value = origem.value
            if origem.has_style:
                destino.font = _copy_style(origem.font)
                destino.border = _copy_style(origem.border)
                destino.fill = _copy_style(origem.fill)
                destino.number_format = origem.number_format
                destino.protection = _copy_style(origem.protection)
                destino.alignment = _copy_style(origem.alignment)
        ultima = ws.cell(row=r, column=max_col)
        ultima.value = None
        ultima.border = Border()


def excluir_colunas_base_reduzida(ws, limite=25, texto_alvo="BASE REDUZIDA"):
    """
    Para toda linha cuja coluna A seja exatamente `texto_alvo` (comparação
    exata, sem diferenciar maiúsc./minúsc.), exclui as colunas desse
    bloco de tabela cujo valor numérico seja menor que `limite` —
    deslocando o resto pra esquerda e reconstruindo corretamente
    qualquer mesclagem de cabeçalho que abranja essas colunas (ex.: um
    título "Regiões" mesclado sobre várias colunas de região encolhe
    para cobrir só as colunas que sobraram, em vez de quebrar).

    Detecta os limites do bloco/tabela automaticamente (sobe/desce até
    achar uma linha totalmente vazia), então funciona tanto num
    relatório grande com várias tabelas quanto num arquivo com uma
    tabela só.

    Returns:
        Número de colunas excluídas no total (soma de todos os blocos).
    """
    max_col_geral = ws.max_column
    linhas_alvo = [
        r for r in range(1, ws.max_row + 1)
        if normalizar_texto_maiusculo(ws.cell(row=r, column=1).value) == texto_alvo
    ]

    colunas_excluidas_total = 0

    for linha_base in linhas_alvo:
        inicio = encontrar_inicio_bloco(ws, linha_base, max_col_geral)
        fim = encontrar_fim_bloco(ws, linha_base, max_col_geral)

        ultima_col_linha = 1
        for c in range(1, max_col_geral + 1):
            if ws.cell(row=linha_base, column=c).value is not None:
                ultima_col_linha = c

        colunas_para_excluir = []
        for c in range(ultima_col_linha, 1, -1):
            valor = ws.cell(row=linha_base, column=c).value
            numero = None
            if isinstance(valor, (int, float)):
                numero = valor
            elif isinstance(valor, str):
                limpo = "".join(ch for ch in valor.replace(",", ".") if ch.isdigit() or ch in ".-")
                if limpo:
                    try:
                        numero = float(limpo)
                    except ValueError:
                        numero = None
            if numero is not None and numero < limite:
                colunas_para_excluir.append(c)

        if not colunas_para_excluir:
            continue

        faixas_capturadas = []
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row <= fim and rng.max_row >= inicio:
                # guarda o valor de verdade (só a célula do canto superior
                # esquerdo de uma mesclagem tem o texto) para restaurar
                # depois — sem isso, o deslocamento de coluna pode
                # sobrescrever esse valor com uma célula vizinha vazia
                valor_original = ws.cell(row=rng.min_row, column=rng.min_col).value
                faixas_capturadas.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col, valor_original))
                try:
                    ws.unmerge_cells(str(rng))
                except KeyError:
                    ws.merged_cells.ranges.discard(rng)

        for col_excluir in colunas_para_excluir:  # já em ordem decrescente
            _excluir_coluna_no_bloco(ws, inicio, fim, col_excluir, max_col_geral)
            colunas_excluidas_total += 1

        for min_row, min_col, max_row, max_col, valor_original in faixas_capturadas:
            excluidas_antes = sum(1 for c in colunas_para_excluir if c < min_col)
            excluidas_dentro = sum(1 for c in colunas_para_excluir if min_col <= c <= max_col)
            novo_min_col = min_col - excluidas_antes
            novo_max_col = max_col - excluidas_antes - excluidas_dentro
            if novo_max_col >= novo_min_col:
                ws.cell(row=min_row, column=novo_min_col).value = valor_original
                # remescla se sobrou mais de 1 COLUNA (novo_max_col >
                # novo_min_col) OU se a mesclagem já era vertical desde o
                # início (max_row > min_row, ex.: "Total" ocupando as 2
                # linhas de cabeçalho) — checar só a largura ignorava
                # mesclagens puramente verticais, perdendo elas mesmo
                # quando nenhuma coluna daquele intervalo foi excluída
                if novo_max_col > novo_min_col or max_row > min_row:
                    ws.merge_cells(
                        start_row=min_row, start_column=novo_min_col,
                        end_row=max_row, end_column=novo_max_col
                    )

    return colunas_excluidas_total
