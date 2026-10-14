"""
Códigos do "Relatório Automatizado" portados de VBA para Python/openpyxl.

Cada função aqui corresponde a um dos códigos originais (numeração igual
à da conversa de origem). Onde o código VBA original depende do AutoFit
real do Excel (algo que não tem como ser replicado com exatidão fora do
Excel, já que depende do motor de renderização de fonte), a altura é
estimada por contagem de caracteres — ver `core.planilha_utils` — e isso
fica marcado explicitamente no retorno de cada função, para o app avisar
a pessoa.

Sem dependência do Streamlit — pode ser testado isoladamente.
"""
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from core.planilha_utils import (
    estimar_altura_calculada,
    estimar_numero_linhas,
    altura_label_por_caracteres,
    eh_label_de_tabela_coluna_a,
    linhas_com_mesclagem_na_coluna_a,
    altura_fixa_por_texto,
    altura_label_celula,
    ALTURA_POR_LINHA_LABEL,
    inserir_linhas_seguro,
    largura_total_mesclagem,
    remover_linhas_seguro,
    _desfazer_mesclagem_rapido,
    _remesclar_rapido,
)

THICK = Side(style="thick")
MEDIUM = Side(style="medium")


def _remover_prefixo_case_insensitive(texto, prefixo):
    if texto[:len(prefixo)].lower() == prefixo.lower():
        return texto[len(prefixo):]
    return texto


# =========================================================
# Código 04 — Formatação de rótulos específicos (quebra de linha manual)
# =========================================================
# Cada regra: texto EXATO a buscar -> texto formatado (com \n),
# altura de linha (None = não mexe), largura de coluna (None = não mexe).
# Lista extraída fielmente do VBA original — específica de projeto, então
# é normal precisar adicionar/editar linhas aqui conforme o projeto mudar.
# Cada regra: (texto original exato, texto formatado com quebras de
# linha, altura, largura, coluna_minima). coluna_minima=None = vale em
# qualquer coluna (rótulo normal, coluna A); coluna_minima=N = só vale
# a partir dessa coluna (cabeçalho de segmento — Escolaridade e
# Religiões só devem quebrar/ganhar altura 45 a partir da coluna C,
# nunca quando o mesmo texto aparecer como rótulo de linha comum).
REGRAS_CODIGO_04_EXATAS = [
    # ---- Não restritas a coluna (rótulos de linha, coluna A) ----
    ("Aluno em domicílio (Educação Pública)", " Aluno em domicílio \n(Educação Pública)", 35, 15, None),
    ("Uso recente do Sistema de Saúde Pública", "Uso recente do Sistema \nde Saúde Pública", 35, 15, None),
    ("Participação em Tratamento Contínuo da Prefeitura",
     " Participação em Tratamento \nContínuo da Prefeitura ", 35, 15, None),
    ("Benefício da Cidade recebido por alguém no domicílio",
     "Benefício da Cidade recebido \n por alguém no domicílio", 35, 15, None),
    ("Renda Média Mensal Domiciliar (em Salário Minímo)",
     " Renda Média \n Mensal Domiciliar\n(em Salário Minímo)", 60, 15, None),

    # ---- Escolaridade — altura 45, a partir da coluna C ----
    ("Ensino Fundamental", "Ensino\nFunda-\nmental", 45, 15, 3),
    ("Ensino Funda- mental", "Ensino\nFunda-\nmental", 45, 15, 3),
    ("Ensino Médio", "Ensino\nMédio", 45, 15, 3),
    ("Ensino Superior", "Ensino\nSuperior", 45, 15, 3),

    # ---- Religiões — altura 45, a partir da coluna C ----
    ("Católica Total", "Católica\nTotal", 45, 15, 3),
    ("Católica de Batismo", "Católica de\nBatismo", 45, 15, 3),
    (" Católica de Batismo", " Católica de\nBatismo", 45, 15, 3),
    ("Católica Praticante", "Católica\nPraticante", 45, 15, 3),
    ("Evangélica Total", "Evangé-\nlica\nTotal", 45, 15, 3),
    ("Evangélicas de Missão", "Evangé-\nlicas de\nMissão", 45, 15, 3),
    ("Pente- costal", "Pente-\ncostal", 45, 15, 3),
    ("Pentecostal", "Pente-\ncostal", 45, 15, 3),
    ("Espírita/Kardecista", "Espírita/\nKarde-\ncista", 45, 15, 3),
    ("Afro-Brasileiras", "Afro-\nBrasi-\nleiras", 45, 15, 3),
    ("Testemunha de Jeová", "Teste-\nmunha\nde Jeová", 45, 15, 3),
    ("Não sabe", "Não\nsabe", 45, 15, 3),
    ("Outras religiões", "Outras\nreligiões", 45, 15, 3),

    # ---- Voto no 2º turno em 2022 — altura 30, a partir da coluna C ----
    ("Jair Bolsonaro", "Jair\nBolsonaro", 30, 15, 3),

    # ---- Avaliação — altura 30, a partir da coluna C ----
    ("Aprovação", "Apro-\nvação", 30, 15, 3),
    ("Reprovação", "Repro-\nvação", 30, 15, 3),
]

# Regras de "Avaliação do Governo X": coluna 1 mantém o texto original
# (altura 18); demais colunas quebram em 2 linhas (altura 34).
GOVERNADORES_CODIGO_04 = [
    "Avaliação do Governo Tarcísio de Freitas",
    "Avaliação do Governo Romeu Zema",
    "Avaliação do Governo Eduardo Riedel",
    "Avaliação do Governo Cláudio Castro",
    "Avaliação do Governo Renato Casagrande",
    "Avaliação do Governo do ex-Presidente Bolsonaro",
]

# Prefixos que, quando a célula COMEÇA com um deles, recebem quebra de
# linha logo depois do prefixo (mantendo o resto do texto na linha de baixo)
PREFIXOS_CODIGO_04 = [
    "Avaliação do Governo Municipal",
    "Avaliação da gestão do",
    "Avaliação da gestão da",
    "Avaliação da gestão de w",
    "Avaliação da gestão atual com",
]


def aplicar_codigo_04(ws):
    """
    Aplica quebra de linha manual e ajustes de altura/largura em rótulos
    específicos (ver REGRAS_CODIGO_04_EXATAS, GOVERNADORES_CODIGO_04,
    PREFIXOS_CODIGO_04) — lista de correções fixas usada há tempos nesse
    tipo de relatório. "Denominações Católicas"/"Denominações
    Evangélicas" também são mescladas com a célula à direita, como no
    VBA original.
    """
    alterados = 0
    centralizado = Alignment(horizontal="center", vertical="center", wrap_text=True)
    max_col = ws.max_column

    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            texto = cel.value
            if not isinstance(texto, str):
                continue

            if texto in ("Denominações Católicas", "Denominações Evangélicas"):
                ja_mesclada = any(
                    rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col
                    for rng in ws.merged_cells.ranges
                )
                if not ja_mesclada:
                    # O grupo pode se espalhar por MAIS de 2 colunas (ex.:
                    # "Denominações Católicas" cobrindo 3 subcolunas —
                    # Total/Batismo/Praticante — contra só 2 de
                    # "Denominações Evangélicas") — sempre mesclar com
                    # "a célula vizinha" (largura fixa de 2) deixava a
                    # 3ª coluna de fora, com o mesmo texto repetido numa
                    # célula separada, sem mesclar. Acha até onde a
                    # repetição do mesmo texto vai antes de mesclar —
                    # usando o valor EFETIVO de cada coluna (se ela já
                    # for filha de uma mesclagem parcial pré-existente,
                    # `.value` sozinho devolve None; sem "enxergar
                    # através" dessa mesclagem, a detecção parava cedo
                    # demais e a mesclagem saía incompleta mesmo assim).
                    def _valor_efetivo(col):
                        v = ws.cell(row=r, column=col).value
                        if v is not None:
                            return v
                        for rng in ws.merged_cells.ranges:
                            if rng.min_row <= r <= rng.max_row and rng.min_col <= col <= rng.max_col:
                                return ws.cell(row=rng.min_row, column=rng.min_col).value
                        return None

                    c_fim = c
                    while c_fim + 1 <= max_col and _valor_efetivo(c_fim + 1) == texto:
                        c_fim += 1
                    # Desfaz qualquer mesclagem que já exista dentro desse
                    # intervalo antes de criar a nova — evita sobreposição
                    # inválida (célula pertencendo a duas mesclagens ao
                    # mesmo tempo), que o Excel descarta silenciosamente
                    # ao abrir o arquivo, deixando a mesclagem pela
                    # metade — foi exatamente essa a causa da mesclagem
                    # não estar acontecendo.
                    for rng in list(ws.merged_cells.ranges):
                        if rng.min_row <= r <= rng.max_row and not (rng.max_col < c or rng.min_col > c_fim):
                            try:
                                ws.unmerge_cells(str(rng))
                            except KeyError:
                                ws.merged_cells.ranges.discard(rng)
                    ws.cell(row=r, column=c, value=texto)
                    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c_fim)
                # aplica a formatação sempre, mesmo quando a célula já veio
                # mesclada de fábrica (o VBA original só formatava quando
                # TAMBÉM precisava mesclar — em arquivos que já chegam
                # pré-mesclados isso deixava a formatação de fora)
                cel.alignment = centralizado
                ws.row_dimensions[r].height = 18
                alterados += 1
                continue

            casou = False
            for original, formatado, altura, largura, coluna_minima in REGRAS_CODIGO_04_EXATAS:
                if texto == original and (coluna_minima is None or c >= coluna_minima):
                    cel.value = formatado
                    cel.alignment = centralizado
                    if altura is not None:
                        ws.row_dimensions[r].height = altura
                    if largura is not None:
                        ws.column_dimensions[get_column_letter(c)].width = largura
                    alterados += 1
                    casou = True
                    break
            if casou:
                continue

            for nome_governo in GOVERNADORES_CODIGO_04:
                if texto == nome_governo:
                    if c == 1:
                        ws.row_dimensions[r].height = 18
                    else:
                        resto = nome_governo[len("Avaliação do Governo"):].strip()
                        cel.value = "Avaliação do Governo\n" + resto
                        ws.row_dimensions[r].height = 34
                    cel.alignment = centralizado
                    alterados += 1
                    casou = True
                    break
            if casou:
                continue

            if texto == "Repro- vação":
                if c == 1:
                    cel.value = "Reprovação"
                    cel.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cel.value = "Repro-\nvação"
                    ws.row_dimensions[r].height = 34
                    cel.alignment = centralizado
                alterados += 1
                continue

            for prefixo in PREFIXOS_CODIGO_04:
                if texto.startswith(prefixo):
                    resto = texto[len(prefixo):]
                    if resto:
                        cel.value = prefixo + "\n" + resto
                        cel.alignment = Alignment(wrap_text=True)
                        ws.row_dimensions[r].height = 34
                        alterados += 1
                    break

    return alterados




# =========================================================
# Código 01 — Preenche células do SPSS pirata
# =========================================================
def aplicar_codigo_01(ws):
    """
    Para cada linha onde a coluna B é 'Total': se não está mesclada
    (tabela de 2 colunas, sem segmento), aplica borda espessa na célula
    A correspondente. Se está mesclada em 2 linhas (tabela com
    segmento), mescla as 2 células correspondentes na coluna A (se
    ambas vazias e ainda não mescladas) e aplica borda espessa ao redor.
    """
    alterados = 0
    for r in range(1, ws.max_row + 1):
        cel_total = ws.cell(row=r, column=2)
        valor = str(cel_total.value).strip() if cel_total.value is not None else ""
        if valor != "Total":
            continue

        faixa_mesclada = None
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= 2 <= rng.max_col:
                faixa_mesclada = rng
                break

        if faixa_mesclada is None:
            a1 = ws.cell(row=r, column=1)
            a1.border = Border(top=THICK, bottom=THICK, left=THICK, right=THICK)
            alterados += 1
        elif faixa_mesclada.max_row - faixa_mesclada.min_row + 1 == 2:
            linha1 = faixa_mesclada.min_row
            linha2 = linha1 + 1
            a1 = ws.cell(row=linha1, column=1)
            a2 = ws.cell(row=linha2, column=1)

            ja_mesclada_a = any(
                rng.min_row <= linha1 <= rng.max_row and rng.min_col <= 1 <= rng.max_col
                for rng in ws.merged_cells.ranges
            )
            v1 = str(a1.value).strip() if a1.value is not None else ""
            v2 = str(a2.value).strip() if a2.value is not None else ""

            if not ja_mesclada_a and v1 == "" and v2 == "":
                # aplica a borda ANTES de mesclar: o merge_cells troca o
                # objeto interno das células que não são a "mestre", então
                # qualquer alteração feita DEPOIS do merge nessas células
                # se perde
                a1.border = Border(top=THICK, left=THICK, right=THICK)
                a2.border = Border(bottom=THICK, left=THICK, right=THICK)
                ws.merge_cells(start_row=linha1, start_column=1, end_row=linha2, end_column=1)
                alterados += 1

    return alterados


# =========================================================
# Código 02 — Muda Bordas (espessa -> média)
# =========================================================
def aplicar_codigo_02(ws):
    """Converte toda borda 'espessa' (thick) da planilha para 'média' (medium)."""
    alterados = 0
    max_col = ws.max_column
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            b = cel.border
            if not any([b.left, b.top, b.right, b.bottom]):
                continue

            mudou = False

            def _rebaixar(lado):
                nonlocal mudou
                if lado is not None and lado.style == "thick":
                    mudou = True
                    return Side(style="medium", color=lado.color)
                return lado

            novo = Border(
                left=_rebaixar(b.left), right=_rebaixar(b.right),
                top=_rebaixar(b.top), bottom=_rebaixar(b.bottom),
            )
            if mudou:
                cel.border = novo
                alterados += 1

    return alterados


# =========================================================
# Código 17 — Bordas espessas no cabeçalho das tabelas segmentadas
# =========================================================
def _com_lados_espessos(borda, lados):
    """Cópia da borda com os `lados` ('left','right','top','bottom')
    trocados por espessa — preserva os demais lados e a cor existente."""
    def _lado(nome):
        atual = getattr(borda, nome)
        if nome not in lados:
            return atual
        cor = atual.color if (atual is not None and atual.color is not None) else None
        return Side(style="thick", color=cor)
    return Border(
        left=_lado("left"), right=_lado("right"),
        top=_lado("top"), bottom=_lado("bottom"),
        diagonal=borda.diagonal, diagonalUp=borda.diagonalUp,
        diagonalDown=borda.diagonalDown, outline=borda.outline,
        vertical=borda.vertical, horizontal=borda.horizontal,
    )


def _engrossar(ws, r, c, lados, mesclagem_de=None):
    """Engrossa os `lados` da célula (r, c).

    Se a célula faz parte de uma mesclagem e o lado está na BORDA EXTERNA
    dela, engrossa o mesmo lado também na célula âncora (canto superior
    esquerdo): ao carregar o arquivo, o openpyxl recopia as bordas
    externas da mesclagem a partir da âncora (`MergedCellRange.format`),
    e como o assistente recarrega o arquivo a cada etapa, sem isso a
    borda espessa da base/direita de uma célula mesclada voltava a ser
    fina na etapa seguinte."""
    cel = ws.cell(row=r, column=c)
    cel.border = _com_lados_espessos(cel.border, lados)
    rng = (mesclagem_de or {}).get((r, c))
    if rng is None:
        return
    lados_ancora = set()
    if "top" in lados and r == rng.min_row:
        lados_ancora.add("top")
    if "bottom" in lados and r == rng.max_row:
        lados_ancora.add("bottom")
    if "left" in lados and c == rng.min_col:
        lados_ancora.add("left")
    if "right" in lados and c == rng.max_col:
        lados_ancora.add("right")
    if lados_ancora and (r, c) != (rng.min_row, rng.min_col):
        ancora = ws.cell(row=rng.min_row, column=rng.min_col)
        ancora.border = _com_lados_espessos(ancora.border, lados_ancora)


def _afinar(ws, r, c, lado, mesclagem_de=None):
    """Deixa o `lado` da célula (r, c) com borda FINA (preserva a cor). Como
    em `_engrossar`, se o lado está na borda externa de uma mesclagem, a
    âncora também recebe a borda — senão o openpyxl a desfaz ao recarregar."""
    def _troca(cel):
        b = cel.border
        atual = getattr(b, lado)
        cor = atual.color if (atual is not None and atual.color is not None) else None
        lados = {n: getattr(b, n) for n in ("left", "right", "top", "bottom")}
        lados[lado] = Side(style="thin", color=cor)
        cel.border = Border(diagonal=b.diagonal, diagonalUp=b.diagonalUp, diagonalDown=b.diagonalDown,
                            outline=b.outline, vertical=b.vertical, horizontal=b.horizontal, **lados)
    _troca(ws.cell(row=r, column=c))
    rng = (mesclagem_de or {}).get((r, c))
    if rng is None or (r, c) == (rng.min_row, rng.min_col):
        return
    na_borda = {"left": c == rng.min_col, "right": c == rng.max_col,
                "top": r == rng.min_row, "bottom": r == rng.max_row}[lado]
    if na_borda:
        _troca(ws.cell(row=rng.min_row, column=rng.min_col))


def _mapa_mesclagens(ws):
    mapa = {}
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                mapa[(r, c)] = rng
    return mapa


def _engrossar_vizinho_se_tiver_borda(ws, r, c, lado, mesclagem_de=None):
    """O Excel guarda a borda compartilhada nas DUAS células vizinhas; se
    a vizinha já tem borda fina/média nesse lado, ela também vira
    espessa pra não brigar com a nossa. Se não tem borda, não mexe."""
    if r < 1 or c < 1:
        return
    cel = ws.cell(row=r, column=c)
    if getattr(cel.border, lado) is not None and getattr(cel.border, lado).style:
        _engrossar(ws, r, c, {lado}, mesclagem_de)


def _ultima_coluna_da_tabela(ws, linha_ini, linha_fim, max_col):
    """Última coluna do cabeçalho: a maior coluna com valor nas linhas
    do cabeçalho ou na primeira linha de dados logo abaixo, ou coberta
    por uma mesclagem do cabeçalho."""
    ultima = 2
    for r in range(linha_ini, linha_fim + 2):
        for c in range(max_col, 2, -1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                ultima = max(ultima, c)
                break
    for rng in ws.merged_cells.ranges:
        if rng.max_row >= linha_ini and rng.min_row <= linha_fim and rng.min_col >= 2:
            ultima = max(ultima, rng.max_col)
    return ultima


def aplicar_codigo_17(ws):
    """Bordas espessas (thick) no contorno do cabeçalho das tabelas segmentadas.

    Identifica cada cabeçalho pela célula MESCLADA na coluna B com o texto
    "Total": as linhas do cabeçalho são as linhas dessa mesclagem (topo e
    base do "Total"), e a largura vai da coluna A até a última coluna da
    tabela. Aplica borda espessa em:
      - linha de cima e linha de baixo do cabeçalho (A até a última coluna);
      - lateral esquerda (coluna A) e lateral direita (última coluna);
      - lateral esquerda da coluna "Total" (B).
    A lateral DIREITA da célula "Total" (divisa com o primeiro segmento) fica
    FINA. As divisórias internas entre as subcategorias continuam como estão.

    Precisa rodar DEPOIS do código 02 (que rebaixa toda borda espessa pra
    média) e das etapas que mudam a estrutura da tabela (05, 07).
    Retorna o número de cabeçalhos formatados.
    """
    max_col = ws.max_column
    faixas_total = []
    for rng in ws.merged_cells.ranges:
        if not (rng.min_col <= 2 <= rng.max_col):
            continue
        if rng.min_col != 2:
            continue
        v = ws.cell(row=rng.min_row, column=2).value
        if isinstance(v, str) and v.strip().lower() == "total":
            faixas_total.append(rng)

    mapa = _mapa_mesclagens(ws)
    formatados = 0
    for rng in sorted(faixas_total, key=lambda x: x.min_row):
        topo, base = rng.min_row, rng.max_row
        col_total_fim = rng.max_col
        ultima = _ultima_coluna_da_tabela(ws, topo, base, max_col)

        for c in range(1, ultima + 1):
            _engrossar(ws, topo, c, {"top"}, mapa)
            _engrossar(ws, base, c, {"bottom"}, mapa)
            _engrossar_vizinho_se_tiver_borda(ws, topo - 1, c, "bottom", mapa)
            _engrossar_vizinho_se_tiver_borda(ws, base + 1, c, "top", mapa)

        for r in range(topo, base + 1):
            _engrossar(ws, r, 1, {"left"}, mapa)
            _engrossar(ws, r, ultima, {"right"}, mapa)
            # lateral esquerda da coluna "Total" (divisa com a coluna A)
            _engrossar(ws, r, 2, {"left"}, mapa)
            _engrossar(ws, r, 1, {"right"}, mapa)
            # lateral direita do "Total" (divisa com o 1º segmento): fina
            if col_total_fim + 1 <= ultima:
                _afinar(ws, r, col_total_fim, "right", mapa)
                _afinar(ws, r, col_total_fim + 1, "left", mapa)
            _engrossar_vizinho_se_tiver_borda(ws, r, ultima + 1, "left", mapa)
        formatados += 1
    return formatados


# =========================================================
# Código 03 — Ordenar tabelas (por valor da coluna B, descendente)
# =========================================================
# Lista de rótulos que NUNCA entram na ordenação — ficam fixos na posição
# original (linha) onde já estão. É extensa e específica de projeto
# (categorias de resposta padrão, "Não sabe", "Branco/Nulo", opções fixas
# de escala etc.) — o VBA original tinha essa lista hardcoded; aqui vira
# parâmetro com esse valor como padrão, editável pela tela de Códigos
# Individuais / Relatório Automatizado sem precisar mexer no código.
# Comparação sempre case-insensitive (equivalente ao
# `d.CompareMode = vbTextCompare` do VBA original).
TERMOS_EXCLUSAO_CODIGO_03 = [
    "Base",
    "Base reduzida",
    "Outros",
    "Direita",
    "Centro",
    "Esquerda",
    "Branco/Nulo",
    "Votaria em todos",
    "Nenhum",
    "Aumenta",
    "Diminui",
    "Não sabe",
    "Não sabe avaliar",
    "Não conhece",
    "Conhece um pouco ",
    "Conhece de ouvir falar ",
    "Conhece bem ",
    "Positiva",
    "Negativa",
    "Regular",
    "Não rejeito nenhum ",
    "Não sabe/Indeciso",
    "Não irá votar ",
    "Não foi votar ",
    "Concordo totalmente",
    "Nem concordo nem discordo",
    "Discordo totalmente",
    "Um candidato opositor a Família Reis ",
    "Um candidato da Família Reis ",
    "Um candidato aliado aos Reis, mas que não seja da família; ",
    "Depende do candidato",
    "Conhece pela foto",
    "Conhece muito",
    "Conhece de nome",
    "Nunca vi",
    "Nunca ouvi falar",
    "Votarei Nulo ou em Branco ",
    "Concordo plenamente",
    "Concordo",
    "Discordo",
    "Discordo plenamente",
    "Nenhuma palavra",
    "É conhecedor de Brasília e das prioridades das pessoas ",
    "Sempre",
    "Frequentemente",
    "Eventualmente",
    "Raramente",
    "Nunca",
    "Afeta muito",
    "Afeta um pouco",
    "Não afeta",
    "Atualmente participa",
    "Já participou",
    "Nunca participou",
    "Algumas vezes por semana",
    "Algumas vezes no mês",
    "Grande interesse",
    "Interesse moderado",
    "Baixo interesse",
    "Diariamente",
    "Certamente votaria",
    "Provavelmente votaria",
    "Provavelmente não votaria",
    "Certamente não votaria",
    "Certamente voto",
    "Certamente não voto",
    "Não conheço suficiente para votar",
    "Provavelmente voto",
    "Provavelmente não voto",
    "Todas as alternativas",
    "Nenhum desses",
    "Branco / Nulo",
    "Ajuda muito na decisão",
    "Ajuda na decisão",
    "Atrapalha na decisão",
    "Atrapalha muito na decisão",
    "Preparado",
    "Despreparado",
    "Experiente",
    "Inexperiente",
    "Tem influência política",
    "Não tem influência política",
    "Conhece a cidade",
    "Não sei",
    "Não votarei",
    "Aprovação",
    "Reprovação",
    "Nenhum deles",
    "Melhorou",
    "Piorou",
    "Continua da mesma forma",
    "Nada a melhorar",
    "Vai melhorar",
    "Ficar como está",
    "Vai piorar",
    "Confia muito",
    "Confia",
    "Confia mais ou menos",
    "Confia pouco",
    "Não confia",
    "Trocaria",
    "Não trocaria",
    "Outras cidades",
    "Não faz nada",
    "Nada contra a gestão atual",
    "Votou branco/nulo",
    "Está preparado",
    "Não está preparado",
    "Tem capacidade",
    "Não tem capacidade",
    "Tem força política",
    "Não tem força política",
    "Tem autonomia",
    "Não tem autonomia",
    "Ajudaria",
    "Não ajudaria",
    "Demais, deveria diminuir",
    "Na medida certa",
    "Pouca, deveria aumentar",
    "Lembra",
    "Não lembra",
    "Desde que nasceu",
    "Mora há mais de 20 anos",
    "Mora entre 20 e 10 anos",
    "Mora entre 10 e 5 anos",
    "Mora há menos de 5 anos",
    "Irá cumprir todas as promessas que faz",
    "Irá cumprir somente algumas",
    "Não irá cumprir nenhuma promessa que faz",
    "Promete muito mais do que conseguirá fazer em Maricá",
    "Promete na medida certa",
    "Promete pouco para Maricá, pois conseguiria fazer mais",
    "Sim",
    "Não",
    "Indiferente",
    "Outros aspectos",
    "Nada de bom",
    "Não vê nada de ruim",
    "Não sente falta de nada",
    "Muito interesse",
    "Interesse regular",
    "Nenhum interesse",
    "Excelente",
    "Boa",
    "Ruim",
    "Péssima",
    "Votou Branco ou Nulo",
    "Repro- vação",
    "Média",
    "Otimista",
    "Pessimista",
    "Nada importante",
    "Pouco importante",
    "Importante",
    "Muito importante",
    "NS",
    "Evoluindo",
    "Mesma coisa / parada",
    "Regredindo",
    "Da esquerda",
    "Da direita",
    "Do centro",
    "Conheço bem",
    "Conheço de ouvir falar",
    "Conheço das redes sociais",
    "Não conheço",
    "Indeciso",
    "Aumenta minha chance de votar nesse candidato",
    "É indiferente o apoio",
    "Diminui a minha chance de votar nesse candidato",
    "Não lembra/ Não tem candidato",
    "Não sabe em quem votar",
    "Outros candidatos",
    "Nulo / Branco",
    "Outras áreas",
    "Nenhuma área",
    "Todas as áreas",
    "Outras necessidades",
    "Nenhuma necessidade",
    "Influencia meu voto no candidato que apoiar",
    "Posso considerar o voto no candidato apoiado",
    "Posso considerar não votar no candidato apoiado",
    "Influencia a não votar no candidato apoiado",
    "Conheço muito",
    "Conheço um pouco",
    "Conheço somente de foto",
    "Conheço de foto, mas, não sabia que era Rodrigo Bacellar",
    "Nunca ouvi falar e nunca vi",
    "Tenho informações positivas",
    "Tenho informações negativas",
    "Não tenho informações",
    "Já viu",
    "Nunca viu",
    "Nenhum dos dois governos está bem",
    "Concorda totalmente",
    "Concorda em parte",
    "Não concorda e nem discorda",
    "Discorda em parte",
    "Discorda totalmente",
    "Melhor",
    "Igual",
    "Pior",
    "Todos devem permanecer presos e ser condenados",
    "A prisão deve ser apenas para quem danificou prédios",
    "Todos devem receber anistia",
    "Maricá oferece vagas de emprego até demais para os moradores",
    "Maricá oferece emprego na medida certa, todos poderiam estar empregados",
    "Maricá oferece poucas vagas de emprego, e vai oferecer cada vez menos",
    "Maricá oferece poucas vagas de emprego, mas, começará a oferecer mais",
    "O turismo está crescendo, mas falta gerar emprego e renda aos moradores",
    "O turismo está crescendo e tem ajudado a gerar emprego e renda aos moradores",
    "O turismo está estagnado",
    "O turismo está diminuindo",
    "Maricá está crescendo, mas sem planejamento e ficará pior",
    "Maricá está crescendo, sem planejamento, mas haverá uma organização",
    "Maricá já está crescendo de forma organizada",
    "Maricá está estagnada",
    "Maricá está regredindo",

    # Termos adicionados por Lucas
    "Candidato que seja de Teresópolis",
    "Candidato que não seja de Teresópolis, mas que tem histórico de realizações na cidade",
    "Nunca viu antes",
    "Nunca ouviu falar",
    "Conhece por vídeo",
    "É preparada para exercer o cargo de deputada estadual",
    "Não é preparada para exercer o cargo de deputado estadual",
    "É conhecedora de Teresópolis e das prioridades das pessoas",
    "Não conhece a cidade e o que as pessoas precisam",
    "Tem influência política e pode trazer benefícios para a cidade",
    "Não tem influência política e não pode trazer benefícios para a cidade",
    "Tem vontade de trazer melhorias para Teresópolis",
    "Não tem vontade de trazer melhorias para Teresópolis",
    "Não acompanha",
    "Ótima",
    "Ótimo",
    "Bom",
    "Péssimo",
    "É preparado para exercer o cargo de deputado federal",
    "Não é preparado para exercer o cargo de deputado federal?",
    "Não é preparado para exercer o cargo de deputado federal",
    "É conhecedor de Duque de Caxias e das prioridades das pessoas",
    "Tem vontade de trazer melhorias para Duque de Caxias",
    "Não tem vontade de trazer melhorias para Duque de Caxias",
    "Nunca vi antes",
    "Não lembro o nome",
    "Não lembra o nome",
    "Não sei/Não lembro",
    "Votei Nulo ou Em branco",
    "Não fui votar",
    # Termos adicionados por Lucas (lote seguinte)
    "É preparado para exercer o cargo de deputado estadual",
    "Não é preparado para exercer o cargo de deputado estadual",
    "Nenhuma prioridade",
    "Centralizadas nos bairros mais ricos",
    "Realizadas em todos os bairros",
    "Centralizadas nos bairros mais pobres",
    "Conhecia bem",
    "Conhecia algumas informações",
    "Apenas tinha ouvido falar",
    "Nunca tinha ouvido falar",
    "Aprova",
    "Desaprova",
    "Não aprova nem desaprova",
    "Não conhece o suficiente para opinar",
    "Geração de empregos para os moradores de Maricá",
    "Aumento do turismo e da atividade econômica",
    "Melhoria da infraestrutura da região",
    "Apenas cerca de 6,6% dos 844 hectares do terreno serem ocupados por edificações",
    "Preservação ou recuperação ambiental de aproximadamente 81% da propriedade",
    "O projeto possuir licenças dos órgãos ambientais e urbanísticos competentes",
    "O projeto será aberto e integrado a cidade, podendo visitar a praia, a lagoa e a restinga",
    "Nenhuma dessas características",
    "Construção dentro de uma área de restinga",
    "Possíveis impactos sobre animais, plantas, lagoas e praias",
    "Aumento do trânsito e da ocupação urbana",
    "Possíveis impactos sobre pescadores e comunidades locais",
    "Participação ou investimento da Prefeitura no empreendimento",
    "Ser um empreendimento voltado ao turismo e às residências de alto padrão",
    "Falta de confiança na fiscalização e no cumprimento das medidas ambientais",
    "Nenhum desses aspectos",
    "Rejeito todos",
    "Não foi votar",
    "Branco/Nulo",
    "Não irá votar",
    "Não sabe",
    "A decisão é definitiva",
    "Está bastante decidido, mas poderia mudar",
    "Tem preferência por um candidato, mas ainda está avaliando",
    "Está realmente em dúvida",
    "Ainda não tem preferência",
    "Nada faria mudar o voto",
    "Outros",
    "Nenhum desses fatores faria mudar o voto",
    "Não há prioridade",
    "Não sabe",
    "Não há prioridade",
    "Nenhum",
    "Ambos",
    "Totalmente a favor",
    "Mais a favor",
    "Mais contra",
    "Totalmente contra",
    "Viu / Ouviu",
    "Não viu / Não ouviu",
    "Viu /Ouviu",
    "Não viu Não ouviu",
    "Viu/Ouviu",
    "Não viu/Não ouviu",
    "Não lembra",
    "Nenhum desses fatores faria mudar meu voto",
    "Nada faria mudar o voto",
    "Não rejeito nenhum deles",
]


def aplicar_codigo_03(ws, termos_exclusao=None):
    """
    Localiza cada célula com o texto exato "Total" (mesclada ou não,
    comparação sem diferenciar maiúsc/minúsc). Abaixo de cada uma,
    considera como "bloco de categorias" as linhas seguintes enquanto a
    coluna B não estiver vazia e não for outro "Total". Dentro desse
    bloco, reordena as linhas cujo rótulo (coluna A) NÃO está na lista de
    exclusão, por valor numérico da coluna B em ordem DECRESCENTE — as
    linhas da lista de exclusão (ex.: "Não sabe", "Branco/Nulo", "Base")
    ficam paradas na posição em que já estavam, servindo de âncora fixa
    entre os grupos que são de fato reordenados.

    Só troca os VALORES das células (todas as colunas usadas), igual ao
    VBA original — não mexe em formatação, então cada linha mantém a
    aparência (fonte, borda, preenchimento) que já tinha na posição onde
    ela está, só o conteúdo é que muda de lugar.

    Args:
        ws: planilha a processar.
        termos_exclusao: lista de rótulos a manter fixos (case-insensitive).
            Se None, usa `TERMOS_EXCLUSAO_CODIGO_03` (lista padrão herdada
            do VBA original).

    Returns:
        Número de tabelas (blocos "Total") que tiveram linhas reordenadas.
    """
    termos = {t.strip().lower() for t in (termos_exclusao or TERMOS_EXCLUSAO_CODIGO_03)}
    max_col = ws.max_column
    max_row = ws.max_row

    # Acha todas as células "Total" (exato, sem diferenciar maiúsc/minúsc),
    # em ordem de leitura (linha por linha, como o Find/FindNext padrão do Excel)
    celulas_total = [
        cel
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col)
        for cel in row
        if cel.value is not None and str(cel.value).strip().lower() == "total"
    ]

    tabelas_reordenadas = 0

    for cel in celulas_total:
        # Se a célula "Total" está mesclada, o bloco de categorias começa
        # logo após o fim da mesclagem; senão, na linha seguinte.
        rng_merge = None
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= cel.row <= rng.max_row and rng.min_col <= cel.column <= rng.max_col:
                rng_merge = rng
                break
        linha_inicio = (rng_merge.max_row + 1) if rng_merge is not None else (cel.row + 1)

        # Bloco vai até a coluna B ficar vazia ou repetir "Total"
        linha_fim = linha_inicio
        while linha_fim <= max_row:
            valor_b = ws.cell(row=linha_fim, column=2).value
            valor_b_str = "" if valor_b is None else str(valor_b).strip()
            if valor_b_str == "" or valor_b_str.lower() == "total":
                break
            linha_fim += 1
        linha_fim -= 1

        if linha_fim < linha_inicio:
            continue

        linhas_para_ordenar = [
            r for r in range(linha_inicio, linha_fim + 1)
            if str(ws.cell(row=r, column=1).value or "").strip().lower() not in termos
        ]
        if len(linhas_para_ordenar) < 2:
            continue

        def _chave_ordenacao(r):
            v = ws.cell(row=r, column=2).value
            try:
                return float(str(v).replace(",", ".")) if v is not None else float("-inf")
            except (TypeError, ValueError):
                return float("-inf")

        # Snapshot dos valores ANTES de escrever nada (senão sobrescreveria
        # linhas que ainda vamos ler, já que a origem/destino se cruzam)
        valores_por_linha = {
            r: [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
            for r in linhas_para_ordenar
        }
        linhas_ordenadas = sorted(linhas_para_ordenar, key=_chave_ordenacao, reverse=True)

        # Reescreve nas mesmas posições originais (linhas_para_ordenar, em
        # ordem), usando os valores na ordem já ordenada — exatamente como
        # o VBA faz ao reler tempArray (ordenado) de volta nas linhas
        # originais na sequência em que apareciam.
        for linha_destino, linha_origem in zip(linhas_para_ordenar, linhas_ordenadas):
            for c, valor in enumerate(valores_por_linha[linha_origem], start=1):
                ws.cell(row=linha_destino, column=c, value=valor)

        tabelas_reordenadas += 1

    return tabelas_reordenadas


# =========================================================
# Código 06 — Ajuste DIN 9 na linha de Pergunta (+ autofit aproximado)
# =========================================================
def aplicar_codigo_06(ws):
    """
    Reduz para tamanho 9 a fonte de qualquer célula contendo 'Pergunta:'.
    NÃO faz o autofit real (não é possível fora do Excel) — só ajusta a
    fonte, que é a parte 100% confiável deste código.
    """
    alterados = 0
    max_col = ws.max_column
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            if isinstance(cel.value, str) and "pergunta:" in cel.value.lower():
                fonte = cel.font
                cel.font = Font(
                    name=fonte.name, size=9, bold=fonte.bold, italic=fonte.italic,
                    vertAlign=fonte.vertAlign, underline=fonte.underline,
                    strike=fonte.strike, color=fonte.color,
                )
                alterados += 1
    return alterados


# =========================================================
# Código 08 — Base pequena
# =========================================================
def aplicar_codigo_08(ws):
    """
    Em toda linha 'Base'/'Base reduzida', marca com '*' os valores abaixo
    de 30 (maiores que zero) e insere uma linha de nota "*Base pequena para
    análise estatística" logo abaixo (pulando a linha 'Média', se ela
    estiver a até 3 linhas de distância) — sempre em fonte DIN tamanho 9.
    """
    linhas_marcadas = 0
    r = 1
    max_col = ws.max_column
    # `ws.max_row` recalcula varrendo a planilha inteira toda vez que é
    # acessado — reconsultar isso a cada volta de um laço `while` que
    # roda linha por linha é caro em relatórios grandes. Rastreia o
    # valor manualmente: só aumenta quando `_analisar_linha_codigo_08`
    # realmente insere a linha de nota (é a única coisa que desloca o
    # fim do arquivo aqui).
    max_row_atual = ws.max_row
    while r <= max_row_atual:
        texto = str(ws.cell(row=r, column=1).value or "").strip().upper()
        if texto in ("BASE", "BASE REDUZIDA"):
            deslocado = _analisar_linha_codigo_08(ws, r, max_col)
            if deslocado:
                linhas_marcadas += 1
                max_row_atual += deslocado
                r += deslocado  # o resto do arquivo desceu — pula a nota recém-inserida
        r += 1
    return linhas_marcadas


def _analisar_linha_codigo_08(ws, linha, max_col):
    col_max = 1
    for c in range(1, max_col + 1):
        if ws.cell(row=linha, column=c).value is not None:
            col_max = c

    mensagem_adicionada = False
    deslocamento_total = 0

    for coluna in range(1, col_max + 1):
        cel = ws.cell(row=linha, column=coluna)
        valor = cel.value
        # abaixo de 30 (antes: só entre 20 e 30). Zero fica de fora: coluna sem
        # entrevistas não é "base pequena". bool é subclasse de int — ignora.
        if isinstance(valor, (int, float)) and not isinstance(valor, bool) and 0 < valor < 30:
            cel.value = f"{valor}*"
            cel.alignment = _com_wrap_text(cel.alignment, False)

            if not mensagem_adicionada:
                tem_media = any(
                    str(ws.cell(row=linha + i, column=1).value or "").strip().upper() == "MÉDIA"
                    for i in range(1, 4)
                )
                offset = 3 if tem_media else 2
                inserir_linhas_seguro(ws, linha + offset, 1)
                nota = ws.cell(row=linha + offset, column=1, value="*Base pequena para análise estatística")
                nota.alignment = _com_wrap_text(nota.alignment, False)
                nota.font = Font(name="DIN", size=9)
                mensagem_adicionada = True
                deslocamento_total = offset

    return deslocamento_total


def _com_wrap_text(alinhamento_atual, valor):
    """Retorna uma cópia do alinhamento existente, só trocando wrap_text — preserva
    horizontal/vertical/indent/etc. (o VBA original só mexe no WrapText)."""
    a = alinhamento_atual
    return Alignment(
        horizontal=a.horizontal, vertical=a.vertical, text_rotation=a.text_rotation,
        wrap_text=valor, shrink_to_fit=a.shrink_to_fit, indent=a.indent,
        justifyLastLine=a.justifyLastLine, readingOrder=a.readingOrder,
    )


def _com_horizontal_vertical(alinhamento_atual, horizontal, vertical):
    """Retorna uma cópia do alinhamento existente, só trocando horizontal/vertical —
    preserva wrap_text/indent/etc. (o VBA original só mexe nesses dois quando
    define HorizontalAlignment/VerticalAlignment, nunca no WrapText)."""
    a = alinhamento_atual
    return Alignment(
        horizontal=horizontal, vertical=vertical, text_rotation=a.text_rotation,
        wrap_text=a.wrap_text, shrink_to_fit=a.shrink_to_fit, indent=a.indent,
        justifyLastLine=a.justifyLastLine, readingOrder=a.readingOrder,
    )


# =========================================================
# Código 09 — Ajusta a altura das labels (autofit aproximado)
# =========================================================
# Cada grupo: (gatilhos de texto, altura fixa, coluna mínima onde vale
# — None = qualquer coluna). Comparação por "começa com",
# case-insensitive. Escolaridade ('Ensino...') foi pra 45 a pedido do
# Lucas (era 30, junto com Sexo, até essa mudança); Sexo
# ('Masculino'/'Feminino') continua em 30, em qualquer coluna. Os
# segmentos de avaliação ('Aprovação'/'Regular'/'Reprovação') também
# ficam em 30, mas SÓ quando aparecem da coluna C em diante — se
# aparecerem na coluna A (rótulo da linha, não valor de segmento), não
# deve ser alterado.
_GATILHOS_ALTURA_FIXA = (
    (("masculino", "feminino"), 30, None),
    (("ensino",), 45, 3),
    (("regular",), 30, 3),
    (("agnóstico", "ateu", "judeu", "orientais"), 45, 3),
    # Gatilhos pelo texto JÁ REFORMATADO pelo código 04 (que roda antes
    # — por isso os prefixos aqui já batem com a versão com quebra de
    # linha/hífen, ex.: "Evangé-\nlica\nTotal" começa com "evangé-").
    # Sem isso, o código 09 (que ativa "Quebrar Texto Automaticamente"
    # em tudo) recalculava a altura por estimativa de caracteres pra
    # qualquer termo que não fosse um gatilho aqui, sobrescrevendo a
    # altura fixa que o código 04 já tinha aplicado.
    (
        (
            "evangé-", "católica", "pente-", "espírita/", "afro-",
            "teste-", "não\nsabe", "outras\nreligiões",
        ),
        45, 3,
    ),
    (("apro-", "repro-", "jair"), 30, 3),
)


def _celulas_gatilho_altura_fixa(ws, r, max_col):
    """
    Se a linha tiver algum gatilho de altura fixa (ver
    `_GATILHOS_ALTURA_FIXA`), retorna (altura, lista de colunas que
    bateram o gatilho) — essas colunas precisam de "Quebrar Texto
    Automaticamente" ativado, mesmo não tendo o texto reformatado (isso
    é feito só pelo código 04; aqui só a altura é fixa). Se não bater
    nenhum gatilho, retorna (None, []).
    """
    for gatilhos, altura, coluna_minima in _GATILHOS_ALTURA_FIXA:
        col_inicio = coluna_minima or 1
        colunas_batidas = []
        for c in range(col_inicio, max_col + 1):
            valor = ws.cell(row=r, column=c).value
            if isinstance(valor, str):
                texto = valor.strip().lower()
                if any(texto.startswith(gatilho) for gatilho in gatilhos):
                    colunas_batidas.append(c)
        if colunas_batidas:
            return altura, colunas_batidas
    return None, []


def _altura_linha_de_label(ws, r, max_col):
    """Altura de uma linha de tabela cuja coluna A é uma label.

    A label (coluna A) segue a regra fixa de caracteres
    (`altura_label_por_caracteres`: até 23 -> 15, até 46 -> 30...) como
    piso, podendo subir se a largura real do texto não couber (ver
    `altura_label_celula` em core/planilha_utils.py), e
    ganha "Quebrar Texto Automaticamente" (sem isso o Excel não quebra e
    a altura extra ficaria sobrando). Se alguma OUTRA célula de texto da
    linha, com quebra ativada, precisar de mais linhas, prevalece a
    maior — também em degraus de 15. Números/percentuais são ignorados.
    """
    cel_a = ws.cell(row=r, column=1)
    if not (cel_a.alignment and cel_a.alignment.wrap_text):
        cel_a.alignment = _com_wrap_text(cel_a.alignment, True)
    altura = altura_label_celula(ws, cel_a)
    for c in range(2, max_col + 1):
        cel = ws.cell(row=r, column=c)
        if not isinstance(cel.value, str) or not cel.value.strip():
            continue
        if not (cel.alignment and cel.alignment.wrap_text):
            continue
        largura = largura_total_mesclagem(ws, cel)
        n_linhas = estimar_numero_linhas(cel.value, largura, cel.font.size or 10)
        altura = max(altura, n_linhas * ALTURA_POR_LINHA_LABEL)
    return altura


def aplicar_codigo_09(ws):
    """
    LABELS DA COLUNA A (regra fixa, pedido do Lucas): em linha de tabela
    cuja coluna A tem uma label (não título, não pergunta, não mesclada),
    a altura é por nº de caracteres — até 23 -> 15, até 46 -> 30, até 69
    -> 45... (ver `_altura_linha_de_label`). As demais linhas continuam
    com a estimativa descrita abaixo.

    PRIORIDADE MÁXIMA: linha que contenha "Denominações Católicas" ou
    "Denominações Evangélicas" (em qualquer coluna) fica com altura 15
    (ver `ALTURAS_FIXAS_POR_TEXTO` em core/planilha_utils.py).

    Para toda célula com quebra de linha (WrapText) ativada, estima a
    altura necessária (aproximação por contagem de caracteres — ver
    aviso em `core.planilha_utils.estimar_altura_calculada`) e aplica na
    linha. O AutoFit real do Excel ajusta a altura pra caber exatamente
    (pode até diminuir uma linha que estava alta demais) — então aqui
    também aplica sempre a estimativa, não só quando ela é maior que a
    altura atual.

    Exceção: linhas com algum gatilho de `_GATILHOS_ALTURA_FIXA`
    (subcategoria de Sexo, Escolaridade, ou segmento de avaliação)
    sempre ficam com a altura fixa correspondente, em vez da
    estimativa — um valor pré-definido, não calculado. A(s) célula(s)
    que bateram o gatilho também ganham "Quebrar Texto Automaticamente"
    ativado.
    """
    alterados = 0
    # `ws.max_column`/`ws.max_row` recalculam varrendo a planilha inteira
    # toda vez que são acessados — chamar isso dentro de um laço por
    # linha (e pior ainda, dentro de outro laço por grupo de gatilho)
    # é a causa principal deste código ficar lento em relatórios
    # grandes. Calcula uma vez só aqui, fora dos laços.
    max_col = ws.max_column
    max_row = ws.max_row
    linhas_mescladas_a = linhas_com_mesclagem_na_coluna_a(ws)
    for r in range(1, max_row + 1):
        altura_texto = altura_fixa_por_texto(ws, r, max_col)
        if altura_texto is not None:
            if ws.row_dimensions[r].height != float(altura_texto):
                ws.row_dimensions[r].height = float(altura_texto)
                alterados += 1
            continue

        altura_fixa, colunas_batidas = _celulas_gatilho_altura_fixa(ws, r, max_col)
        if altura_fixa is not None:
            for c in colunas_batidas:
                cel = ws.cell(row=r, column=c)
                cel.alignment = _com_wrap_text(cel.alignment, True)
            if ws.row_dimensions[r].height != float(altura_fixa):
                ws.row_dimensions[r].height = float(altura_fixa)
                alterados += 1
            continue

        if eh_label_de_tabela_coluna_a(ws, r, max_col, linhas_mescladas_a):
            nova_altura = float(_altura_linha_de_label(ws, r, max_col))
            if ws.row_dimensions[r].height != nova_altura:
                ws.row_dimensions[r].height = nova_altura
                alterados += 1
            continue

        maior_estim = None
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            if not (cel.alignment and cel.alignment.wrap_text):
                continue
            if not cel.value:
                continue

            largura = largura_total_mesclagem(ws, cel)
            tamanho_fonte = cel.font.size or 10
            estim = estimar_altura_calculada(cel.value, largura, tamanho_fonte)
            if maior_estim is None or estim > maior_estim:
                maior_estim = estim

        if maior_estim is None:
            continue

        altura_atual = ws.row_dimensions[r].height
        nova_altura = round(maior_estim, 1)
        if altura_atual != nova_altura:
            ws.row_dimensions[r].height = nova_altura
            alterados += 1
    return alterados


# =========================================================
# Código 11 — Ajusta a altura dos Títulos (+ remove "Titulo: ")
# =========================================================
def aplicar_codigo_11(ws):
    """
    Para toda célula contendo 'Titulo: ' OU 'Título: ' (com acento —
    algumas exportações do SPSS usam a forma acentuada): estima a
    altura (aproximada, igual ao código 09) e aplica na "escada"
    30/45/60/75... Remove o prefixo ('Titulo: ' ou 'Título: ', o que
    estiver presente) do texto, mantendo só o título em si, centraliza
    o texto horizontal e verticalmente, e deixa em negrito.
    """
    alterados = 0
    max_col = ws.max_column
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            if not isinstance(cel.value, str):
                continue
            valor_lower = cel.value.lower()
            if "titulo: " not in valor_lower and "título: " not in valor_lower:
                continue

            largura = largura_total_mesclagem(ws, cel)
            tamanho_fonte = cel.font.size or 10
            altura_calc = estimar_altura_calculada(cel.value, largura, tamanho_fonte) + 2
            altura_final = _altura_staircase_titulo(altura_calc)
            ws.row_dimensions[r].height = altura_final

            texto_novo = _remover_prefixo_case_insensitive(cel.value, "Titulo: ")
            texto_novo = _remover_prefixo_case_insensitive(texto_novo, "Título: ")
            cel.value = texto_novo
            cel.alignment = _com_horizontal_vertical(cel.alignment, "center", "center")

            fonte_atual = cel.font
            cel.font = Font(
                name=fonte_atual.name, size=fonte_atual.size, bold=True,
                italic=fonte_atual.italic, color=fonte_atual.color,
            )
            alterados += 1
    return alterados


def _altura_staircase_titulo(altura_calculada):
    altura = 30
    while altura < altura_calculada:
        altura += 15
    return altura


# =========================================================
# Código 12 — Ajuste da altura das Perguntas
# =========================================================
def aplicar_codigo_12(ws):
    """
    Para toda célula contendo 'Pergunta: ': estima quantas linhas o
    texto vai ocupar (quebra automática, mesma aproximação usada nos
    demais códigos de altura — ver `estimar_numero_linhas`) e aplica
    altura = número de linhas × 15 (1 linha -> 15, 2 linhas -> 30, 3
    linhas -> 45...). Fórmula fixa a pedido do Lucas — substituiu a
    "escada" antiga (20 -> 30 -> +15/+15...). Não remove o texto (esse
    código, diferente do 11, mantém 'Pergunta: ' no texto).
    """
    alterados = 0
    max_col = ws.max_column
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            if not (isinstance(cel.value, str) and "pergunta: " in cel.value.lower()):
                continue

            largura = largura_total_mesclagem(ws, cel)
            tamanho_fonte = cel.font.size or 10
            n_linhas = estimar_numero_linhas(cel.value, largura, tamanho_fonte)
            altura_final = n_linhas * 15
            if ws.row_dimensions[r].height != altura_final:
                ws.row_dimensions[r].height = altura_final
                alterados += 1
    return alterados


# =========================================================
# Código 05 — Base reduzida (exclusão de colunas, direto na planilha)
# =========================================================
# A lógica de detectar blocos de tabela e excluir colunas de "Base
# reduzida" abaixo de um limite (com mesclagem preservada) mora em
# core/planilha_utils.py — compartilhada com o Processador de Base
# Reduzida standalone (analises/base_reduzida.py), pra não duplicar essa
# lógica (já testada com arquivo real, inclusive com cabeçalho mesclado
# tipo "Regiões") em dois lugares.
from core.planilha_utils import (
    normalizar_texto_maiusculo as _normalizar_texto_base,
    linha_vazia_ate_coluna as _linha_vazia_ate_coluna,
    encontrar_inicio_bloco as _encontrar_inicio_bloco,
    encontrar_fim_bloco as _encontrar_fim_bloco,
    excluir_colunas_base_reduzida as _excluir_colunas_base_reduzida,
)

def _remover_linha_extra_multipla(ws):
    """
    Logo abaixo de um título mesclado de pergunta do tipo Múltipla
    ('Múltipla', 'Multipla' ou 'Estimulada e Múltipla'), se a linha
    imediatamente seguinte estiver totalmente em branco, remove essa
    linha (efeito de uma linha extra que sobra nesse tipo de tabela).
    Processa de baixo para cima.

    A comparação ignora acento (usa `unicodedata` pra tirar diacríticos
    antes de comparar) — sem isso, o texto real do relatório ('Múltipla',
    com acento) nunca batia contra os literais 'MULTIPLA'/'ESTIMULADA E
    MULTIPLA' (sem acento), e a linha extra nunca era removida na
    prática. Bug antigo, corrigido aqui — não é específico do Total
    Automático, também afeta o código 05 do Relatório Automatizado.
    """
    import unicodedata

    def _sem_acento(texto):
        return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")

    max_col = ws.max_column
    removidas = 0
    r = ws.max_row
    while r >= 1:
        cel = ws.cell(row=r, column=1)
        faixa = None
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= 1 <= rng.max_col:
                faixa = rng
                break

        if faixa is not None and faixa.min_row == r:
            texto = _sem_acento(_normalizar_texto_base(ws.cell(row=faixa.min_row, column=1).value))
            eh_multipla = "ESTIMULADA E MULTIPLA" in texto or "MULTIPLA" in texto
            if eh_multipla:
                linha_abaixo = faixa.max_row + 1
                if linha_abaixo <= ws.max_row and _linha_vazia_ate_coluna(ws, linha_abaixo, max_col):
                    ws.delete_rows(linha_abaixo, 1)
                    removidas += 1
        elif faixa is None:
            texto = _sem_acento(_normalizar_texto_base(cel.value))
            eh_multipla = "ESTIMULADA E MULTIPLA" in texto or "MULTIPLA" in texto
            if eh_multipla:
                linha_abaixo = r + 1
                if linha_abaixo <= ws.max_row and _linha_vazia_ate_coluna(ws, linha_abaixo, max_col):
                    ws.delete_rows(linha_abaixo, 1)
                    removidas += 1
        r -= 1

    return removidas


def aplicar_layout_basico_planilha(ws):
    """
    Parte "layout" do VBA 'ConfigurarUmaPlanilha' — tudo, EXCETO a
    exclusão de colunas de Base reduzida (essa fica isolada em
    `aplicar_codigo_05`, que chama esta função por baixo). Extraída à
    parte pra ser reaproveitada por outros fluxos que precisam do mesmo
    layout sem a exclusão por limite (ex.: Total Automático).

    - fonte DIN 10 e alinhamento centralizado em toda a planilha;
    - coluna A alinhada à esquerda (horizontal) e centralizada (vertical);
    - margens, orientação PAISAGEM e papel A4;
    - largura da coluna A = 21, coluna B = 8, demais colunas = 8.67
      (+ compensação empírica AJUSTE_LARGURA_COLUNA_DIN);
    - escala de impressão 100% (não "ajustar à largura");
    - toda célula mesclada na coluna A contendo 'Pergunta' ou '*' alinhada
      à esquerda;
    - linhas de grade ATIVADAS (a planilha do SPSS às vezes já vem com
      elas desligadas — força ligado em vez de só herdar o que quer que
      o arquivo de origem tiver).

    (Não remove mais a linha extra abaixo de títulos 'Múltipla' — tinha
    uma tentativa de correção aqui, mas não funcionou como esperado num
    arquivo real e foi desativada a pedido do Lucas. A função
    `_remover_linha_extra_multipla` continua no arquivo, só não é mais
    chamada; ver comentário logo abaixo.)
    """
    from openpyxl.worksheet.page import PageMargins

    ws.sheet_view.showGridLines = True

    fonte_padrao = Font(name="DIN", size=10)

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cel = ws.cell(row=r, column=c)
            cel.font = fonte_padrao
            horizontal = "left" if c == 1 else "center"
            cel.alignment = _com_horizontal_vertical(cel.alignment, horizontal, "center")

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(
        left=1.5 / 2.54, right=1.3 / 2.54, top=2.5 / 2.54, bottom=2 / 2.54,
        header=1.3 / 2.54, footer=1.3 / 2.54,
    )
    # Escala de impressão em 100% (Layout de Página > Escala), não o
    # modo "ajustar a X página(s) de largura" — os dois são mutuamente
    # exclusivos no Excel via pageSetUpPr.fitToPage.
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.fitToWidth = None
    ws.page_setup.fitToHeight = None
    ws.page_setup.scale = 100

    # O Excel parece recalcular a largura "em caracteres" exibida usando
    # a métrica da fonte DIN (aplicada nas células logo abaixo) em vez da
    # fonte "Normal" do arquivo (Calibri, que é o que o openpyxl usa como
    # referência ao gravar o número puro) — resultado: gravar 20.67 faz
    # o Excel mostrar ~19.89 na caixa de diálogo de largura de coluna, um
    # desvio de ~0.78 constante (não proporcional) nos testes feitos.
    # Compensa gravando o valor-alvo + 0.78, calibrado empiricamente
    # contra esse comportamento — não é uma unidade "oficial" documentada
    # do Excel, então se um dia isso mudar de comportamento (nova versão
    # do Excel, fonte diferente etc.), essa constante pode precisar ser
    # recalibrada.
    AJUSTE_LARGURA_COLUNA_DIN = 0.78
    ws.column_dimensions["A"].width = 21 + AJUSTE_LARGURA_COLUNA_DIN
    if ws.max_column >= 2:
        ws.column_dimensions["B"].width = 8 + AJUSTE_LARGURA_COLUNA_DIN
    for c in range(3, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = 8.67 + AJUSTE_LARGURA_COLUNA_DIN

    for r in range(1, ws.max_row + 1):
        cel = ws.cell(row=r, column=1)
        faixa = None
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= 1 <= rng.max_col:
                faixa = rng
                break
        if faixa is not None:
            texto = str(ws.cell(row=faixa.min_row, column=1).value or "")
            if "pergunta" in texto.lower() or "*" in texto:
                cel_ref = ws.cell(row=faixa.min_row, column=1)
                a = cel_ref.alignment
                cel_ref.alignment = Alignment(
                    horizontal="left", vertical=a.vertical, text_rotation=a.text_rotation,
                    wrap_text=a.wrap_text, shrink_to_fit=a.shrink_to_fit, indent=a.indent,
                    justifyLastLine=a.justifyLastLine, readingOrder=a.readingOrder,
                )

    # Removida a chamada de `_remover_linha_extra_multipla(ws)` — a
    # correção do bug de acentuação (comparava com 'MULTIPLA' sem
    # acento contra texto real 'MÚLTIPLA' com acento) não resolveu o
    # problema na prática num arquivo real; desativada a pedido do
    # Lucas até investigar com mais calma. A função continua definida
    # acima, só não é mais chamada daqui.


def aplicar_codigo_05(ws, limite=25):
    """
    Porta o VBA 'ConfigurarUmaPlanilha' por completo: aplica o layout
    básico (ver `aplicar_layout_basico_planilha`) e, por fim, exclui as
    colunas de Base reduzida abaixo do limite escolhido.
    """
    aplicar_layout_basico_planilha(ws)
    excluidas = _excluir_colunas_base_reduzida(ws, limite)

    return excluidas



def aplicar_codigo_06_autofit(ws):
    """
    Aproxima o `usedRange.Rows.AutoFit` do VBA original: para toda célula
    com texto, estima a altura necessária (mesma aproximação usada nos
    códigos 09/11/12) e aplica na linha, se maior que a altura atual.
    Não sobrescreve pra baixo — só aumenta quando o texto realmente
    precisar de mais espaço.
    """
    alterados = 0
    for r in range(1, ws.max_row + 1):
        maior_estim = None
        for c in range(1, ws.max_column + 1):
            cel = ws.cell(row=r, column=c)
            if not cel.value or not isinstance(cel.value, str):
                continue

            largura = largura_total_mesclagem(ws, cel)
            tamanho_fonte = cel.font.size or 10
            estim = estimar_altura_calculada(cel.value, largura, tamanho_fonte)
            if maior_estim is None or estim > maior_estim:
                maior_estim = estim

        if maior_estim is None:
            continue

        altura_atual = ws.row_dimensions[r].height or 15
        if maior_estim > altura_atual:
            ws.row_dimensions[r].height = round(maior_estim, 1)
            alterados += 1

    return alterados


# =========================================================
# Código 07 — Adiciona linhas + quebra de página
# =========================================================
def aplicar_codigo_07(ws):
    """
    Para toda célula na coluna A contendo 'Pergunta' ou '*' (mesclada ou
    não — o VBA original só olha células mescladas, mas isso deixa o
    código mais tolerante a variações do arquivo real): normaliza o vão
    logo abaixo pra ficar com exatamente 4 linhas em branco (insere a
    mais se faltar, remove o excedente se já houver linhas em branco ali
    de sobra) e adiciona uma quebra de página manual bem no meio dessas
    4 linhas (depois da 2ª, antes da 3ª). Separadamente, sempre que
    encontrar um trecho de 4 ou mais linhas em branco seguidas que NÃO
    esteja logo depois de um bloco "Pergunta"/"*" (esse já ganhou sua
    própria quebra ali do lado, então outra aqui seria redundante),
    insere uma quebra de página no meio dele — no máximo uma por trecho,
    não uma a cada 4 linhas dentro do mesmo vão vazio longo.

    OTIMIZAÇÃO: a versão anterior processava de baixo pra cima chamando
    `inserir_linhas_seguro`/`remover_linhas_seguro` uma vez POR
    "Pergunta" encontrada — cada chamada custa O(tamanho da planilha)
    porque é isso que `ws.insert_rows()`/`ws.delete_rows()` do openpyxl
    custam por baixo (percorrem TODAS as células pra deslocar, não só a
    região afetada), então N "Perguntas" custavam O(N × tamanho da
    planilha) no total — em relatórios com muitas tabelas, esse era o
    gargalo conhecido do código 07. Esta versão faz uma única LEITURA
    (sem mutar nada) pra decidir TODAS as inserções/remoções de uma vez,
    calcula o deslocamento acumulado de cada linha com um vetor de
    diferenças (soma prefixada, O(tamanho da planilha) no total, não por
    operação) e move célula, mesclagem, altura de linha e quebra de
    página numa única passada. Mesmo resultado final, só que O(tamanho
    da planilha) no total em vez de O(nº de Perguntas × tamanho da
    planilha).

    A leitura pura funciona porque cada "Pergunta" só edita linhas
    estritamente abaixo dela mesma, numa vizinhança curta (até a faixa
    de 4 linhas) — essa faixa nunca se sobrepõe ao trecho que OUTRA
    "Pergunta" lê como "vão logo abaixo", porque sempre há conteúdo de
    tabela de verdade separando um bloco do próximo. Processar de baixo
    pra cima mutando OU ler tudo de uma vez de cima pra baixo sem mutar
    dão exatamente as mesmas decisões.
    """
    from openpyxl.worksheet.pagebreak import Break

    # Fixa a impressão em escala 100% (modo "Ajustar a: 100% do tamanho
    # normal", não o modo "Ajustar a N página(s) de largura") — o Excel
    # trata os dois como mutuamente exclusivos via pageSetUpPr.fitToPage.
    # Com fitToPage=False + scale=100, o Excel usa a paginação natural da
    # planilha (incluindo as quebras manuais inseridas abaixo) sem
    # recalcular nada pra "caber" num número de páginas.
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.fitToWidth = None
    ws.page_setup.fitToHeight = None
    ws.page_setup.scale = 100

    def _e_gatilho(valor):
        texto = str(valor) if valor is not None else ""
        return bool(texto) and ("pergunta" in texto.lower() or "*" in texto)

    max_row_original = ws.max_row
    if max_row_original < 1:
        return 0, 0

    # ---- FASE 1: planejamento — uma leitura só, sem mutar nada -------
    valores_col_a = [None] * (max_row_original + 2)
    for r in range(1, max_row_original + 1):
        valores_col_a[r] = ws.cell(row=r, column=1).value

    operacoes = []  # (tipo, posicao_original, quantidade) — tipo: "inserir"/"remover"
    quebras_planejadas = []  # ("gatilho", linha_do_gatilho) | ("livre", posicao_original)

    r = 1
    while r <= max_row_original:
        valor = valores_col_a[r]
        if _e_gatilho(valor):
            gap_existente = 0
            rr = r + 1
            while rr <= max_row_original and valores_col_a[rr] is None:
                gap_existente += 1
                rr += 1

            if gap_existente < 4:
                operacoes.append(("inserir", r + 1, 4 - gap_existente))
            elif gap_existente > 4:
                operacoes.append(("remover", r + 5, gap_existente - 4))

            if r < max_row_original:
                # Quebra "no meio das 4, depois da 2ª" — guarda a linha do
                # GATILHO (não r+2 direto): quando o gap precisava de
                # inserção, r+2 pode cair numa linha recém-criada que não
                # existia na numeração original, então a posição final tem
                # que ser calculada a partir de onde o gatilho mesmo vai
                # parar (`_nova_posicao(r) + 2`), não remapeada como se
                # fosse uma linha original comum.
                quebras_planejadas.append(("gatilho", r))

            r = rr
            continue

        if valor is None:
            inicio_trecho = r
            rr = r
            while rr <= max_row_original and valores_col_a[rr] is None:
                rr += 1
            fim_trecho = rr - 1
            tamanho_trecho = fim_trecho - inicio_trecho + 1
            logo_apos_pergunta = inicio_trecho > 1 and _e_gatilho(valores_col_a[inicio_trecho - 1])

            if tamanho_trecho >= 4 and not logo_apos_pergunta:
                meio = (inicio_trecho + fim_trecho) // 2
                if meio - 1 >= 1:
                    # Trecho vazio livre, sem inserção/remoção nenhuma
                    # dentro dele — a posição em si já é uma linha
                    # original de verdade, remapeia direto.
                    quebras_planejadas.append(("livre", meio - 1))

            r = fim_trecho + 1
            continue

        r += 1

    if not operacoes and not quebras_planejadas:
        return 0, 0

    # ---- FASE 2: vetor de diferenças -> deslocamento acumulado -------
    delta = [0] * (max_row_original + 3)
    linha_removida = [False] * (max_row_original + 2)
    for tipo, pos, qtd in operacoes:
        if tipo == "inserir":
            delta[pos] += qtd
        else:
            for rr in range(pos, pos + qtd):
                linha_removida[rr] = True
            delta[pos + qtd] -= qtd

    deslocamento_acumulado = [0] * (max_row_original + 2)
    acumulado = 0
    for r in range(1, max_row_original + 1):
        acumulado += delta[r]
        deslocamento_acumulado[r] = acumulado

    def _nova_posicao(r_original):
        return r_original + deslocamento_acumulado[r_original]

    # ---- FASE 3: aplica tudo numa passada só --------------------------
    linhas_inseridas_total = sum(qtd for tipo, _pos, qtd in operacoes if tipo == "inserir")
    linhas_removidas_total = sum(qtd for tipo, _pos, qtd in operacoes if tipo == "remover")

    # Mesclagens: desfaz agora (registra os alvos pra recriar depois).
    # A RECRIAÇÃO (`_remesclar_rapido`) só acontece DEPOIS do bloco
    # "Células" abaixo — ver nota naquele bloco pra explicação do porquê.
    faixas_a_remesclar = []
    for rng in list(ws.merged_cells.ranges):
        dentro_de_removida = (
            rng.min_row <= max_row_original and linha_removida[rng.min_row]
        ) or (
            rng.max_row <= max_row_original and linha_removida[rng.max_row]
        )
        if not dentro_de_removida:
            novo_min = rng.min_row if rng.min_row > max_row_original else _nova_posicao(rng.min_row)
            novo_max = rng.max_row if rng.max_row > max_row_original else _nova_posicao(rng.max_row)
            faixas_a_remesclar.append((novo_min, rng.min_col, novo_max, rng.max_col))
        _desfazer_mesclagem_rapido(ws, rng)

    # Alturas de linha customizadas.
    alturas_a_mover = []
    for idx in list(ws.row_dimensions.keys()):
        dim = ws.row_dimensions.pop(idx)
        if idx <= max_row_original:
            if not linha_removida[idx]:
                alturas_a_mover.append((_nova_posicao(idx), dim))
        else:
            alturas_a_mover.append((idx + linhas_inseridas_total - linhas_removidas_total, dim))
    for nova_idx, dim in alturas_a_mover:
        dim.index = nova_idx
        ws.row_dimensions[nova_idx] = dim

    # Células: move valor + estilo de cada uma pra posição final, numa
    # única passada sobre o dicionário interno do openpyxl — é
    # exatamente o que `ws.insert_rows()`/`ws.delete_rows()` fazem por
    # baixo dos panos a cada chamada, só que aqui é feito uma vez só
    # para TODAS as operações combinadas.
    novas_cells = {}
    for (r, c), cel in ws._cells.items():
        if r > max_row_original:
            novas_cells[(r + linhas_inseridas_total - linhas_removidas_total, c)] = cel
            cel.row = r + linhas_inseridas_total - linhas_removidas_total
            continue
        if linha_removida[r]:
            continue
        nr = _nova_posicao(r)
        cel.row = nr
        novas_cells[(nr, c)] = cel
    ws._cells = novas_cells

    # SÓ AGORA recria as mesclagens nas posições finais. Motivo: a nova
    # posição de uma mesclagem (ex.: título que vai de A10:B10 pra
    # A13:B13) pode coincidir com a posição ANTIGA de uma célula de
    # dado que ainda não tinha sido movida — `_remesclar_rapido` limpa
    # o conteúdo das células não-âncora da faixa (igual o
    # `merge_cells()` oficial faz), então recriar a mesclagem ANTES de
    # mover as células apagava valores que só por coincidência de
    # numeração ainda estavam de passagem por aquela linha (bug real,
    # encontrado com arquivo de produção: "Regular", nomes de
    # candidatos e "Base" sumindo em várias tabelas — sempre a coluna
    # de valor, nunca o rótulo, porque a coluna A é a âncora do merge e
    # só a não-âncora é limpa). Fazendo a recriação depois, a célula
    # não-âncora que a mesclagem limpa já é, de fato, a célula vazia
    # que pertence à nova posição do título — não um valor de outra
    # linha que só passou por ali de raspão.
    for min_row, min_col, max_row_faixa, max_col in faixas_a_remesclar:
        _remesclar_rapido(ws, min_row=min_row, min_col=min_col, max_row=max_row_faixa, max_col=max_col)

    # Quebras de página já existentes: remapeia ou descarta as que
    # caíam dentro de uma faixa removida.
    quebras_remapeadas = []
    for quebra in ws.row_breaks.brk:
        if quebra.id <= max_row_original:
            if linha_removida[quebra.id]:
                continue
            quebra.id = _nova_posicao(quebra.id)
        else:
            quebra.id += linhas_inseridas_total - linhas_removidas_total
        quebras_remapeadas.append(quebra)
    ws.row_breaks.brk = quebras_remapeadas

    # Adiciona as quebras novas planejadas na Fase 1, já na posição final.
    quebras_adicionadas = 0
    ids_existentes = {b.id for b in ws.row_breaks.brk}
    for tipo_quebra, valor in quebras_planejadas:
        if tipo_quebra == "gatilho":
            # `valor` é a linha do gatilho — a posição final da quebra é
            # sempre 2 linhas depois de onde o próprio gatilho parar,
            # independente de o vão abaixo ter sido criado por inserção
            # (linha nova, sem posição original) ou já existir.
            pos_final = _nova_posicao(valor) + 2
        else:  # "livre" — trecho vazio já existente, remapeia direto
            if linha_removida[valor]:
                continue
            pos_final = _nova_posicao(valor)
        if pos_final not in ids_existentes:
            ws.row_breaks.append(Break(id=pos_final))
            ids_existentes.add(pos_final)
            quebras_adicionadas += 1

    # O Excel espera a lista de quebras em ordem CRESCENTE de linha, ou
    # descarta a lista inteira silenciosamente.
    ws.row_breaks.brk = sorted(ws.row_breaks.brk, key=lambda b: b.id)

    return quebras_adicionadas, linhas_inseridas_total


# =========================================================
# Código 10 — Ajusta o título da Renda
# =========================================================
def aplicar_codigo_10(ws):
    """Quebra 'Renda Média Mensal Domiciliar (em Salário Mínimo)' em 2 linhas.

    A altura da linha é CALCULADA a partir do texto já quebrado e da
    largura real da mesclagem (mesmo método do código 09), não mais um
    valor fixo de 30. Corrige um bug real: 30 dava conta só de 2 linhas
    visuais, mas "Renda Média Mensal Domiciliar" sozinho já não cabe
    numa linha só dentro da largura da coluna — o Excel quebra
    automaticamente em 2 linhas, e com a quebra manual do "(em Salário
    Mínimo)" dá 3 linhas visuais no total. Com a altura travada em 30,
    o texto estourava pra cima e sobrepunha a borda superior da célula
    (visualmente parecia uma borda mais fina/cortada bem em cima dessa
    coluna especificamente, e só dessa — as demais colunas do mesmo
    cabeçalho, com texto de 1 linha, cabiam sem problema)."""
    alvo = "Renda Média Mensal Domiciliar (em Salário Mínimo)"
    formatado = "Renda Média Mensal Domiciliar\n(em Salário Mínimo)"
    alterados = 0
    max_col = ws.max_column
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            if cel.value == alvo:
                cel.value = formatado
                cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                largura = largura_total_mesclagem(ws, cel)
                tamanho_fonte = cel.font.size or 10
                estim = estimar_altura_calculada(cel.value, largura, tamanho_fonte)
                # Piso de 45 (3 linhas de 15pt) em vez do valor genérico
                # da estimativa: confirmado por render real (LibreOffice)
                # que esse texto específico precisa de 3 linhas visuais
                # em larguras de coluna típicas — "Renda Média Mensal
                # Domiciliar" sozinho já quebra em 2 linhas automáticas,
                # mais a linha do "(em Salário Mínimo)" — mas a fórmula
                # genérica de caracteres-por-linha (calibrada pra
                # rótulos em geral, não pra este texto específico)
                # ainda estimava só 2 linhas, insuficiente.
                nova_altura = max(45.0, round(estim, 1))
                if ws.row_dimensions[r].height != nova_altura or ws.row_dimensions[r].height is None:
                    ws.row_dimensions[r].height = nova_altura
                alterados += 1
    return alterados


# =========================================================
# Código 14 — Cabeçalho/rodapé de impressão fixos
# =========================================================
def aplicar_codigo_14(wb):
    """Aplica cabeçalho fixo (aviso legal) e rodapé (nº de página) em todas as abas."""
    texto_cabecalho = (
        "NÃO ATUAMOS COM DIVULGAÇÃO DE PESQUISAS ELEITORAIS DESDE 2010.\n"
        "A PRESENTE PESQUISA NÃO ESTÁ REGISTRADA, SUA DIVULGAÇÃO ESTÁ SUJEITA A\n"
        "PENALIDADES PREVISTAS NA LEI ELEITORAL."
    )
    for ws in wb.worksheets:
        ws.oddHeader.left.text = f'&"DIN,Regular"&9{texto_cabecalho}'
        ws.oddHeader.center.text = ""
        ws.oddHeader.right.text = ""
        ws.oddFooter.left.text = ""
        ws.oddFooter.center.text = ""
        ws.oddFooter.right.text = '&"DIN,Regular"&9&P'
        ws.page_margins.header = 1.8 / 2.54
        ws.page_margins.top = 2.4 / 2.54
        ws.page_margins.footer = 0.5 / 2.54
        ws.page_margins.bottom = 1.5 / 2.54
    return len(wb.worksheets)


# =========================================================
# Código 16 — Índice / Sumário (páginas reais pelas quebras)
# =========================================================
def _ultima_linha_com_conteudo(ws):
    """Última linha com algum valor em qualquer coluna, varrendo de baixo pra cima."""
    for r in range(ws.max_row, 0, -1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                return r
    return 1


def _primeira_pagina(ws):
    """Espelha GetFirstPageNumber do VBA: só usa o número customizado se
    'usar número de primeira página' estiver realmente marcado; senão
    (Automático), a primeira página é 1."""
    try:
        if ws.page_setup.useFirstPageNumber and ws.page_setup.firstPageNumber:
            return int(ws.page_setup.firstPageNumber)
    except (TypeError, ValueError, AttributeError):
        pass
    return 1


def _pagina_da_linha(linha, quebras_ordenadas, primeira_pagina):
    """Conta quantas quebras de página já começaram até (e incluindo) `linha`."""
    cnt = sum(1 for inicio in quebras_ordenadas if inicio <= linha)
    return primeira_pagina + cnt


TERMOS_EXCLUIR_SUMARIO = ("CONTINUACAO", "PARA QUEM")


def _sumario_norm(texto):
    """Maiúsculo, sem acento — usado só pra comparar contra TOTAL/
    SEGMENTO/os termos excluídos do sumário, sem depender de acentuação
    exata."""
    import unicodedata
    t = str(texto or "").strip().upper()
    return "".join(ch for ch in unicodedata.normalize("NFD", t) if unicodedata.category(ch) != "Mn")


def aplicar_codigo_16(ws):
    """
    Cria (recriando se já existir) a aba "Sumário", listando o título de
    cada bloco (mesclagem na coluna A, exceto blocos "Total"/dentro de um
    "Total" até aparecer um "Segmento", "Pergunta..." e "*...") junto com
    o número de página real — calculado a partir das quebras de página
    manuais já presentes na planilha (ex.: inseridas pelo código 07 ou
    pelo código 15/InserirCapasResultados), não da posição na tela.

    Diferente do VBA original (que precisava alternar pra Page Break
    Preview e capturar `ws.HPageBreaks`), aqui as quebras já estão
    disponíveis diretamente em `ws.row_breaks` (mesmo mecanismo que o
    código 07 usa para inseri-las), então não há nada "visual" a simular.

    Formatação (a pedido do Lucas):
        - tudo em DIN Book 9;
        - sem linha de cabeçalho "Título"/"Página";
        - "RESULTADOS PELO TOTAL" vira a primeira linha, em negrito, com
          o número de página real dela (antes era só um texto solto sem
          página nenhuma);
        - "RESULTADOS PELOS SEGMENTOS" (já listada como qualquer outro
          título) fica em negrito;
        - linhas "Continuação" ou "Para quem..." (que às vezes acabam
          mescladas na coluna A perto de uma tabela dividida ou de uma
          Base reduzida) nunca entram na lista.

    Retorna a quantidade de linhas listadas no sumário.
    """
    wb = ws.parent

    if "Sumário" in wb.sheetnames:
        del wb["Sumário"]
    resumo = wb.create_sheet("Sumário")

    fonte_padrao = Font(name="DIN Book", size=9)
    fonte_negrito = Font(name="DIN Book", size=9, bold=True)

    ultima_linha = _ultima_linha_com_conteudo(ws)

    # Quebras de página já presentes: Break(id=n) => nova página começa na linha n+1
    quebras_inicio = sorted({b.id + 1 for b in ws.row_breaks.brk if b.id is not None})
    primeira_pagina = _primeira_pagina(ws)

    # Mescalgens da coluna A, indexadas pela linha do topo (só interessa
    # onde a mesclagem começa exatamente na linha i, igual ao VBA)
    merges_col_a = {
        rng.min_row: rng
        for rng in ws.merged_cells.ranges
        if rng.min_col <= 1 <= rng.max_col
    }

    entradas = []  # (titulo, pagina, negrito)
    titulos_vistos = set()
    ignorar = False
    pagina_total_capa = None

    i = 1
    while i <= ultima_linha:
        rng = merges_col_a.get(i)
        if rng is not None:
            altura_merge = rng.max_row - rng.min_row + 1
            titulo = str(ws.cell(row=rng.min_row, column=rng.min_col).value or "").strip()
            titulo_norm = _sumario_norm(titulo)

            if "TOTAL" in titulo_norm:
                ignorar = True
                if pagina_total_capa is None:
                    pagina_total_capa = _pagina_da_linha(i, quebras_inicio, primeira_pagina)
            elif "SEGMENTO" in titulo_norm:
                ignorar = False

            eh_termo_excluido = any(titulo_norm.startswith(t) for t in TERMOS_EXCLUIR_SUMARIO)

            if not ignorar and not eh_termo_excluido and titulo:
                if not titulo.startswith("Pergunta") and not titulo.startswith("*"):
                    if titulo not in titulos_vistos:
                        negrito = "SEGMENTO" in titulo_norm
                        pagina = _pagina_da_linha(i, quebras_inicio, primeira_pagina)
                        entradas.append((titulo, pagina, negrito))
                        titulos_vistos.add(titulo)

            i += altura_merge
        else:
            i += 1

    linhas_finais = []
    if pagina_total_capa is not None:
        linhas_finais.append(("RESULTADOS PELO TOTAL", pagina_total_capa, True))
    linhas_finais.extend(entradas)

    for idx, (titulo, pagina, negrito) in enumerate(linhas_finais, start=1):
        fonte = fonte_negrito if negrito else fonte_padrao
        c1 = resumo.cell(row=idx, column=1, value=titulo)
        c2 = resumo.cell(row=idx, column=2, value=pagina)
        c1.font = fonte
        c2.font = fonte

    # Larguras aproximadas (sem AutoFit real disponível fora do Excel)
    maior_titulo = max((len(t) for t, _, _ in linhas_finais), default=10)
    resumo.column_dimensions["A"].width = max(15, min(60, maior_titulo + 2))
    resumo.column_dimensions["B"].width = 10

    return len(linhas_finais)


def remover_termo_do_sumario(wb, termo):
    """
    Remove `termo` (ex.: '(Estimulada e Única)') do texto de cada linha
    da coluna A da aba "Sumário", colapsando o espaço duplo que sobra
    no lugar. Não mexe em mais nada (número de página, negrito etc.
    ficam intactos). Retorna quantas linhas foram alteradas.

    Não faz nada (retorna 0) se a aba "Sumário" ainda não existir —
    precisa rodar o código 16 primeiro.
    """
    if "Sumário" not in wb.sheetnames:
        return 0
    resumo = wb["Sumário"]
    alterados = 0
    for r in range(1, resumo.max_row + 1):
        cel = resumo.cell(row=r, column=1)
        if isinstance(cel.value, str) and termo in cel.value:
            novo = cel.value.replace(termo, "")
            novo = " ".join(novo.split())
            cel.value = novo
            alterados += 1
    return alterados
