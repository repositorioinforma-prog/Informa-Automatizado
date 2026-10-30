"""
Exclusão de colunas por Base / Base reduzida — com regras configuráveis.

Generaliza o que o código 05 e o Processador de Base Reduzida fazem (esses só
olham a linha "Base reduzida" e só aceitam "menor que N"): aqui a pessoa
escolhe QUAL linha de base vale (Base, Base reduzida ou as duas), o valor de
corte, o operador (< ou ≤) e, se quiser, restringe a regra a certos grupos de
segmento (ex.: só "Região"). Uma coluna é apagada se QUALQUER regra a pegar.

Fluxo em três passos, todos sem Streamlit (testáveis isoladamente):

    scan  = escanear_planilha(ws)                      # lê tabelas, cabeçalhos e bases
    plano = planejar_exclusoes(scan, regras, ...)      # decide o que apagar (só dados)
    resumo = aplicar_exclusoes(ws, plano)              # apaga na planilha

Separar "planejar" de "aplicar" permite mostrar a prévia na tela, em tempo
real, sem recarregar o Excel a cada clique.

Como a formatação é preservada
------------------------------
- Cada tabela é deslocada só nas PRÓPRIAS linhas (tabelas empilhadas podem
  perder colunas diferentes) e de uma vez só, não coluna por coluna.
- Valor E estilo completo da célula (fonte, preenchimento, borda,
  alinhamento, formato numérico) andam juntos.
- Mesclagens de cabeçalho encolhem para cobrir só as colunas que sobraram;
  se a âncora da mesclagem (a 1ª coluna do grupo) for apagada, o texto e o
  estilo dela passam para a nova âncora. Grupo que perde todas as colunas some.
- Se a última coluna da tabela sai, a nova última coluna herda a borda
  direita (estilo e cor) que a antiga tinha — o contorno da tabela não abre.
- As colunas que sobram vazias à direita são removidas (não deixam
  preenchimento/borda para trás).
- Larguras de coluna acompanham o deslocamento quando isso é seguro (todas as
  tabelas da aba apagam exatamente as mesmas colunas e não há outro conteúdo
  na aba); senão ficam como estão, pois as tabelas dividem as mesmas colunas.

A coluna A (rótulos) nunca é apagada. A coluna B (Total) fica protegida por
padrão.
"""
from __future__ import annotations

import re
import unicodedata
from copy import copy
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string, get_column_letter

from core.planilha_utils import (
    _desfazer_mesclagem_rapido,
    _fechar_borda_direita,
    _remesclar_rapido,
)

TIPO_BASE = "BASE"
TIPO_BASE_REDUZIDA = "BASE REDUZIDA"
TIPOS_DE_BASE = (TIPO_BASE, TIPO_BASE_REDUZIDA)

OPERADORES = ("<", "<=")


# ---------------------------------------------------------------------------
# Regras
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Regra:
    """
    limite    : valor de corte digitado.
    linhas    : quais linhas de base a regra consulta (TIPO_BASE / TIPO_BASE_REDUZIDA).
    operador  : "<" (menor que) ou "<=" (menor ou igual a).
    grupos    : grupos de segmento a que a regra se restringe; vazio = todos.
    """
    limite: float
    linhas: tuple = TIPOS_DE_BASE
    operador: str = "<"
    grupos: tuple = ()

    def __post_init__(self):
        if self.operador not in OPERADORES:
            raise ValueError(f"operador inválido: {self.operador!r} (use '<' ou '<=')")
        if not self.linhas or any(x not in TIPOS_DE_BASE for x in self.linhas):
            raise ValueError(f"linhas inválidas: {self.linhas!r}")

    def casa_valor(self, valor: float) -> bool:
        return valor < self.limite if self.operador == "<" else valor <= self.limite

    def casa_grupo(self, grupo_efetivo: str) -> bool:
        if not self.grupos:
            return True
        return _chave(grupo_efetivo) in {_chave(g) for g in self.grupos}

    def descricao(self) -> str:
        if set(self.linhas) == set(TIPOS_DE_BASE):
            alvo = "Base ou Base reduzida"
        else:
            alvo = "Base reduzida" if self.linhas[0] == TIPO_BASE_REDUZIDA else "Base"
        n = int(self.limite) if float(self.limite).is_integer() else self.limite
        texto = f"{alvo} {self.operador.replace('<=', '≤')} {n}"
        if self.grupos:
            texto += f" (só em: {', '.join(self.grupos)})"
        return texto


# ---------------------------------------------------------------------------
# Texto e número
# ---------------------------------------------------------------------------
def _chave(texto) -> str:
    """Minúsculas, sem acento, espaços colapsados — comparação tolerante."""
    s = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def _texto(valor) -> str:
    return "" if valor is None else str(valor).strip()


def _limpar_rotulo(valor) -> str:
    return re.sub(r"\s+", " ", str(valor).replace("-\n", "")).strip()


_RE_NUMERO = re.compile(r"^\s*(\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:[.,]\d+)?)\s*\**\s*$")
_RE_MILHAR = re.compile(r"^\d{1,3}(?:\.\d{3})+(?:,\d+)?$")


def numero_da_base(valor):
    """
    Número de uma célula de base, ou None se não for número. Aceita 604, 604.0,
    '25*' (base pequena já marcada pelo código 08), '12,5' e '1.234' (milhar).
    """
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor) if valor == valor else None  # descarta NaN
    m = _RE_NUMERO.match(str(valor))
    if not m:
        return None
    t = m.group(1)
    t = t.replace(".", "").replace(",", ".") if _RE_MILHAR.match(t) else t.replace(",", ".")
    return float(t)


def tipo_de_base(valor_coluna_a):
    """'BASE' / 'BASE REDUZIDA' se o rótulo da coluna A for isso (tolerante a
    acento, caixa e ':' no fim), senão None."""
    if valor_coluna_a is None:
        return None
    chave = _chave(valor_coluna_a).rstrip(":").strip()
    if chave == "base":
        return TIPO_BASE
    if chave == "base reduzida":
        return TIPO_BASE_REDUZIDA
    return None


# ---------------------------------------------------------------------------
# 1) Leitura: tabelas, cabeçalhos e linhas de base
# ---------------------------------------------------------------------------
def _linha_vazia(linha) -> bool:
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in linha)


def escanear_planilha(ws) -> dict:
    """
    Lê a aba e devolve:
      {"aba", "tabelas": [...], "linhas_fora": int, "n_base": int,
       "n_base_reduzida": int, "rotulos_base_desconhecidos": [...]}

    Tabela = trecho de linhas contíguas (separadas por linha em branco) que
    tem linha 'Base'/'Base reduzida'; se um mesmo trecho traz duas tabelas
    coladas (um título novo depois de uma base), ele é dividido em duas.
    Cada tabela traz o rótulo de grupo/segmento de cada coluna e os valores
    de suas linhas de base (ver `planejar_exclusoes`).
    """
    n_linhas, n_colunas = ws.max_row, ws.max_column
    grade = [
        tuple(l) for l in ws.iter_rows(
            min_row=1, max_row=n_linhas, min_col=1, max_col=n_colunas, values_only=True,
        )
    ]

    # mesclagens: texto de cada célula coberta + linhas de título mescladas na A
    valor_mesclado = {}
    linhas_titulo_mescladas = set()
    ultima_col_mesclada_na_linha = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row > n_linhas or rng.min_col > n_colunas:
            continue
        ancora = grade[rng.min_row - 1][rng.min_col - 1]
        if ancora is not None:
            for r in range(rng.min_row, min(rng.max_row, n_linhas) + 1):
                for c in range(rng.min_col, min(rng.max_col, n_colunas) + 1):
                    valor_mesclado[(r, c)] = ancora
        for r in range(rng.min_row, min(rng.max_row, n_linhas) + 1):
            if rng.max_col > ultima_col_mesclada_na_linha.get(r, 0):
                ultima_col_mesclada_na_linha[r] = rng.max_col
        if rng.min_row == rng.max_row and rng.min_col == 1 and rng.max_col > 1 and _texto(ancora):
            linhas_titulo_mescladas.add(rng.min_row)

    def celula(r, c):
        v = grade[r - 1][c - 1]
        return valor_mesclado.get((r, c)) if v is None else v

    def eh_titulo(r):
        a = _texto(grade[r - 1][0])
        if not a:
            return False
        chave = _chave(a)
        if chave.startswith("pergunta"):
            return False
        return r in linhas_titulo_mescladas or chave.startswith("titulo")

    vazia = [_linha_vazia(l) for l in grade]

    tabelas = []
    linhas_fora = 0
    desconhecidos = []
    n_base = n_base_reduzida = 0

    r = 0
    while r < n_linhas:
        if vazia[r]:
            r += 1
            continue
        ini = r + 1
        while r < n_linhas and not vazia[r]:
            r += 1
        fim = r  # 1-based, inclusive

        bases = []
        for rr in range(ini, fim + 1):
            tipo = tipo_de_base(grade[rr - 1][0])
            if tipo:
                bases.append((rr, tipo))
            else:
                a = _texto(grade[rr - 1][0])
                if a and _chave(a).startswith("base") and a not in desconhecidos:
                    desconhecidos.append(a)
        if not bases:
            linhas_fora += fim - ini + 1
            continue

        # divide o trecho em tabelas: um título novo entre duas bases abre outra
        inicios = [ini]
        for (ant, _), (prox, _) in zip(bases, bases[1:]):
            titulos = [x for x in range(ant + 1, prox) if eh_titulo(x)]
            if titulos:
                inicios.append(titulos[-1])
        fins = [x - 1 for x in inicios[1:]] + [fim]

        for t_ini, t_fim in zip(inicios, fins):
            bases_t = [(rr, tp) for rr, tp in bases if t_ini <= rr <= t_fim]
            primeira_base = bases_t[0][0]

            titulos = [x for x in range(t_ini, primeira_base) if eh_titulo(x)]
            linha_titulo = titulos[-1] if titulos else None
            titulo = "(sem título)"
            if linha_titulo:
                titulo = re.sub(r"^\s*titulo\s*:\s*", "", _texto(grade[linha_titulo - 1][0]),
                                flags=re.IGNORECASE) or "(sem título)"

            # primeira linha de dados: rótulo na A e algum número depois da A
            inicio_dados = primeira_base
            for x in range((linha_titulo or t_ini - 1) + 1, primeira_base + 1):
                a = _texto(grade[x - 1][0])
                if (a and not _chave(a).startswith("pergunta") and not eh_titulo(x)
                        and any(numero_da_base(v) is not None for v in grade[x - 1][1:])):
                    inicio_dados = x
                    break
            linhas_cabecalho = range((linha_titulo or t_ini - 1) + 1, inicio_dados)

            largura = 1
            for x in range(t_ini, t_fim + 1):
                ult = 0
                for c in range(n_colunas, 0, -1):
                    v = grade[x - 1][c - 1]
                    if v is not None and not (isinstance(v, str) and not v.strip()):
                        ult = c
                        break
                largura = max(largura, ult, min(ultima_col_mesclada_na_linha.get(x, 0), n_colunas))

            colunas = {}
            for c in range(2, largura + 1):
                rotulos = []
                for x in linhas_cabecalho:
                    v = celula(x, c)
                    if v is None or not _texto(v):
                        continue
                    s = _limpar_rotulo(v)
                    if not rotulos or rotulos[-1] != s:
                        rotulos.append(s)
                colunas[c] = {
                    "grupo": rotulos[0] if len(rotulos) > 1 else "",
                    "segmento": rotulos[-1] if rotulos else "",
                }

            bases_info = []
            for rr, tipo in bases_t:
                valores, vazias, formulas = {}, [], 0
                for c in range(2, largura + 1):
                    bruto = grade[rr - 1][c - 1]
                    if bruto is None or (isinstance(bruto, str) and not bruto.strip()):
                        vazias.append(c)
                        valores[c] = None
                        continue
                    if isinstance(bruto, str) and bruto.lstrip().startswith("="):
                        formulas += 1
                    valores[c] = numero_da_base(bruto)
                bases_info.append({
                    "linha": rr, "tipo": tipo, "valores": valores, "vazias": vazias, "formulas": formulas,
                })
                if tipo == TIPO_BASE:
                    n_base += 1
                else:
                    n_base_reduzida += 1

            tabelas.append({
                "aba": ws.title, "titulo": titulo, "linha_titulo": linha_titulo,
                "linha_ini": t_ini, "linha_fim": t_fim, "largura": largura,
                "colunas": colunas, "bases": bases_info,
            })

    return {
        "aba": ws.title, "tabelas": tabelas, "linhas_fora": linhas_fora,
        "n_base": n_base, "n_base_reduzida": n_base_reduzida,
        "rotulos_base_desconhecidos": desconhecidos[:10],
    }


def grupo_efetivo(info_coluna: dict) -> str:
    """Grupo da coluna; se o cabeçalho não tem grupo, o próprio segmento."""
    return info_coluna.get("grupo") or info_coluna.get("segmento") or ""


def _coluna_b_e_total(tabela: dict) -> bool:
    """A coluna B é o Total? Sim se o cabeçalho diz 'total' ou está sem rótulo."""
    info = tabela["colunas"].get(2)
    if info is None:
        return False
    rotulo = _chave(grupo_efetivo(info))
    return rotulo == "" or "total" in rotulo


def listar_grupos(scans) -> list:
    """Grupos de segmento encontrados (ordem de aparição, sem repetir),
    sem o Total — alimenta o filtro por grupo da tela."""
    vistos = {}
    for scan in scans:
        for t in scan["tabelas"]:
            col_min = 3 if _coluna_b_e_total(t) else 2
            for c in range(col_min, t["largura"] + 1):
                g = grupo_efetivo(t["colunas"].get(c, {}))
                if g and _chave(g) not in vistos:
                    vistos[_chave(g)] = g
    return list(vistos.values())


# ---------------------------------------------------------------------------
# 2) Plano: o que apagar (só dados — nada é alterado na planilha)
# ---------------------------------------------------------------------------
def planejar_exclusoes(scan, regras, proteger_total=True, vazias_como_zero=False) -> dict:
    """
    Aplica as regras sobre o scan e devolve o plano:
      {"aba", "itens": [{"tabela", "excluir": {col: info}}], "total_colunas",
       "tabelas_afetadas", "tabelas_sem_segmentos", "larguras_deslocaveis"}

    proteger_total   : não apaga a coluna B quando ela é o Total.
    vazias_como_zero : célula de base vazia conta como 0 (senão é ignorada).
    """
    regras = list(regras)
    itens = []
    for t in scan["tabelas"]:
        col_min = 3 if (proteger_total and _coluna_b_e_total(t)) else 2
        excluir = {}
        for base in t["bases"]:
            vazias = set(base["vazias"])
            for idx, regra in enumerate(regras, start=1):
                if base["tipo"] not in regra.linhas:
                    continue
                for c in range(col_min, t["largura"] + 1):
                    if c in excluir:
                        continue
                    n = base["valores"].get(c)
                    if n is None:
                        if vazias_como_zero and c in vazias:
                            n = 0.0
                        else:
                            continue
                    info = t["colunas"].get(c, {})
                    if not regra.casa_valor(n) or not regra.casa_grupo(grupo_efetivo(info)):
                        continue
                    excluir[c] = {
                        "coluna": c, "letra": get_column_letter(c),
                        "grupo": info.get("grupo", ""), "segmento": info.get("segmento", ""),
                        "valor": n, "linha_base": base["linha"], "tipo": base["tipo"],
                        "regra": idx, "regra_descricao": regra.descricao(),
                    }
        itens.append({"tabela": t, "excluir": excluir})

    afetados = [i for i in itens if i["excluir"]]
    sem_segmentos = [
        i["tabela"] for i in afetados
        if i["tabela"]["largura"] >= 3
        and all(c in i["excluir"] for c in range(3, i["tabela"]["largura"] + 1))
    ]

    conjuntos = {frozenset(i["excluir"]) for i in itens}
    larguras_deslocaveis = (
        bool(itens) and scan["linhas_fora"] == 0 and len(conjuntos) == 1 and bool(next(iter(conjuntos)))
    )

    return {
        "aba": scan["aba"], "itens": itens,
        "total_colunas": sum(len(i["excluir"]) for i in itens),
        "tabelas_afetadas": len(afetados),
        "tabelas_sem_segmentos": sem_segmentos,
        "larguras_deslocaveis": larguras_deslocaveis,
    }


def linhas_do_plano(plano) -> list:
    """Uma linha por coluna a apagar, pronta para virar tabela na tela."""
    linhas = []
    for item in plano["itens"]:
        t = item["tabela"]
        for c in sorted(item["excluir"]):
            e = item["excluir"][c]
            linhas.append({
                "Aba": plano["aba"],
                "Tabela": t["titulo"],
                "Linha do título": t["linha_titulo"] or t["linha_ini"],
                "Grupo": e["grupo"],
                "Segmento": e["segmento"] or f"(coluna {e['letra']})",
                "Coluna": e["letra"],
                "Linha de base": "Base reduzida" if e["tipo"] == TIPO_BASE_REDUZIDA else "Base",
                "Valor da base": int(e["valor"]) if float(e["valor"]).is_integer() else e["valor"],
                "Regra": f"{e['regra']}: {e['regra_descricao']}",
            })
    return linhas


# ---------------------------------------------------------------------------
# 3) Aplicação na planilha
# ---------------------------------------------------------------------------
def _excluir_colunas_na_faixa(ws, linha_ini, linha_fim, excluir, max_col, largura):
    """
    Apaga as colunas `excluir` (índices 1-based) só nas linhas
    [linha_ini, linha_fim], deslocando o resto para a esquerda (valor +
    estilo), refazendo as mesclagens e fechando o contorno direito.
    Devolve quantas células com fórmula foram deslocadas (fórmulas não são
    reescritas — é só para avisar).
    """
    excluir = set(excluir)
    sobreviventes = [c for c in range(1, max_col + 1) if c not in excluir]
    n_sobreviventes = len(sobreviventes)

    # 1) se a última coluna da tabela sai, lembra a borda direita dela (linha a linha)
    contorno = {}
    if largura in excluir:
        for r in range(linha_ini, linha_fim + 1):
            cel = ws._cells.get((r, largura))
            lado = cel.border.right if cel is not None else None
            if lado is not None and lado.style:
                contorno.setdefault(lado.style, []).append((r, lado.color))

    # 2) guarda as mesclagens da faixa (texto + estilo da âncora) e desfaz
    faixas = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row <= linha_fim and rng.max_row >= linha_ini:
            ancora = ws._cells.get((rng.min_row, rng.min_col))
            faixas.append((
                rng.min_row, rng.min_col, rng.max_row, rng.max_col,
                ancora._value if ancora is not None else None,
                ancora.data_type if ancora is not None else "n",
                copy(ancora._style) if ancora is not None else None,
            ))
            _desfazer_mesclagem_rapido(ws, rng)

    # 3) desloca as células (da esquerda para a direita: destino <= origem)
    formulas = 0
    for r in range(linha_ini, linha_fim + 1):
        for novo, orig in enumerate(sobreviventes, start=1):
            if novo == orig:
                continue
            origem = ws._cells.get((r, orig))
            destino = ws._cells.get((r, novo))
            if origem is None:
                if destino is not None:
                    del ws._cells[(r, novo)]
                continue
            if destino is None:
                destino = ws.cell(row=r, column=novo)
            destino._value = origem._value
            destino.data_type = origem.data_type
            destino._style = copy(origem._style)
            if origem.data_type == "f":
                formulas += 1
        # colunas que sobraram à direita: somem (sem preenchimento/borda órfãos)
        for c in range(n_sobreviventes + 1, max_col + 1):
            ws._cells.pop((r, c), None)

    # 4) refaz as mesclagens encolhidas; a âncora herda texto e estilo originais
    for min_row, min_col, max_row, max_c, valor, tipo_dado, estilo in faixas:
        antes = sum(1 for c in excluir if c < min_col)
        dentro = sum(1 for c in excluir if min_col <= c <= max_c)
        novo_min, novo_max = min_col - antes, max_c - antes - dentro
        if novo_max < novo_min:
            continue  # o grupo inteiro foi apagado
        ancora = ws.cell(row=min_row, column=novo_min)
        ancora._value = valor
        ancora.data_type = tipo_dado
        if estilo is not None:
            ancora._style = copy(estilo)
        # mesclagem puramente vertical também precisa voltar (max_row > min_row)
        if novo_max > novo_min or max_row > min_row:
            _remesclar_rapido(ws, min_row=min_row, min_col=novo_min, max_row=max_row, max_col=novo_max)

    # 5) contorno direito da tabela
    if contorno:
        nova_ultima = largura - len(excluir)
        for estilo, linhas_cor in contorno.items():
            _fechar_borda_direita(ws, nova_ultima, linhas_cor, estilo)

    return formulas


def _deslocar_larguras(ws, excluir, max_col):
    """Acompanha o deslocamento nas larguras (e ocultação/agrupamento) das
    colunas. Só chamada quando todas as tabelas da aba apagam as mesmas colunas."""
    from openpyxl.worksheet.dimensions import ColumnDimension

    excluir = set(excluir)
    sobreviventes = [c for c in range(1, max_col + 1) if c not in excluir]
    dims = ws.column_dimensions

    por_coluna = {}
    for letra, cd in list(dims.items()):
        lo = cd.min or column_index_from_string(letra)
        hi = cd.max or lo
        for c in range(lo, min(hi, max_col) + 1):
            por_coluna[c] = cd
        if hi <= max_col:
            del dims[letra]
        elif lo <= max_col:  # grupo que passa de max_col: mantém só a cauda
            del dims[letra]
            cd.min, cd.max, cd.index = max_col + 1, hi, get_column_letter(max_col + 1)
            dims[cd.index] = cd

    def clonar(cd, c):
        novo = ColumnDimension(
            ws, index=get_column_letter(c), width=cd.width, hidden=cd.hidden,
            outlineLevel=cd.outlineLevel, collapsed=cd.collapsed, bestFit=cd.bestFit, min=c, max=c,
        )
        novo._style = copy(cd._style)
        return novo

    for novo_idx, orig in enumerate(sobreviventes, start=1):
        cd = por_coluna.get(orig)
        if cd is not None:
            dims[get_column_letter(novo_idx)] = clonar(cd, novo_idx)
    for c in range(len(sobreviventes) + 1, max_col + 1):  # colunas agora vazias: mantêm o que tinham
        cd = por_coluna.get(c)
        if cd is not None:
            dims[get_column_letter(c)] = clonar(cd, c)


def aplicar_exclusoes(ws, plano, ajustar_larguras=True) -> dict:
    """
    Executa o plano (de `planejar_exclusoes`) na aba `ws`. Devolve
      {"colunas_excluidas", "tabelas_afetadas", "larguras_ajustadas", "avisos"}.
    """
    resumo = {"colunas_excluidas": 0, "tabelas_afetadas": 0, "larguras_ajustadas": False, "avisos": []}
    itens = [i for i in plano["itens"] if i["excluir"]]
    if not itens:
        return resumo

    max_col = ws.max_column
    formulas = 0
    for item in itens:
        t = item["tabela"]
        formulas += _excluir_colunas_na_faixa(
            ws, t["linha_ini"], t["linha_fim"], sorted(item["excluir"]), max_col, t["largura"],
        )
        resumo["colunas_excluidas"] += len(item["excluir"])
        resumo["tabelas_afetadas"] += 1

    if ajustar_larguras and plano["larguras_deslocaveis"]:
        _deslocar_larguras(ws, itens[0]["excluir"], max_col)
        resumo["larguras_ajustadas"] = True

    if formulas:
        resumo["avisos"].append(
            f"{formulas} célula(s) com fórmula foram deslocadas sem reescrever as referências — confira."
        )
    if len(ws.conditional_formatting):
        resumo["avisos"].append("A aba tem formatação condicional, que não é deslocada junto com as colunas.")
    if ws.data_validations is not None and ws.data_validations.dataValidation:
        resumo["avisos"].append("A aba tem validação de dados, que não é deslocada junto com as colunas.")
    return resumo


def excluir_colunas_por_regras(ws, regras, proteger_total=True, vazias_como_zero=False,
                               ajustar_larguras=True) -> dict:
    """Atalho: escaneia, planeja e aplica numa chamada (scripts e testes)."""
    plano = planejar_exclusoes(escanear_planilha(ws), regras, proteger_total, vazias_como_zero)
    resumo = aplicar_exclusoes(ws, plano, ajustar_larguras)
    resumo["plano"] = plano
    return resumo
