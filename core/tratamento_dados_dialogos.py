"""
Tratamento de Dados — partes puras (sem Streamlit) da tela no estilo SPSS.

Reúne a organização dos comandos em faixas por tarefa, os textos de ajuda de cada janela e as conversões das grades
(tabela de regras do Recode, rótulos de valores, valores ausentes) e monta as tabelas HTML do Data View e do Variable View para o
texto/parâmetros que o núcleo (`tratamento_dados_math`) entende. Tudo aqui é testável sem tela.
"""
from __future__ import annotations

import html
import math
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

from core.tratamento_dados_math import ErroTratamento, tabela_variaveis

# --------------------------------------------------------------------------------------- faixas de comandos
# A tela organiza os comandos por TAREFA, sempre visíveis, em abas (faixas) com grupos:
#   faixa -> grupos -> comandos (id da janela, ícone Material, rótulo, descrição curta)
Comando = Tuple[str, str, str, str]
FAIXAS: List[Tuple[str, List[Tuple[str, List[Comando]]]]] = [
    ("Variáveis", [
        ("Criar e transformar", [
            ("compute", "calculate", "Calcular variável", "Cria uma variável por fórmula (COMPUTE / IF)"),
            ("recode_outras", "swap_horiz", "Recodificar em outras variáveis", "Troca códigos e grava numa variável nova"),
            ("recode_mesmas", "edit_note", "Recodificar nas mesmas variáveis", "Troca códigos na própria variável"),
        ]),
        ("Organizar", [
            ("renomear_variavel", "drive_file_rename_outline", "Renomear variável", "Muda o nome de uma variável"),
            ("excluir_variaveis", "delete", "Excluir variáveis", "Remove colunas do banco"),
        ]),
    ]),
    ("Casos", [
        ("Filtrar", [
            ("selecionar_casos", "filter_alt", "Selecionar casos", "Mantém ou exclui linhas por uma condição"),
            ("remover_duplicados", "difference", "Remover casos duplicados", "Deixa uma linha por chave (ex.: telefone)"),
        ]),
        ("Organizar", [
            ("ordenar_casos", "sort", "Ordenar casos", "Reordena as linhas por até 3 variáveis"),
        ]),
        ("Resumir", [
            ("agregar", "functions", "Agregar dados", "Média, soma ou contagem por grupo"),
        ]),
    ]),
    ("Rótulos e ausentes", [
        ("Descrever as variáveis", [
            ("rotular_variavel", "label", "Rótulo da variável", "Texto descritivo da variável"),
            ("rotular_valores", "format_list_numbered", "Rótulos dos valores", "Nome de cada código (1 = Sim)"),
            ("definir_medida", "straighten", "Nível de medida", "Nominal, ordinal ou escala"),
        ]),
        ("Valores ausentes", [
            ("definir_missing", "block", "Valores ausentes", "Códigos como 98 / 99 fora das contas"),
        ]),
    ]),
    ("Juntar bancos", [
        ("Combinar com outro .sav", [
            ("adicionar_casos", "playlist_add", "Adicionar casos", "Empilha as linhas de outro banco"),
            ("adicionar_variaveis", "view_column", "Adicionar variáveis", "Traz colunas de outro banco por uma chave"),
        ]),
    ]),
    ("Arquivo", [
        ("Salvar e exportar", [
            ("salvar_sav", "save", "Salvar como .sav", "Grava o banco tratado"),
            ("exportar", "file_download", "Exportar para Excel / CSV", "Dados e uma aba com as variáveis"),
            ("salvar_sintaxe", "code", "Salvar syntax (.sps)", "Todos os passos, para rodar no SPSS"),
        ]),
    ]),
    ("Ajuda", [
        ("Consultar", [
            ("ajuda_expressoes", "menu_book", "Funções e operadores", "O que dá para usar nas expressões"),
            ("ajuda_geral", "help", "Como usar este módulo", "Visão geral da tela"),
        ]),
    ]),
]

ACOES_DIRETAS = {"desfazer", "desfazer_tudo"}          # botões de desfazer: não abrem janela

# Botões da Visualização de variáveis: (rótulo, id da janela) — abrem a janela já na variável escolhida
ACOES_DA_VARIAVEL: List[Tuple[str, str]] = [
    ("Nome", "renomear_variavel"),
    ("Rótulo", "rotular_variavel"),
    ("Valores", "rotular_valores"),
    ("Ausentes", "definir_missing"),
    ("Medida", "definir_medida"),
]

# id da janela -> operação do núcleo (None = janela que não altera o banco)
OPERACAO_DA_JANELA: Dict[str, Optional[str]] = {
    "compute": "compute",
    "recode_mesmas": "recode", "recode_outras": "recode",
    "selecionar_casos": "selecionar_casos", "ordenar_casos": "ordenar_casos",
    "remover_duplicados": "remover_duplicados", "agregar": "agregar",
    "excluir_variaveis": "excluir_variaveis", "renomear_variavel": "renomear_variavel",
    "rotular_variavel": "rotular_variavel", "rotular_valores": "rotular_valores",
    "definir_missing": "definir_missing", "definir_medida": "definir_medida",
    "adicionar_casos": "adicionar_casos", "adicionar_variaveis": "adicionar_variaveis",
    "salvar_sav": None, "exportar": None, "salvar_sintaxe": None,
    "ajuda_expressoes": None, "ajuda_geral": None,
}

VISOES = ("Visualização de dados", "Visualização de variáveis", "Syntax")


def ids_dos_comandos() -> List[str]:
    """Todos os ids de comando das faixas, na ordem em que aparecem."""
    return [c[0] for _, grupos in FAIXAS for _, comandos in grupos for c in comandos]


# --------------------------------------------------------------------------------------- ajuda
AJUDA: Dict[str, str] = {
    "compute": "Cria uma variável nova (ou muda uma existente) a partir de uma fórmula — é o COMPUTE do SPSS. "
               "Exemplos: `Idade * 12` · `MEAN(Q1 TO Q5)`. A condição em “Se…” funciona como o IF: "
               "só os casos que a satisfazem recebem o valor.",
    "recode_mesmas": "Troca códigos dentro da própria variável (RECODE sem INTO). Os valores originais são "
                     "perdidos — se quiser guardar a original, use “Recodificar em outras variáveis”.",
    "recode_outras": "Cria uma variável nova com os códigos trocados (RECODE … INTO). Exemplo: de `Idade` para "
                     "`Faixa` com `Do menor valor até 29 → 1`, `Intervalo 30 a 49 → 2`, `Todos os outros → 3`. "
                     "A primeira regra que casar vale. Sem “Todos os outros”, o que não foi citado fica vazio.",
    "selecionar_casos": "Mantém só as linhas que atendem à condição (ou exclui as que atendem) — é o SELECT IF. "
                        "Os casos não selecionados saem do banco. Exemplo: `Cidade = 'Rio' AND Idade >= 18`.",
    "ordenar_casos": "Reordena as linhas por até 3 variáveis (SORT CASES). Valores ausentes ficam no início "
                     "na ordem crescente, como no SPSS.",
    "remover_duplicados": "Deixa uma linha por chave (ex.: telefone ou ID), guardando a primeira ou a última.",
    "agregar": "Resume por grupo (AGGREGATE): média, soma, contagem etc. Pode acrescentar o resumo a cada caso "
               "(ex.: idade média da cidade em cada linha) ou substituir o banco por uma linha por grupo.",
    "adicionar_casos": "Empilha as linhas de um segundo .sav embaixo das do banco ativo (ADD FILES), "
                       "casando as variáveis pelo nome.",
    "adicionar_variaveis": "Traz colunas de um segundo .sav ligando as linhas por uma chave comum, como `ID` "
                           "(MATCH FILES).",
    "rotular_variavel": "O rótulo é o texto descritivo da variável (VARIABLE LABELS). Em branco, remove.",
    "rotular_valores": "Dá nome aos códigos (VALUE LABELS): `1 = Baixa`, `2 = Média`… A grade já vem com os "
                       "rótulos atuais da variável de referência.",
    "definir_missing": "Define quais códigos (ex.: 98, 99) são “ausentes do usuário” e ficam fora das contas "
                       "(MISSING VALUES). O SPSS aceita até 3 valores, ou 1 faixa + 1 valor.",
    "definir_medida": "Nominal (categorias), ordinal (categorias com ordem) ou escala (números contínuos).",
    "renomear_variavel": "Muda o nome da variável. As fórmulas já gravadas no histórico continuam valendo "
                         "porque o histórico é refeito do início.",
    "excluir_variaveis": "Remove colunas do banco (DELETE VARIABLES).",
    "salvar_sav": "Grava o banco tratado em um novo .sav, com rótulos, missings e nível de medida.",
    "exportar": "Exporta os dados para Excel (com uma aba de variáveis) ou CSV, com códigos ou rótulos.",
    "salvar_sintaxe": "Baixa o arquivo .sps com todos os passos, para rodar no SPSS ou guardar como registro.",
}

AJUDA_EXPRESSOES = """
**Operadores:** `+ - * / **` · comparação `= <> < > <= >=` (ou `EQ NE LT GT LE GE`) · lógicos `AND OR NOT` (ou `& | ~`)

**Funções numéricas:** `SUM(a,b,c)` `MEAN(a,b,c)` `MIN` `MAX` `SD` `MEDIAN` `NVALID` — aceitam `Q1 TO Q5` e o mínimo
de válidos (`MEAN.2(...)`) · `ABS` `SQRT` `EXP` `LN` `LG10` `RND` `TRUNC` `MOD(x,y)`

**Testes:** `ANY(x, 1, 2)` · `RANGE(x, 1, 5, 9, 9)` · `MISSING(x)` (ausente ou missing de usuário) ·
`SYSMIS(x)` · `VALUE(x)` (valor bruto, ignorando o missing de usuário)

**Texto:** `CONCAT(a, '-', b)` `UPCASE` `LOWER` `LTRIM` `RTRIM` `LENGTH` `SUBSTR(x, pos, tam)` `NUMBER(x, F8)`

**Outros:** `$CASENUM` (nº do caso) · `$SYSMIS`

Nomes de variável não diferenciam maiúsculas de minúsculas. Valor ausente em uma conta dá ausente; em uma
comparação, o caso não é nem verdadeiro nem falso (e por isso fica de fora do IF / SELECT IF), como no SPSS.
"""

AJUDA_GERAL = """
**Os comandos ficam em abas, organizados por tarefa.**

- **Variáveis** — criar (calcular, recodificar) e organizar (renomear, excluir).
- **Casos** — filtrar (selecionar, remover duplicados), ordenar e resumir (agregar).
- **Rótulos e ausentes** — rótulo da variável, rótulos dos valores, nível de medida e valores ausentes.
- **Juntar bancos** — adicionar casos (empilhar) ou variáveis (acrescentar colunas) de outro .sav.
- **Arquivo** — salvar o .sav, exportar para Excel/CSV e salvar a syntax (.sps).
- Clique em um comando para abrir sua **janela**, com **OK** (aplica), **Ver syntax** (mostra o comando equivalente
  sem aplicar), **Reiniciar** (limpa a janela) e **Cancelar**.
- **Desfazer** fica sempre à vista, acima das abas. O banco original nunca é alterado: desfazer refaz tudo desde o
  início sem o passo removido.
- **Abas de baixo**: *Visualização de dados* (as linhas), *Visualização de variáveis* (escolha uma variável e use
  os botões para editar) e *Syntax* (o .sps de tudo o que você fez, com os passos).
- **Barra de status** (rodapé): arquivo, casos × variáveis e quantos passos foram aplicados.
"""

# --------------------------------------------------------------------------------------- Calcular
TECLADO: List[List[Tuple[str, str]]] = [
    [("+", "+"), ("−", "-"), ("×", "*"), ("÷", "/"), ("**", "**"), ("(", "("), (")", ")"), (",", ",")],
    [("=", "="), ("<>", "<>"), ("<", "<"), (">", ">"), ("<=", "<="), (">=", ">="), ("E", "AND"), ("OU", "OR")],
    [("NÃO", "NOT"), ("0", "0"), ("1", "1"), ("2", "2"), ("3", "3"), ("4", "4"), ("5", "5"), ("9", "9")],
]

FUNCOES: Dict[str, List[str]] = {
    "Estatísticas": ["MEAN(Q1 TO Q5)", "MEAN.2(Q1 TO Q5)", "SUM(Q1, Q2)", "MIN(Q1, Q2)", "MAX(Q1, Q2)",
                     "SD(Q1 TO Q5)", "MEDIAN(Q1 TO Q5)", "NVALID(Q1 TO Q5)"],
    "Aritméticas": ["ABS(x)", "SQRT(x)", "EXP(x)", "LN(x)", "LG10(x)", "RND(x)", "TRUNC(x)", "MOD(x, y)"],
    "Testes e lógicas": ["ANY(x, 1, 2)", "RANGE(x, 1, 5)", "MISSING(x)", "SYSMIS(x)", "VALUE(x)"],
    "Texto": ["CONCAT(a, '-', b)", "UPCASE(x)", "LOWER(x)", "LTRIM(x)", "RTRIM(x)", "LENGTH(x)",
              "SUBSTR(x, 1, 3)", "NUMBER(x, F8)"],
    "Outras": ["$CASENUM", "$SYSMIS"],
}


def inserir_trecho(atual: str, trecho: str) -> str:
    """Acrescenta `trecho` ao fim de `atual` com o espaçamento certo (usado pelos botões de inserir)."""
    atual = atual or ""
    if not atual.strip():
        return trecho
    if trecho in {",", ")"} or atual.endswith(("(", " ")):
        return atual + trecho
    return f"{atual} {trecho}"


# --------------------------------------------------------------------------------------- Recode
TIPOS_ANTIGO = [
    "Valor",
    "Intervalo (de … até …)",
    "Do menor valor até …",
    "De … até o maior valor",
    "Ausente do sistema",
    "Ausente (sistema ou usuário)",
    "Todos os outros valores",
]
_T_VALOR, _T_INTERVALO, _T_ATE, _T_A_PARTIR, _T_SYSMIS, _T_MISSING, _T_ELSE = TIPOS_ANTIGO


def _s(x: Any) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and math.isnan(x):
        return ""
    try:
        if pd.isna(x):
            return ""
    except (TypeError, ValueError):
        pass
    return str(x).strip()


def regras_de_tabela(linhas: Iterable[Mapping[str, Any]]) -> str:
    """Converte a grade de 'valores antigos e novos' no texto de regras do RECODE.

    Colunas esperadas: Tipo, De, Até, Novo valor. Linhas totalmente vazias do tipo 'Valor' (e as de intervalo
    sem nada preenchido) são ignoradas. 'Todos os outros valores' vai sempre por último, como exige o SPSS.
    """
    regras: List[str] = []
    final: List[str] = []
    for n, lin in enumerate(linhas, 1):
        tipo = _s(lin.get("Tipo")) or _T_VALOR
        de, ate, novo = _s(lin.get("De")), _s(lin.get("Até")), _s(lin.get("Novo valor"))
        if tipo not in TIPOS_ANTIGO:
            raise ErroTratamento(f"Linha {n}: tipo de valor antigo desconhecido ('{tipo}').")
        usa_de = tipo in (_T_VALOR, _T_INTERVALO, _T_A_PARTIR)
        usa_ate = tipo in (_T_INTERVALO, _T_ATE)
        if tipo in (_T_VALOR, _T_INTERVALO, _T_ATE, _T_A_PARTIR) and not (de or ate or novo):
            continue                                           # linha em branco da grade
        if usa_de and not de:
            raise ErroTratamento(f"Linha {n}: informe o valor antigo (coluna “De”).")
        if usa_ate and not ate:
            raise ErroTratamento(f"Linha {n}: informe o limite (coluna “Até”).")
        if not novo:
            raise ErroTratamento(f"Linha {n}: informe o novo valor (um número, SYSMIS ou COPY).")
        if tipo == _T_VALOR:
            antigo = de
        elif tipo == _T_INTERVALO:
            antigo = f"{de} THRU {ate}"
        elif tipo == _T_ATE:
            antigo = f"LO THRU {ate}"
        elif tipo == _T_A_PARTIR:
            antigo = f"{de} THRU HI"
        elif tipo == _T_SYSMIS:
            antigo = "SYSMIS"
        elif tipo == _T_MISSING:
            antigo = "MISSING"
        else:
            antigo = "ELSE"
        (final if tipo == _T_ELSE else regras).append(f"{antigo} = {novo}")
    todas = regras + final
    if not todas:
        raise ErroTratamento("Informe ao menos uma regra de recodificação.")
    return "\n".join(todas)


# --------------------------------------------------------------------------------------- Missing
MODOS_MISSING = {
    "nenhum": "Sem valores ausentes",
    "discretos": "Valores ausentes discretos (até 3)",
    "faixa": "Faixa mais um valor discreto opcional",
}


def _aspas(valor: str, texto: bool) -> str:
    valor = valor.strip()
    if not texto or not valor or valor[0] in "'\"":
        return valor
    return "'" + valor.replace("'", "''") + "'"


def especificacao_missing(modo: str, discretos: Sequence[str] = (), faixa_de: str = "", faixa_ate: str = "",
                          faixa_extra: str = "", texto: bool = False) -> str:
    """Monta o texto de MISSING VALUES (ex.: '98, 99' ou 'LO THRU 0, 99'). Vazio = remove os missings."""
    if modo == "nenhum":
        return ""
    if modo == "discretos":
        vals = [_aspas(_s(v), texto) for v in discretos if _s(v)]
        if not vals:
            raise ErroTratamento("Informe ao menos um valor ausente.")
        if len(vals) > 3:
            raise ErroTratamento("O SPSS aceita no máximo 3 valores ausentes discretos.")
        return ", ".join(vals)
    if modo == "faixa":
        de, ate, extra = _s(faixa_de), _s(faixa_ate), _s(faixa_extra)
        if not de and not ate:
            raise ErroTratamento("Informe ao menos um limite da faixa (mínimo ou máximo).")
        partes = [f"{_aspas(de, texto) or 'LO'} THRU {_aspas(ate, texto) or 'HI'}"]
        if extra:
            partes.append(_aspas(extra, texto))
        return ", ".join(partes)
    raise ErroTratamento(f"Modo de valores ausentes desconhecido: {modo}.")


# --------------------------------------------------------------------------------------- grades (HTML)
MEDIDA_NOME = {"": "", "nominal": "Nominal", "ordinal": "Ordinal", "scale": "Escala"}
TIPO_NOME = {"Numérica": "Numérico", "Texto": "String"}


def _e(x: Any) -> str:
    return html.escape("" if x is None else str(x), quote=True)


def _formatador_numerico(serie: pd.Series):
    """Inteiro quando todos os valores são inteiros; senão 2 casas. Ausente do sistema = '.', como no SPSS."""
    v = pd.to_numeric(serie, errors="coerce").dropna()
    inteiro = len(v) == 0 or bool(np.all(np.mod(v.to_numpy(dtype=float), 1) == 0))

    def f(x: Any) -> str:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return "."
        try:
            return f"{x:.0f}" if inteiro else f"{x:.2f}"
        except (TypeError, ValueError):
            return str(x)
    return f


def tabela_html_dados(banco: Any, linhas: int = 100, com_rotulos: bool = False) -> str:
    """Data View: cabeçalho cinza com os nomes (rótulo na dica), coluna de nº do caso e células com grade fina."""
    df = banco.df.head(int(linhas))
    cols = [str(c) for c in df.columns]
    texto = [not (pd.api.types.is_numeric_dtype(df[c]) or pd.api.types.is_datetime64_any_dtype(df[c])) for c in df.columns]
    fmts = [None if t else _formatador_numerico(banco.df[c]) for c, t in zip(df.columns, texto)]
    rotulos = [banco.rotulos_valor.get(c, {}) if (com_rotulos and not t) else {} for c, t in zip(cols, texto)]
    cab = "".join(f'<th title="{_e(banco.rotulos_var.get(c, ""))}">{_e(c)}</th>' for c in cols)
    corpo: List[str] = []
    for n, linha in enumerate(df.itertuples(index=False, name=None), 1):
        cel = [f'<th class="td-rn">{n}</th>']
        for i, x in enumerate(linha):
            if texto[i]:
                cel.append(f'<td>{_e("" if x is None or (isinstance(x, float) and math.isnan(x)) else x)}</td>')
            else:
                rot = rotulos[i].get(x) if rotulos[i] and not (isinstance(x, float) and math.isnan(x)) else None
                cel.append(f'<td class="td-n">{_e(rot) if rot is not None else fmts[i](x)}</td>')
        corpo.append("<tr>" + "".join(cel) + "</tr>")
    return (f'<div class="td-grid"><table><thead><tr><th class="td-rn"></th>{cab}</tr></thead>'
            f'<tbody>{"".join(corpo)}</tbody></table></div>')


def tabela_html_variaveis(banco: Any) -> str:
    """Variable View: uma linha por variável, com as colunas que o arquivo realmente informa."""
    tab = tabela_variaveis(banco)
    nomes = ["Nome", "Tipo", "Rótulo", "Valores", "Ausentes", "Medida", "Casos válidos"]
    cab = "".join(f"<th>{n}</th>" for n in nomes)
    corpo: List[str] = []
    for n, r in enumerate(tab.itertuples(index=False), 1):
        valores = [r[0], TIPO_NOME.get(r[1], r[1]), r[2], r[3], r[4], MEDIDA_NOME.get(r[5], r[5]), r[6]]
        cel = "".join(f'<td class="td-n">{_e(v)}</td>' if k == 6 else f"<td>{_e(v)}</td>" for k, v in enumerate(valores))
        corpo.append(f'<tr><th class="td-rn">{n}</th>{cel}</tr>')
    return (f'<div class="td-grid"><table><thead><tr><th class="td-rn"></th>{cab}</tr></thead>'
            f'<tbody>{"".join(corpo)}</tbody></table></div>')
