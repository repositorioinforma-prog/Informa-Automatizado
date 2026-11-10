"""
Tratamento do Banco — etapa 1 (lógica pura, sem Streamlit).

Junta, em um fluxo só:
  1. Revisão da Syntax (reaproveita o Revisor de Syntax);
  2. Execução da Syntax sobre o banco original (core.syntax_executor_math);
  3. Conferência da recodificação (variáveis criadas, códigos sem rótulo,
     rótulos sem casos, conjuntos múltiplos);
  4. Comparação, variável a variável, com o banco que o SPSS gerou
     (quando o usuário envia um) — valores, rótulos de valor e de variável;
  5. Exportação do banco tratado em .sav.

O banco só é considerado "aprovado" quando não há erro crítico nem possível
erro e, havendo banco do SPSS para comparar, todas as variáveis batem.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Any

import numpy as np
import pandas as pd

from core.analitica_dados import BancoCarregado
from core.revisor_syntax_math import SEVERITY_ORDER, Finding, review_syntax
from core.syntax_executor_math import ResultadoExecucao, executar_syntax, variaveis_geradas

LIMITE_VALORES_CATEGORICA = 40       # acima disso, a variável é tratada como contínua
SEM_CONFERENCIA_DE_ROTULO = {"TOTAL"}


# ---------------------------------------------------------------------------
# Estruturas
# ---------------------------------------------------------------------------

@dataclass
class ComparacaoVariavel:
    variavel: str
    existe_no_spss: bool
    valores_iguais: bool | None = None
    casos_diferentes: int = 0
    rotulos_de_valor_iguais: bool | None = None
    rotulo_de_variavel_igual: bool | None = None
    exemplo: str = ""
    detalhe: str = ""

    @property
    def igual(self) -> bool:
        if not self.existe_no_spss:
            return False
        return bool(self.valores_iguais) and bool(self.rotulos_de_valor_iguais) and bool(self.rotulo_de_variavel_igual)


@dataclass
class ResultadoTratamento:
    execucao: ResultadoExecucao
    revisao: dict[str, Any]
    conferencias: list[Finding]
    comparacao: list[ComparacaoVariavel] | None
    achados: list[Finding]                      # tudo junto, ordenado por gravidade
    erro_comparacao: str | None = None
    _revisao_findings: list[Finding] = field(default_factory=list)

    @property
    def achados_do_tratamento(self) -> list[Finding]:
        """Execução + conferências + comparação (o que este módulo decide)."""
        return [f for f in self.achados if f not in self.achados_da_revisao]

    @property
    def achados_da_revisao(self) -> list[Finding]:
        return list(self._revisao_findings)

    @property
    def contagem(self) -> dict[str, int]:
        out = {k: 0 for k in SEVERITY_ORDER}
        for f in self.achados_do_tratamento:
            out[f.severity] = out.get(f.severity, 0) + 1
        return out

    @property
    def contagem_revisao(self) -> dict[str, int]:
        out = {k: 0 for k in SEVERITY_ORDER}
        for f in self.achados_da_revisao:
            out[f.severity] = out.get(f.severity, 0) + 1
        return out

    @property
    def pendencias_da_revisao(self) -> int:
        c = self.contagem_revisao
        return c.get("critico", 0) + c.get("possivel_erro", 0)

    @property
    def divergentes(self) -> list[ComparacaoVariavel]:
        return [c for c in (self.comparacao or []) if not c.igual]

    @property
    def aprovado(self) -> bool:
        c = self.contagem
        if c.get("critico", 0) or c.get("possivel_erro", 0):
            return False
        if self.comparacao is not None and self.divergentes:
            return False
        if self.erro_comparacao:
            return False
        return True


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _tocadas(exe: ResultadoExecucao) -> list[str]:
    """Variáveis criadas + alteradas pela Syntax (sem repetir, sem ligar para maiúsculas)."""
    vistos = {n.upper() for n in exe.criadas}
    return list(exe.criadas) + [m for m in exe.modificadas if m.upper() not in vistos]


def _coluna(df: pd.DataFrame, nome: str):
    n = nome.upper()
    for c in df.columns:
        if str(c).upper() == n:
            return c
    return None


def _num(v: Any) -> float | None:
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _mapa_rotulos(value_labels: dict[Any, str] | None) -> dict[Any, str]:
    out: dict[Any, str] = {}
    for k, v in (value_labels or {}).items():
        n = _num(k)
        out[n if n is not None else str(k)] = str(v).strip()
    return out


def _valores_presentes(serie: pd.Series) -> list[Any]:
    if pd.api.types.is_numeric_dtype(serie):
        return sorted(float(v) for v in serie.dropna().unique())
    return sorted(str(v).strip() for v in serie.dropna().unique() if str(v).strip())


def _fmt(v: Any) -> str:
    n = _num(v)
    if n is not None:
        return str(int(n)) if n.is_integer() else str(n)
    return str(v)


# ---------------------------------------------------------------------------
# Conferência da recodificação
# ---------------------------------------------------------------------------

def conferir_recodificacao(exe: ResultadoExecucao) -> list[Finding]:
    """Confere as variáveis que a Syntax criou/alterou, olhando o resultado."""
    achados: list[Finding] = []
    linha_de = {}
    for r in exe.registros:
        for v in r.variaveis:
            linha_de.setdefault(v.upper(), r.linha)

    sem_rotulo_explicito = {
        r.variaveis[0].upper() for r in exe.registros
        if r.tipo == "VAR LABEL" and r.variaveis
    }

    sem_rotulo_de_variavel: list[str] = []
    for nome in _tocadas(exe):
        col = _coluna(exe.df, nome)
        if col is None:
            continue
        serie = exe.df[col]
        linha = linha_de.get(nome.upper())
        n_validos = int(serie.notna().sum()) if pd.api.types.is_numeric_dtype(serie) \
            else int((serie.astype(str).str.strip() != "").sum())

        if n_validos == 0:
            selecao = _eh_recode_de_selecao(exe, nome)
            achados.append(Finding(
                "revisar" if selecao else "possivel_erro", "TRAT_VARIAVEL_VAZIA",
                f"{nome} ficou sem nenhum caso",
                f"A variável {nome} foi criada, mas nenhum caso recebeu valor. "
                + ("Como o RECODE só separa uma categoria, pode ser que ninguém tenha dado essa resposta."
                   if selecao else
                   "Provavelmente a regra não casa com os códigos da variável de origem."),
                linha, nome, None, "Confira os códigos de origem no RECODE.",
            ))
            continue

        if nome.upper() in SEM_CONFERENCIA_DE_ROTULO:
            continue

        rotulos = _mapa_rotulos(exe.value_labels.get(nome.upper()))
        rotulo_var = exe.var_labels.get(nome.upper())
        distintos = _valores_presentes(serie)

        if rotulo_var is None and nome.upper() not in sem_rotulo_explicito:
            sem_rotulo_de_variavel.append(nome)

        if not rotulos:
            if len(distintos) <= LIMITE_VALORES_CATEGORICA:
                achados.append(Finding(
                    "revisar", "TRAT_SEM_ROTULO_DE_VALOR", f"{nome} não tem rótulos de valor",
                    f"{nome} tem {len(distintos)} código(s) distinto(s) e nenhum VALUE LABELS.",
                    linha, nome, None, "Inclua o VALUE LABELS dessa variável.",
                ))
            continue

        sem_rotulo = [v for v in distintos if v not in rotulos]
        if sem_rotulo:
            achados.append(Finding(
                "revisar", "TRAT_CODIGO_SEM_ROTULO", f"{nome}: código(s) sem rótulo",
                f"Há casos em {nome} com código(s) {', '.join(_fmt(v) for v in sem_rotulo[:10])} "
                "que não têm rótulo no VALUE LABELS.",
                linha, nome, None, "Inclua o rótulo desses códigos ou revise o RECODE.",
            ))
        sem_caso = [k for k in rotulos if k not in set(distintos)]
        if sem_caso:
            achados.append(Finding(
                "informativo", "TRAT_ROTULO_SEM_CASO", f"{nome}: rótulo(s) sem nenhum caso",
                f"Os códigos {', '.join(_fmt(v) for v in sem_caso[:10])} têm rótulo em {nome}, "
                "mas nenhum caso recebeu esse valor.",
                linha, nome, None,
                "Normal se ninguém escolheu a opção; senão, revise o RECODE.",
            ))

    if sem_rotulo_de_variavel:
        achados.append(Finding(
            "revisar", "TRAT_SEM_ROTULO_DE_VARIAVEL",
            f"{len(sem_rotulo_de_variavel)} variável(is) criada(s) sem rótulo de variável",
            "Estas variáveis foram criadas e não receberam VARIABLE LABELS: "
            + ", ".join(sem_rotulo_de_variavel) + ".",
            None, sem_rotulo_de_variavel[0], None, "Inclua o rótulo (ou use '' se for intencional).",
        ))

    achados.extend(_conferir_codigos_entre_variaveis(exe))
    achados.extend(_conferir_rotulos_parecidos(exe))

    # conjuntos de respostas múltiplas
    for conj, membros in exe.mrsets.items():
        faltam = [m for m in membros if _coluna(exe.df, m) is None]
        if faltam:
            achados.append(Finding(
                "critico", "TRAT_MRSET_MEMBRO_AUSENTE", f"{conj}: variável(is) inexistente(s)",
                f"O conjunto {conj} lista {', '.join(faltam)}, que não existe(m) no banco tratado.",
                None, conj, None, "Confira os nomes em MRSETS.",
            ))

    return achados


def _eh_recode_de_selecao(exe: ResultadoExecucao, destino: str) -> bool:
    for rc in exe.recodes:
        for dest, regras in zip(rc.destinos, rc.regras):
            if dest.upper() == destino.upper():
                return len(regras) == 1
    return False


CODIGOS_NAO_RESPOSTA = (97.0, 98.0, 99.0)   # nulo/branco, não irá votar, não sabe — padrão do Instituto


def _canon(txt: str) -> str:
    import unicodedata
    t = unicodedata.normalize("NFD", txt or "")
    t = "".join(ch for ch in t if unicodedata.category(ch) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def _conferir_codigos_entre_variaveis(exe: ResultadoExecucao) -> list[Finding]:
    """Um mesmo rótulo especial (ex.: 'Branco/Nulo') deve ter o MESMO código em
    todas as variáveis criadas. Se uma variável usa outro código, é quase
    sempre uma troca de 97/98 — o tipo de erro que passa despercebido."""
    achados: list[Finding] = []
    por_rotulo: dict[str, dict[str, list[tuple[str, float]]]] = {}
    for nome in _tocadas(exe):
        for cod, rot in _mapa_rotulos(exe.value_labels.get(nome.upper())).items():
            if cod not in CODIGOS_NAO_RESPOSTA:
                continue
            c = _canon(rot)
            if c:
                por_rotulo.setdefault(c, {}).setdefault(_fmt(cod), []).append((nome, cod))
    for c, codigos in por_rotulo.items():
        if len(codigos) < 2:
            continue
        maioria, membros_maioria = max(codigos.items(), key=lambda kv: len(kv[1]))
        if len(membros_maioria) < 3:
            continue
        for cod, membros in codigos.items():
            if cod == maioria:
                continue
            for nome, _ in membros:
                rot = _mapa_rotulos(exe.value_labels.get(nome.upper())).get(float(cod), c)
                achados.append(Finding(
                    "possivel_erro", "TRAT_CODIGO_INCONSISTENTE",
                    f"{nome}: '{rot}' usa o código {cod} (nas outras, {maioria})",
                    f"Em {len(membros_maioria)} variáveis o rótulo '{rot}' tem o código {maioria}, "
                    f"mas em {nome} ele tem o código {cod}. Se os códigos 97/98 estiverem trocados, "
                    "os percentuais saem nas linhas erradas.",
                    None, nome, None,
                    "Confira o questionário/banco: o código está trocado no VALUE LABELS ou na codificação do campo.",
                ))
    return achados


LIMIAR_ROTULO_PARECIDO = 0.94


def _conferir_rotulos_parecidos(exe: ResultadoExecucao) -> list[Finding]:
    """Acha o mesmo nome escrito de formas diferentes em variáveis diferentes
    (ex.: 'Alex da Confecção' x 'Alexda Confecção', 'UNIÃO BRASIL' x
    'UNIÃOBRASIL'). Rótulos parecidos dentro da MESMA variável são opções
    distintas e não entram."""
    import difflib

    donos: dict[str, set[str]] = {}
    for nome in _tocadas(exe):
        for cod, rot in _mapa_rotulos(exe.value_labels.get(nome.upper())).items():
            if len(rot) >= 6:
                donos.setdefault(rot, set()).add(nome.upper())

    chave = {rot: re.sub(r"\s+", "", _canon(rot)) for rot in donos}
    blocos: dict[str, list[str]] = {}
    for rot, k in chave.items():
        blocos.setdefault(k[:3], []).append(rot)

    achados: list[Finding] = []
    vistos: set[frozenset] = set()
    for grupo in blocos.values():
        for i, a in enumerate(grupo):
            for b in grupo[i + 1:]:
                if donos[a] & donos[b]:
                    continue
                par = frozenset((a, b))
                if par in vistos:
                    continue
                igual = chave[a] == chave[b]
                if not igual and difflib.SequenceMatcher(None, chave[a], chave[b]).ratio() < LIMIAR_ROTULO_PARECIDO:
                    continue
                vistos.add(par)
                va, vb = sorted(donos[a]), sorted(donos[b])
                so_caixa = a.strip().lower() == b.strip().lower()
                achados.append(Finding(
                    "informativo" if so_caixa else "revisar", "TRAT_ROTULO_PARECIDO",
                    f"Nome escrito de dois jeitos: '{a}' e '{b}'",
                    f"'{a}' aparece em {', '.join(va[:6])} e '{b}' aparece em {', '.join(vb[:6])}. "
                    "Parece o mesmo nome com grafia diferente.",
                    None, (va + vb)[0], None,
                    "Padronize a grafia no VALUE LABELS para a tabela não separar a mesma opção em duas.",
                ))
    return achados


# ---------------------------------------------------------------------------
# Comparação com o banco processado pelo SPSS
# ---------------------------------------------------------------------------

def _iguais(a: pd.Series, b: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(a) and pd.api.types.is_numeric_dtype(b):
        af, bf = a.astype(float), b.astype(float)
        return (af == bf) | (af.isna() & bf.isna())
    sa = a.astype(object).where(a.notna(), "").astype(str).str.rstrip()
    sb = b.astype(object).where(b.notna(), "").astype(str).str.rstrip()
    return sa == sb


def comparar_com_spss(exe: ResultadoExecucao, spss: BancoCarregado) -> list[ComparacaoVariavel]:
    """Compara cada variável criada/alterada pela Syntax com a do banco que o
    SPSS gerou (valores, rótulos de valor e rótulo de variável)."""
    out: list[ComparacaoVariavel] = []
    spss_labels = {str(k).upper(): v for k, v in spss.var_labels.items()}
    spss_vl = {str(k).upper(): v for k, v in spss.value_labels.items()}

    mesmo_tamanho = len(exe.df) == len(spss.df)
    for nome in _tocadas(exe):
        col_m = _coluna(exe.df, nome)
        col_s = _coluna(spss.df, nome)
        if col_s is None:
            out.append(ComparacaoVariavel(nome, False, detalhe="não existe no banco do SPSS"))
            continue
        cmp = ComparacaoVariavel(nome, True)
        if not mesmo_tamanho:
            cmp.valores_iguais = False
            cmp.detalhe = f"bancos com números de casos diferentes ({len(exe.df)} × {len(spss.df)})"
        else:
            ok = _iguais(exe.df[col_m].reset_index(drop=True), spss.df[col_s].reset_index(drop=True))
            cmp.casos_diferentes = int((~ok).sum())
            cmp.valores_iguais = cmp.casos_diferentes == 0
            if not cmp.valores_iguais:
                idx = np.flatnonzero(~ok.to_numpy())[:3]
                partes = []
                for i in idx:
                    partes.append(f"caso {int(i) + 1}: meu={_fmt(exe.df[col_m].iloc[i])} · SPSS={_fmt(spss.df[col_s].iloc[i])}")
                cmp.exemplo = "; ".join(partes)
        mine = _mapa_rotulos(exe.value_labels.get(nome.upper()))
        theirs = _mapa_rotulos(spss_vl.get(nome.upper()))
        cmp.rotulos_de_valor_iguais = mine == theirs
        if not cmp.rotulos_de_valor_iguais:
            dif = sorted({k for k in set(mine) | set(theirs) if mine.get(k) != theirs.get(k)}, key=str)
            cmp.detalhe = (cmp.detalhe + " " if cmp.detalhe else "") + \
                "rótulos de valor diferentes nos códigos: " + ", ".join(_fmt(k) for k in dif[:8])
        cmp.rotulo_de_variavel_igual = (exe.var_labels.get(nome.upper(), "") or "").strip() == \
                                       (spss_labels.get(nome.upper(), "") or "").strip()
        if not cmp.rotulo_de_variavel_igual:
            cmp.detalhe = (cmp.detalhe + " " if cmp.detalhe else "") + "rótulo da variável difere"
        out.append(cmp)
    return out


def achados_da_comparacao(comparacao: list[ComparacaoVariavel]) -> list[Finding]:
    achados: list[Finding] = []
    for c in comparacao:
        if c.igual:
            continue
        if not c.existe_no_spss:
            achados.append(Finding(
                "revisar", "TRAT_FORA_DO_BANCO_SPSS", f"{c.variavel} não existe no banco do SPSS",
                f"{c.variavel} foi criada pela Syntax aqui, mas não está no banco que o SPSS gerou.",
                None, c.variavel, None, "Confira se o banco do SPSS é o mesmo desta Syntax.",
            ))
            continue
        if c.valores_iguais is False:
            achados.append(Finding(
                "critico", "TRAT_VALORES_DIFERENTES",
                f"{c.variavel}: valores diferentes do SPSS ({c.casos_diferentes} caso(s))",
                c.detalhe or f"{c.casos_diferentes} caso(s) com valor diferente. {c.exemplo}",
                None, c.variavel, c.exemplo or None,
                "A recodificação feita aqui não bate com a do SPSS: não use este banco antes de entender a diferença.",
            ))
        if c.rotulos_de_valor_iguais is False or c.rotulo_de_variavel_igual is False:
            achados.append(Finding(
                "possivel_erro", "TRAT_ROTULOS_DIFERENTES", f"{c.variavel}: rótulos diferentes do SPSS",
                c.detalhe, None, c.variavel, None, "Compare os rótulos da Syntax com os do banco do SPSS.",
            ))
    return achados


# ---------------------------------------------------------------------------
# Fluxo completo
# ---------------------------------------------------------------------------

def _metadados_simples(banco: BancoCarregado) -> SimpleNamespace:
    return SimpleNamespace(
        column_names_to_labels=dict(banco.var_labels),
        variable_value_labels={k: dict(v) for k, v in banco.value_labels.items()},
        column_names=list(banco.df.columns),
    )


def _na_ordem_original(df: pd.DataFrame, ordem: list[str]) -> pd.DataFrame:
    """Devolve as colunas na ordem do banco enviado; as novas ficam no fim."""
    por_nome = {str(c).upper(): c for c in df.columns}
    cols: list[Any] = []
    for n in ordem:
        c = por_nome.pop(n.upper(), None)
        if c is not None:
            cols.append(c)
    cols.extend(por_nome.values())
    return df[cols]


def banco_sem_geradas(sps_bytes: bytes, banco: BancoCarregado) -> tuple[BancoCarregado, list[str]]:
    """Tira do banco as variáveis que a Syntax cria (se já estiverem lá),
    para a Syntax rodar do zero. Devolve (banco limpo, nomes removidos)."""
    geradas = {g.upper() for g in variaveis_geradas(sps_bytes)}
    remover = [c for c in banco.df.columns if str(c).upper() in geradas]
    if not remover:
        return banco, []
    nomes = {str(c).upper() for c in remover}
    limpo = BancoCarregado(
        df=banco.df.drop(columns=remover),
        nome_arquivo=banco.nome_arquivo,
        var_labels={k: v for k, v in banco.var_labels.items() if str(k).upper() not in nomes},
        value_labels={k: v for k, v in banco.value_labels.items() if str(k).upper() not in nomes},
        missing_ranges={k: v for k, v in banco.missing_ranges.items() if str(k).upper() not in nomes},
    )
    return limpo, [str(c) for c in remover]


def tratar_banco(
    sps_bytes: bytes,
    banco: BancoCarregado,
    banco_spss: BancoCarregado | None = None,
    partir_do_zero: bool = True,
) -> ResultadoTratamento:
    removidas: list[str] = []
    ordem_original = [str(c) for c in banco.df.columns]
    if partir_do_zero:
        banco, removidas = banco_sem_geradas(sps_bytes, banco)
    try:
        revisao = review_syntax(sps_bytes, banco.df, _metadados_simples(banco),
                                banco_spss.df if banco_spss is not None else None)
    except Exception as e:  # noqa: BLE001 — a revisão não pode derrubar o tratamento
        revisao = {"counts": {k: 0 for k in SEVERITY_ORDER}, "findings": [], "erro": str(e)}

    exe = executar_syntax(sps_bytes, banco.df, banco.var_labels, banco.value_labels, banco.missing_ranges)
    exe.df = _na_ordem_original(exe.df, ordem_original)
    conferencias = conferir_recodificacao(exe)

    comparacao = None
    erro_cmp = None
    extras: list[Finding] = []
    if banco_spss is not None:
        try:
            comparacao = comparar_com_spss(exe, banco_spss)
            extras = achados_da_comparacao(comparacao)
        except Exception as e:  # noqa: BLE001
            erro_cmp = f"Não foi possível comparar com o banco do SPSS: {e}"

    revisao_findings = [Finding(**{k: f.get(k) for k in
                                   ("severity", "code", "title", "message", "line", "variable", "snippet", "suggestion")})
                        for f in revisao.get("findings", [])]
    extras_info: list[Finding] = []
    if removidas:
        extras_info.append(Finding(
            "informativo", "TRAT_GERADAS_REMOVIDAS",
            f"{len(removidas)} variável(is) já existiam no banco e foram refeitas",
            "O banco enviado já tinha variáveis que a Syntax cria (" + ", ".join(removidas[:12])
            + (" …" if len(removidas) > 12 else "") + "). Elas foram retiradas e recalculadas pela Syntax, "
            "para a conferência valer de verdade.",
            None, None, None, None,
        ))
    longos = [k for k in rotulos_longos(exe.var_labels) if k in {str(c).upper() for c in exe.df.columns}]
    if longos:
        extras_info.append(Finding(
            "informativo", "TRAT_ROTULO_LONGO",
            f"{len(longos)} rótulo(s) de variável passam de {LIMITE_ROTULO_BYTES} bytes",
            "O .sav gravado pelo app encurta estes rótulos de variável (limite da gravação): "
            + ", ".join(longos[:12]) + (" …" if len(longos) > 12 else "")
            + ". São textos longos do questionário; não afetam as tabelas.",
            None, longos[0], None, None,
        ))
    todos = revisao_findings + exe.findings + conferencias + extras + extras_info
    todos.sort(key=lambda f: (SEVERITY_ORDER.get(f.severity, 99), f.line or 10**9, f.code))
    return ResultadoTratamento(exe, revisao, conferencias, comparacao, todos, erro_cmp, revisao_findings)


# ---------------------------------------------------------------------------
# Exportação
# ---------------------------------------------------------------------------

def nome_arquivo_tratado(nome_original: str) -> str:
    base = re.sub(r"\.sav$", "", nome_original or "Banco", flags=re.I)
    base = re.sub(r"^Banco_Original_", "Banco_Tratado_", base, flags=re.I)
    if not base.lower().startswith("banco_tratado"):
        base = f"{base}_tratado"
    return f"{base}.sav"


LIMITE_ROTULO_BYTES = 255        # limite do formato .sav na gravação


def _cortar_bytes(txt: str, limite: int = LIMITE_ROTULO_BYTES) -> str:
    """Corta em `limite` bytes sem deixar um caractere acentuado pela metade."""
    b = txt.encode("utf-8")
    if len(b) <= limite:
        return txt
    return b[:limite].decode("utf-8", errors="ignore").rstrip()


def rotulos_longos(var_labels: dict[str, str]) -> list[str]:
    return [k for k, v in var_labels.items() if v and len(v.encode("utf-8")) > LIMITE_ROTULO_BYTES]


def exportar_sav(exe: ResultadoExecucao) -> bytes:
    from controle.exporters.sav import write_sav_bytes

    nomes = {str(c).upper(): c for c in exe.df.columns}
    column_labels = {nomes[k]: _cortar_bytes(v) for k, v in exe.var_labels.items() if k in nomes and v is not None}
    value_labels: dict[str, dict] = {}
    for k, mp in exe.value_labels.items():
        if k not in nomes or not mp:
            continue
        limpo = {}
        for cod, rot in mp.items():
            n = _num(cod)
            limpo[n if n is not None else str(cod)] = str(rot)
        value_labels[nomes[k]] = limpo
    return write_sav_bytes(exe.df, column_labels, value_labels, file_label="Banco tratado - Informa")
