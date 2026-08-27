"""Serialização e validação de projetos do Gerador de Amostra.

Os dados territoriais continuam no SQLite. O JSON guarda somente as escolhas
feitas pelo usuário, permitindo baixar, arquivar e reabrir um desenho amostral
sem duplicar o banco demográfico.
"""

from __future__ import annotations

import json
import re
from copy import deepcopy
from datetime import datetime, timezone
from typing import Any, Mapping

VERSAO_SCHEMA_PROJETO = "1.2"
VERSOES_SCHEMA_COMPATIVEIS = {"1.0", "1.1", VERSAO_SCHEMA_PROJETO}
FONTES_DADOS_VALIDAS = {"banco_sqlite", "bases_legadas"}
MODOS_VALIDOS = {"classico", "camadas"}
BASES_POPULACIONAIS_VALIDAS = {"16_mais", "total"}
METODOS_DISTRIBUICAO_VALIDOS = {"proporcional", "manual"}
EXIBICOES_COTAS_VALIDAS = {"totais", "percentuais", "ambos"}
NIVEIS_VALIDOS = {
    "uf",
    "regiao_intermediaria",
    "regiao_imediata",
    "municipio",
    "distrito",
}
CAMPOS_RECORTES_TERRITORIAIS = (
    "regioes_intermediarias",
    "regioes_imediatas",
    "municipios",
    "distritos",
)


class ProjetoAmostraErro(ValueError):
    """Configuração JSON inválida ou incompatível."""


def _inteiro_positivo(valor: Any, campo: str) -> int:
    try:
        convertido = int(valor)
    except (TypeError, ValueError) as exc:
        raise ProjetoAmostraErro(f"{campo} precisa ser um número inteiro.") from exc
    if convertido <= 0:
        raise ProjetoAmostraErro(f"{campo} precisa ser maior que zero.")
    return convertido


def _percentual(valor: Any, campo: str) -> int:
    try:
        convertido = int(valor)
    except (TypeError, ValueError) as exc:
        raise ProjetoAmostraErro(f"{campo} precisa ser um número inteiro.") from exc
    if convertido < 0 or convertido > 100:
        raise ProjetoAmostraErro(f"{campo} precisa estar entre 0 e 100.")
    return convertido


def _normalizar_recorte_territorial(valor: Any, campo: str) -> list[dict[str, str]]:
    """Valida uma lista portátil de unidades territoriais por código e nome."""
    itens = valor or []
    if not isinstance(itens, list):
        raise ProjetoAmostraErro(f"{campo} precisa ser uma lista.")

    normalizados: list[dict[str, str]] = []
    codigos_vistos: set[str] = set()
    for item in itens:
        if not isinstance(item, Mapping):
            raise ProjetoAmostraErro(f"Cada item de {campo} precisa ser um objeto.")
        codigo = str(item.get("codigo") or "").strip()
        nome = str(item.get("nome") or "").strip()
        if not codigo or not nome:
            raise ProjetoAmostraErro(
                f"Cada item de {campo} precisa possuir código e nome."
            )
        if codigo in codigos_vistos:
            continue
        codigos_vistos.add(codigo)
        normalizados.append({"codigo": codigo, "nome": nome})
    return normalizados


def _normalizar_alocacoes_manuais(valor: Any) -> list[dict[str, Any]]:
    itens = valor or []
    if not isinstance(itens, list):
        raise ProjetoAmostraErro("alocacoes_manuais precisa ser uma lista.")
    normalizados = []
    vistos = set()
    for item in itens:
        if not isinstance(item, Mapping):
            raise ProjetoAmostraErro("Cada alocação manual precisa ser um objeto.")
        codigo = str(item.get("codigo") or "").strip()
        if not codigo:
            raise ProjetoAmostraErro("Cada alocação manual precisa possuir código.")
        if codigo in vistos:
            raise ProjetoAmostraErro(f"Código repetido na alocação manual: {codigo}.")
        vistos.add(codigo)
        try:
            amostra_float = float(item.get("amostra"))
        except (TypeError, ValueError) as exc:
            raise ProjetoAmostraErro(
                f"A amostra manual de {codigo} precisa ser inteira."
            ) from exc
        if not amostra_float.is_integer() or amostra_float < 0:
            raise ProjetoAmostraErro(
                f"A amostra manual de {codigo} precisa ser um inteiro não negativo."
            )
        normalizados.append({"codigo": codigo, "amostra": int(amostra_float)})
    return normalizados


def criar_configuracao_projeto(
    *,
    nome: str,
    fonte_dados: str,
    modo: str,
    amostra_total: int,
    base_populacional: str,
    faixa_flex_pct: int,
    uf: str | None,
    regioes_intermediarias: list[dict[str, str]] | None = None,
    regioes_imediatas: list[dict[str, str]] | None = None,
    municipios: list[dict[str, str]] | None = None,
    distritos: list[dict[str, str]] | None = None,
    nivel_quebra: str | None = None,
    incluir_distritos: bool = False,
    camadas: list[str] | None = None,
    metodo_distribuicao: str = "proporcional",
    alocacoes_manuais: list[dict[str, Any]] | None = None,
    exibicao_cotas: str = "ambos",
) -> dict[str, Any]:
    """Cria e valida a representação portável de um projeto de amostra."""
    configuracao = {
        "schema_version": VERSAO_SCHEMA_PROJETO,
        "tipo": "projeto_amostra",
        "nome": str(nome or "Projeto de amostra").strip(),
        "salvo_em_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "fonte_dados": fonte_dados,
        "modo": modo,
        "amostra_total": amostra_total,
        "base_populacional": base_populacional,
        "faixa_flex_pct": faixa_flex_pct,
        "territorio": {
            "uf": uf,
            "regioes_intermediarias": regioes_intermediarias or [],
            "regioes_imediatas": regioes_imediatas or [],
            "municipios": municipios or [],
            "distritos": distritos or [],
        },
        "classico": {
            "nivel_quebra": nivel_quebra,
            "incluir_distritos": bool(incluir_distritos),
        },
        "camadas": list(camadas or []),
        "distribuicao": {
            "metodo": metodo_distribuicao,
            "alocacoes_manuais": alocacoes_manuais or [],
        },
        "exportacao": {
            "exibicao_cotas": exibicao_cotas,
        },
    }
    return validar_configuracao_projeto(configuracao)


def validar_configuracao_projeto(configuracao: Mapping[str, Any]) -> dict[str, Any]:
    """Valida, migra e normaliza um dicionário carregado de JSON."""
    if not isinstance(configuracao, Mapping):
        raise ProjetoAmostraErro("O arquivo precisa conter um objeto JSON.")

    dados = deepcopy(dict(configuracao))
    if dados.get("tipo") != "projeto_amostra":
        raise ProjetoAmostraErro("O JSON não é um projeto de amostra reconhecido.")

    versao_recebida = str(dados.get("schema_version") or "")
    if versao_recebida not in VERSOES_SCHEMA_COMPATIVEIS:
        raise ProjetoAmostraErro(
            "Versão de projeto incompatível. "
            f"Compatíveis: {', '.join(sorted(VERSOES_SCHEMA_COMPATIVEIS))}; "
            f"recebida: {versao_recebida!r}."
        )

    nome = str(dados.get("nome") or "Projeto de amostra").strip()
    if len(nome) > 150:
        raise ProjetoAmostraErro("O nome do projeto deve ter no máximo 150 caracteres.")

    fonte = str(dados.get("fonte_dados") or "")
    if fonte not in FONTES_DADOS_VALIDAS:
        raise ProjetoAmostraErro(f"Fonte de dados inválida: {fonte!r}.")

    modo = str(dados.get("modo") or "")
    if modo not in MODOS_VALIDOS:
        raise ProjetoAmostraErro(f"Modo de montagem inválido: {modo!r}.")

    base = str(dados.get("base_populacional") or "")
    if base not in BASES_POPULACIONAIS_VALIDAS:
        raise ProjetoAmostraErro(f"Base populacional inválida: {base!r}.")

    territorio = dados.get("territorio") or {}
    if not isinstance(territorio, Mapping):
        raise ProjetoAmostraErro("O campo territorio precisa ser um objeto JSON.")
    uf = territorio.get("uf")
    if uf is not None:
        uf = str(uf).strip().upper()
        if not re.fullmatch(r"[A-Z]{2}", uf):
            raise ProjetoAmostraErro("A UF precisa ter duas letras ou ser nula para Brasil.")

    recortes = {
        campo: _normalizar_recorte_territorial(territorio.get(campo), campo)
        for campo in CAMPOS_RECORTES_TERRITORIAIS
    }

    classico = dados.get("classico") or {}
    if not isinstance(classico, Mapping):
        raise ProjetoAmostraErro("O campo classico precisa ser um objeto JSON.")
    nivel_quebra = classico.get("nivel_quebra")
    if nivel_quebra is not None and nivel_quebra not in NIVEIS_VALIDOS - {"uf"}:
        raise ProjetoAmostraErro(f"Nível de quebra inválido: {nivel_quebra!r}.")
    incluir_distritos = bool(classico.get("incluir_distritos", False))
    if nivel_quebra == "distrito" or recortes["distritos"]:
        incluir_distritos = True

    camadas = dados.get("camadas") or []
    if not isinstance(camadas, list):
        raise ProjetoAmostraErro("O campo camadas precisa ser uma lista.")
    camadas = [str(camada) for camada in camadas]
    invalidas = [camada for camada in camadas if camada not in NIVEIS_VALIDOS]
    if invalidas:
        raise ProjetoAmostraErro(f"Camada(s) inválida(s): {', '.join(invalidas)}.")
    if len(camadas) != len(set(camadas)):
        raise ProjetoAmostraErro("Uma camada não pode aparecer mais de uma vez.")
    if modo == "camadas" and not camadas:
        raise ProjetoAmostraErro("Projetos no modo por camadas precisam de ao menos uma camada.")
    if modo == "classico" and nivel_quebra is None:
        raise ProjetoAmostraErro("Projetos no modo clássico precisam de um nível de quebra.")

    distribuicao = dados.get("distribuicao") or {}
    if not isinstance(distribuicao, Mapping):
        raise ProjetoAmostraErro("O campo distribuicao precisa ser um objeto JSON.")
    metodo_distribuicao = str(distribuicao.get("metodo") or "proporcional")
    if metodo_distribuicao not in METODOS_DISTRIBUICAO_VALIDOS:
        raise ProjetoAmostraErro(
            f"Método de distribuição inválido: {metodo_distribuicao!r}."
        )
    alocacoes_manuais = _normalizar_alocacoes_manuais(
        distribuicao.get("alocacoes_manuais")
    )

    exportacao = dados.get("exportacao") or {}
    if not isinstance(exportacao, Mapping):
        raise ProjetoAmostraErro("O campo exportacao precisa ser um objeto JSON.")
    exibicao_cotas = str(exportacao.get("exibicao_cotas") or "ambos")
    if exibicao_cotas not in EXIBICOES_COTAS_VALIDAS:
        raise ProjetoAmostraErro(f"Exibição de cotas inválida: {exibicao_cotas!r}.")

    amostra_total_normalizada = _inteiro_positivo(
        dados.get("amostra_total"), "amostra_total"
    )
    if alocacoes_manuais and sum(item["amostra"] for item in alocacoes_manuais) != amostra_total_normalizada:
        raise ProjetoAmostraErro(
            "As alocações manuais salvas precisam somar exatamente amostra_total."
        )

    return {
        "schema_version": VERSAO_SCHEMA_PROJETO,
        "tipo": "projeto_amostra",
        "nome": nome,
        "salvo_em_utc": dados.get("salvo_em_utc"),
        "fonte_dados": fonte,
        "modo": modo,
        "amostra_total": amostra_total_normalizada,
        "base_populacional": base,
        "faixa_flex_pct": _percentual(dados.get("faixa_flex_pct", 20), "faixa_flex_pct"),
        "territorio": {
            "uf": uf,
            **recortes,
        },
        "classico": {
            "nivel_quebra": nivel_quebra,
            "incluir_distritos": incluir_distritos,
        },
        "camadas": camadas,
        "distribuicao": {
            "metodo": metodo_distribuicao,
            "alocacoes_manuais": alocacoes_manuais,
        },
        "exportacao": {
            "exibicao_cotas": exibicao_cotas,
        },
    }


def serializar_configuracao_projeto(configuracao: Mapping[str, Any]) -> bytes:
    """Valida e converte uma configuração para JSON UTF-8 legível."""
    normalizada = validar_configuracao_projeto(configuracao)
    texto = json.dumps(normalizada, ensure_ascii=False, indent=2, sort_keys=False)
    return (texto + "\n").encode("utf-8")


def desserializar_configuracao_projeto(conteudo: bytes | str) -> dict[str, Any]:
    """Carrega bytes/texto JSON e devolve a configuração normalizada."""
    try:
        texto = conteudo.decode("utf-8-sig") if isinstance(conteudo, bytes) else str(conteudo)
        dados = json.loads(texto)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProjetoAmostraErro(f"Não foi possível ler o JSON: {exc}") from exc
    return validar_configuracao_projeto(dados)


def nome_arquivo_projeto(nome: str) -> str:
    """Gera um nome de arquivo estável sem depender de bibliotecas externas."""
    base = re.sub(r"[^a-zA-Z0-9_-]+", "_", str(nome).strip()).strip("_")
    return f"{base or 'projeto_amostra'}.json"
