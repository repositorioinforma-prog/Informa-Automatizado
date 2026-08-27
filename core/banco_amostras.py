"""Acesso ao banco territorial e demográfico usado na montagem de amostras.

Além das consultas de exploração, o módulo adapta município e distrito ao
contrato esperado pelos motores clássico e por camadas.
"""

from __future__ import annotations

from contextlib import closing
import re
import sqlite3
import unicodedata
from pathlib import Path
from typing import Iterable

import pandas as pd

CAMINHO_BANCO_PADRAO = (
    Path(__file__).resolve().parent.parent
    / "dados"
    / "banco_amostras"
    / "banco_amostras.sqlite"
)

NIVEIS_DEMOGRAFICOS = ("municipio", "distrito", "subdistrito", "bairro")
NIVEIS_MOTOR_AMOSTRA = ("municipio", "distrito")

ROTULOS_NIVEIS_BANCO = {
    "brasil": "Brasil",
    "grande_regiao": "Grande Região",
    "uf": "Estado (UF)",
    "regiao_intermediaria": "Região Intermediária",
    "regiao_imediata": "Região Imediata",
    "municipio": "Município",
    "distrito": "Distrito",
    "subdistrito": "Subdistrito",
    "bairro": "Bairro",
}


class BancoAmostrasErro(RuntimeError):
    """Erro de configuração ou consulta do banco de amostras."""


def _normalizar_texto(valor: str) -> str:
    texto = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[^a-zA-Z0-9]+", " ", texto.lower())
    return re.sub(r"\s+", " ", texto).strip()


def conectar_banco(
    caminho: str | Path | None = None,
    *,
    somente_leitura: bool = True,
) -> sqlite3.Connection:
    """Abre uma conexão SQLite com chaves estrangeiras e linhas nomeadas."""
    banco = Path(caminho) if caminho is not None else CAMINHO_BANCO_PADRAO
    banco = banco.expanduser().resolve()
    if not banco.exists():
        raise BancoAmostrasErro(f"Banco de amostras não encontrado: {banco}")

    conexao = sqlite3.connect(str(banco), timeout=30)
    conexao.row_factory = sqlite3.Row
    conexao.execute("PRAGMA foreign_keys = ON")
    if somente_leitura:
        conexao.execute("PRAGMA query_only = ON")
    return conexao


def verificar_integridade(caminho: str | Path | None = None) -> dict[str, object]:
    """Executa as verificações nativas de integridade e chaves estrangeiras."""
    with closing(conectar_banco(caminho)) as conexao:
        integridade = conexao.execute("PRAGMA integrity_check").fetchone()[0]
        chaves_orfas = conexao.execute("PRAGMA foreign_key_check").fetchall()
        versao = conexao.execute(
            "SELECT valor FROM metadados_banco WHERE chave = 'versao_schema'"
        ).fetchone()
    return {
        "integridade": integridade,
        "chaves_estrangeiras_invalidas": len(chaves_orfas),
        "versao_schema": versao[0] if versao else None,
    }


def resumo_banco(caminho: str | Path | None = None) -> pd.DataFrame:
    """Retorna a quantidade de unidades e disponibilidade por nível."""
    consulta = """
        SELECT
            nivel,
            COUNT(*) AS quantidade,
            SUM(dados_demograficos_disponiveis) AS com_demografia,
            SUM(codigo_duplicado) AS registros_com_codigo_duplicado
        FROM territorios
        GROUP BY nivel, ordem_nivel
        ORDER BY ordem_nivel, nivel
    """
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao)


def obter_metadados_banco(caminho: str | Path | None = None) -> pd.DataFrame:
    """Retorna os metadados de versão e finalidade do banco."""
    consulta = """
        SELECT chave, valor, descricao
        FROM metadados_banco
        ORDER BY chave
    """
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao)


def listar_fontes(caminho: str | Path | None = None) -> pd.DataFrame:
    """Lista as fontes ingeridas e suas referências temporais."""
    consulta = """
        SELECT
            fonte_id,
            nome_logico,
            nome_arquivo,
            nivel_territorial,
            formato,
            total_registros,
            ano_referencia,
            origem_declarada,
            data_ingestao_utc,
            checksum_sha256,
            observacoes
        FROM fontes
        ORDER BY fonte_id
    """
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao)


def listar_ufs(caminho: str | Path | None = None) -> pd.DataFrame:
    """Retorna as UFs disponíveis no banco em ordem alfabética."""
    consulta = """
        SELECT DISTINCT uf, nome_uf
        FROM vw_territorios
        WHERE uf IS NOT NULL
        ORDER BY uf
    """
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao)


def resumo_alertas(caminho: str | Path | None = None) -> pd.DataFrame:
    """Agrupa os alertas de qualidade por severidade e categoria."""
    consulta = """
        SELECT severidade, categoria, COUNT(*) AS quantidade
        FROM alertas_qualidade
        GROUP BY severidade, categoria
        ORDER BY
            CASE severidade WHEN 'alta' THEN 1 WHEN 'media' THEN 2 ELSE 3 END,
            quantidade DESC,
            categoria
    """
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao)


def buscar_territorios(
    nivel: str,
    *,
    texto: str | None = None,
    uf: str | None = None,
    codigo_municipio: str | None = None,
    somente_com_demografia: bool = True,
    limite: int = 200,
    caminho: str | Path | None = None,
) -> pd.DataFrame:
    """Localiza unidades territoriais e devolve suas chaves não ambíguas.

    Para bairros, use ``chave_territorial`` nas etapas seguintes, pois uma
    fonte recebida contém um código repetido com dois registros distintos.
    """
    if nivel not in NIVEIS_DEMOGRAFICOS:
        raise ValueError(f"Nível inválido: {nivel!r}. Opções: {NIVEIS_DEMOGRAFICOS}")
    if limite < 1 or limite > 10_000:
        raise ValueError("O limite deve estar entre 1 e 10.000.")

    filtros = ["nivel = ?"]
    parametros: list[object] = [nivel]
    if texto:
        filtros.append("nome_normalizado LIKE ?")
        parametros.append(f"%{_normalizar_texto(texto)}%")
    if uf:
        filtros.append("UPPER(uf) = UPPER(?)")
        parametros.append(uf.strip())
    if codigo_municipio:
        filtros.append("codigo_municipio = ?")
        parametros.append(str(codigo_municipio).strip())
    if somente_com_demografia:
        filtros.append("dados_demograficos_disponiveis = 1")

    consulta = f"""
        SELECT
            territorio_id,
            chave_territorial,
            nivel,
            codigo,
            ocorrencia_codigo,
            nome,
            uf,
            nome_uf,
            codigo_municipio,
            nome_municipio,
            codigo_distrito,
            nome_distrito,
            codigo_duplicado,
            dados_demograficos_disponiveis
        FROM vw_territorios
        WHERE {' AND '.join(filtros)}
        ORDER BY nome_normalizado, codigo, ocorrencia_codigo
        LIMIT ?
    """
    parametros.append(limite)
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao, params=parametros)


def carregar_base_amostragem(
    nivel: str,
    *,
    chaves_territoriais: Iterable[str] | None = None,
    uf: str | None = None,
    codigo_municipio: str | None = None,
    somente_com_demografia: bool = True,
    caminho: str | Path | None = None,
) -> pd.DataFrame:
    """Carrega a visão pronta para cálculos de amostra e cotas."""
    if nivel not in NIVEIS_DEMOGRAFICOS:
        raise ValueError(f"Nível inválido: {nivel!r}. Opções: {NIVEIS_DEMOGRAFICOS}")

    filtros = ["nivel = ?"]
    parametros: list[object] = [nivel]
    if chaves_territoriais is not None:
        chaves = [str(chave).strip() for chave in chaves_territoriais]
        if not chaves:
            return pd.DataFrame()
        filtros.append(f"chave_territorial IN ({','.join('?' for _ in chaves)})")
        parametros.extend(chaves)
    if uf:
        filtros.append("UPPER(uf) = UPPER(?)")
        parametros.append(uf.strip())
    if codigo_municipio:
        filtros.append("codigo_municipio = ?")
        parametros.append(str(codigo_municipio).strip())
    if somente_com_demografia:
        filtros.append("dados_demograficos_disponiveis = 1")

    consulta = f"""
        SELECT *
        FROM vw_base_amostragem
        WHERE {' AND '.join(filtros)}
        ORDER BY codigo, ocorrencia_codigo
    """
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao, params=parametros)


def consultar_cobertura(
    *,
    nivel_pai: str,
    codigo_pai: str,
    nivel_filho: str,
    caminho: str | Path | None = None,
) -> pd.DataFrame:
    """Mostra se a soma dos filhos cobre a população do território pai."""
    consulta = """
        SELECT
            t.chave_territorial,
            t.nivel,
            t.codigo,
            t.nome,
            t.uf,
            c.nivel_filho,
            c.aplicavel,
            c.quantidade_registros_filhos,
            c.quantidade_codigos_filhos,
            c.quantidade_filhos_com_dados,
            c.populacao_pai,
            c.populacao_filhos,
            c.diferenca_populacao,
            c.proporcao_cobertura,
            c.fecha_exatamente,
            c.observacoes
        FROM cobertura_hierarquia c
        JOIN vw_territorios t ON t.territorio_id = c.territorio_pai_id
        WHERE t.nivel = ? AND t.codigo = ? AND c.nivel_filho = ?
        ORDER BY t.ocorrencia_codigo
    """
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(
            consulta,
            conexao,
            params=[nivel_pai, str(codigo_pai).strip(), nivel_filho],
        )


def listar_alertas(
    *,
    categoria: str | None = None,
    severidade: str | None = None,
    limite: int = 1_000,
    caminho: str | Path | None = None,
) -> pd.DataFrame:
    """Retorna alertas de qualidade preservados durante a ingestão."""
    filtros = ["1 = 1"]
    parametros: list[object] = []
    if categoria:
        filtros.append("categoria = ?")
        parametros.append(categoria)
    if severidade:
        filtros.append("severidade = ?")
        parametros.append(severidade)
    consulta = f"""
        SELECT * FROM vw_alertas_qualidade
        WHERE {' AND '.join(filtros)}
        ORDER BY
            CASE severidade WHEN 'alta' THEN 1 WHEN 'media' THEN 2 ELSE 3 END,
            categoria, nivel, codigo
        LIMIT ?
    """
    parametros.append(limite)
    with closing(conectar_banco(caminho)) as conexao:
        return pd.read_sql_query(consulta, conexao, params=parametros)


def carregar_base_motor_amostragem(
    nivel: str,
    *,
    somente_com_demografia: bool = True,
    caminho: str | Path | None = None,
) -> pd.DataFrame:
    """Converte o banco padronizado para o contrato usado pelos motores atuais.

    Nesta primeira integração, o motor aceita município e distrito. Bairro e
    subdistrito continuam disponíveis no explorador, mas só entrarão no cálculo
    quando houver uma política explícita para cobertura territorial parcial.
    """
    if nivel not in NIVEIS_MOTOR_AMOSTRA:
        raise ValueError(
            f"Nível inválido para o motor: {nivel!r}. "
            f"Opções: {NIVEIS_MOTOR_AMOSTRA}"
        )

    filtros = ["nivel = ?"]
    parametros: list[object] = [nivel]
    if somente_com_demografia:
        filtros.append("dados_demograficos_disponiveis = 1")

    consulta = f"""
        SELECT
            chave_territorial,
            uf,
            nome_uf,
            codigo_regiao_intermediaria,
            nome_regiao_intermediaria,
            codigo_regiao_imediata,
            nome_regiao_imediata,
            codigo_municipio,
            nome_municipio,
            codigo_distrito,
            nome_distrito,
            populacao_total,
            masculino_total,
            feminino_total,
            masculino_15_19,
            feminino_15_19,
            masculino_20_24,
            masculino_25_29,
            masculino_30_39,
            masculino_40_49,
            masculino_50_59,
            masculino_60_69,
            masculino_70_mais,
            feminino_20_24,
            feminino_25_29,
            feminino_30_39,
            feminino_40_49,
            feminino_50_59,
            feminino_60_69,
            feminino_70_mais,
            populacao_16_mais_referencia,
            masculino_16_mais_referencia,
            feminino_16_mais_referencia,
            idade_16_19_referencia,
            idade_20_29,
            "idade_30_39:1" AS idade_30_39_referencia,
            "idade_40_49:1" AS idade_40_49_referencia,
            idade_50_mais,
            metodo_populacao_16_mais_referencia,
            renda_ate_2sm,
            renda_2_a_5sm,
            renda_5_a_10sm,
            renda_mais_10sm,
            renda_disponivel
        FROM vw_base_amostragem
        WHERE {' AND '.join(filtros)}
        ORDER BY codigo, ocorrencia_codigo
    """
    with closing(conectar_banco(caminho)) as conexao:
        df = pd.read_sql_query(consulta, conexao, params=parametros)

    if df.empty:
        return df

    renomear = {
        "codigo_regiao_intermediaria": "cod_regiao_intermediaria",
        "nome_regiao_intermediaria": "regiao_intermediaria",
        "codigo_regiao_imediata": "cod_regiao_imediata",
        "nome_regiao_imediata": "regiao_imediata",
        "codigo_municipio": "cod_municipio",
        "nome_municipio": "municipio",
        "codigo_distrito": "cod_distrito",
        "nome_distrito": "distrito",
        "masculino_total": "masc_total",
        "feminino_total": "fem_total",
        "masculino_16_mais_referencia": "masc_16mais",
        "feminino_16_mais_referencia": "fem_16mais",
    }
    df = df.rename(columns=renomear)

    colunas_numericas = [
        "populacao_total",
        "masc_total",
        "fem_total",
        "masculino_15_19",
        "feminino_15_19",
        "masculino_20_24",
        "masculino_25_29",
        "masculino_30_39",
        "masculino_40_49",
        "masculino_50_59",
        "masculino_60_69",
        "masculino_70_mais",
        "feminino_20_24",
        "feminino_25_29",
        "feminino_30_39",
        "feminino_40_49",
        "feminino_50_59",
        "feminino_60_69",
        "feminino_70_mais",
        "populacao_16_mais_referencia",
        "masc_16mais",
        "fem_16mais",
        "idade_16_19_referencia",
        "idade_20_29",
        "idade_30_39_referencia",
        "idade_40_49_referencia",
        "idade_50_mais",
        "renda_ate_2sm",
        "renda_2_a_5sm",
        "renda_5_a_10sm",
        "renda_mais_10sm",
        "renda_disponivel",
    ]
    for coluna in colunas_numericas:
        if coluna in df.columns:
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    # Divide a faixa 16-19 de referência segundo a composição sexual observada
    # em 15-19. O total 16-19 permanece exatamente igual ao valor de referência.
    denominador_15_19 = df["masculino_15_19"].fillna(0) + df["feminino_15_19"].fillna(0)
    proporcao_masc = (df["masculino_15_19"].fillna(0) / denominador_15_19.where(denominador_15_19 > 0))
    proporcao_masc = proporcao_masc.fillna(0.5).clip(0, 1)
    total_16_19 = df["idade_16_19_referencia"].fillna(0).round().astype(int)
    df["masc_16_19"] = (total_16_19 * proporcao_masc).round().astype(int)
    df["fem_16_19"] = total_16_19 - df["masc_16_19"]

    def dividir_total_por_sexo(total, masculino, feminino):
        denominador = masculino.fillna(0) + feminino.fillna(0)
        proporcao = (masculino.fillna(0) / denominador.where(denominador > 0)).fillna(0.5)
        total_inteiro = total.fillna(0).round().astype(int)
        parte_masculina = (total_inteiro * proporcao.clip(0, 1)).round().astype(int)
        return parte_masculina, total_inteiro - parte_masculina

    masc_20_29_bruto = df[["masculino_20_24", "masculino_25_29"]].sum(axis=1, min_count=1)
    fem_20_29_bruto = df[["feminino_20_24", "feminino_25_29"]].sum(axis=1, min_count=1)
    df["masc_20_29"], df["fem_20_29"] = dividir_total_por_sexo(
        df["idade_20_29"], masc_20_29_bruto, fem_20_29_bruto
    )
    df["masc_30_39"], df["fem_30_39"] = dividir_total_por_sexo(
        df["idade_30_39_referencia"], df["masculino_30_39"], df["feminino_30_39"]
    )
    df["masc_40_49"], df["fem_40_49"] = dividir_total_por_sexo(
        df["idade_40_49_referencia"], df["masculino_40_49"], df["feminino_40_49"]
    )
    masc_50_bruto = df[
        ["masculino_50_59", "masculino_60_69", "masculino_70_mais"]
    ].sum(axis=1, min_count=1)
    fem_50_bruto = df[
        ["feminino_50_59", "feminino_60_69", "feminino_70_mais"]
    ].sum(axis=1, min_count=1)
    df["masc_50mais"], df["fem_50mais"] = dividir_total_por_sexo(
        df["idade_50_mais"], masc_50_bruto, fem_50_bruto
    )

    # Códigos precisam permanecer como texto para não perder zeros ou formato.
    for coluna in [
        "cod_regiao_intermediaria",
        "cod_regiao_imediata",
        "cod_municipio",
        "cod_distrito",
    ]:
        if coluna in df.columns:
            df[coluna] = df[coluna].astype("string")

    colunas_saida = [
        "chave_territorial",
        "uf",
        "nome_uf",
        "cod_regiao_intermediaria",
        "regiao_intermediaria",
        "cod_regiao_imediata",
        "regiao_imediata",
        "cod_municipio",
        "municipio",
    ]
    if nivel == "distrito":
        colunas_saida.extend(["cod_distrito", "distrito"])
    colunas_saida.extend([
        "populacao_total",
        "masc_total",
        "fem_total",
        "masc_16mais",
        "fem_16mais",
        "masc_16_19",
        "fem_16_19",
        "masc_20_29",
        "fem_20_29",
        "masc_30_39",
        "fem_30_39",
        "masc_40_49",
        "fem_40_49",
        "masc_50mais",
        "fem_50mais",
        "renda_ate_2sm",
        "renda_2_a_5sm",
        "renda_5_a_10sm",
        "renda_mais_10sm",
        "renda_disponivel",
        "populacao_16_mais_referencia",
        "metodo_populacao_16_mais_referencia",
    ])
    return df[colunas_saida].copy()
