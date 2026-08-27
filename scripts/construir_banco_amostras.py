#!/usr/bin/env python3
"""Constrói o banco SQLite de apoio à montagem de amostras territoriais.

O banco preserva os agregados demográficos recebidos, adiciona a hierarquia
territorial já existente no projeto, registra procedência e cria visões
amigáveis para uso no Streamlit.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

VAR_COLUMNS = [f"V010{i:02d}" for i in range(6, 42)]
VAR_COLUMNS_SQL = [c.lower() for c in VAR_COLUMNS]

UF_INFO = {
    "11": ("RO", "Rondônia", "Norte", "1"),
    "12": ("AC", "Acre", "Norte", "1"),
    "13": ("AM", "Amazonas", "Norte", "1"),
    "14": ("RR", "Roraima", "Norte", "1"),
    "15": ("PA", "Pará", "Norte", "1"),
    "16": ("AP", "Amapá", "Norte", "1"),
    "17": ("TO", "Tocantins", "Norte", "1"),
    "21": ("MA", "Maranhão", "Nordeste", "2"),
    "22": ("PI", "Piauí", "Nordeste", "2"),
    "23": ("CE", "Ceará", "Nordeste", "2"),
    "24": ("RN", "Rio Grande do Norte", "Nordeste", "2"),
    "25": ("PB", "Paraíba", "Nordeste", "2"),
    "26": ("PE", "Pernambuco", "Nordeste", "2"),
    "27": ("AL", "Alagoas", "Nordeste", "2"),
    "28": ("SE", "Sergipe", "Nordeste", "2"),
    "29": ("BA", "Bahia", "Nordeste", "2"),
    "31": ("MG", "Minas Gerais", "Sudeste", "3"),
    "32": ("ES", "Espírito Santo", "Sudeste", "3"),
    "33": ("RJ", "Rio de Janeiro", "Sudeste", "3"),
    "35": ("SP", "São Paulo", "Sudeste", "3"),
    "41": ("PR", "Paraná", "Sul", "4"),
    "42": ("SC", "Santa Catarina", "Sul", "4"),
    "43": ("RS", "Rio Grande do Sul", "Sul", "4"),
    "50": ("MS", "Mato Grosso do Sul", "Centro-Oeste", "5"),
    "51": ("MT", "Mato Grosso", "Centro-Oeste", "5"),
    "52": ("GO", "Goiás", "Centro-Oeste", "5"),
    "53": ("DF", "Distrito Federal", "Centro-Oeste", "5"),
}

GRANDES_REGIOES = {
    "1": "Norte",
    "2": "Nordeste",
    "3": "Sudeste",
    "4": "Sul",
    "5": "Centro-Oeste",
}

SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE metadados_banco (
    chave TEXT PRIMARY KEY,
    valor TEXT NOT NULL,
    descricao TEXT
);

CREATE TABLE fontes (
    fonte_id INTEGER PRIMARY KEY,
    nome_logico TEXT NOT NULL,
    nome_arquivo TEXT,
    tipo_conteudo TEXT NOT NULL,
    nivel_territorial TEXT,
    formato TEXT,
    planilha TEXT,
    encoding TEXT,
    checksum_sha256 TEXT,
    tamanho_bytes INTEGER,
    total_registros INTEGER,
    ano_referencia TEXT,
    origem_declarada TEXT,
    data_ingestao_utc TEXT NOT NULL,
    observacoes TEXT
);

CREATE TABLE ufs (
    codigo_uf TEXT PRIMARY KEY,
    sigla TEXT NOT NULL UNIQUE,
    nome TEXT NOT NULL,
    grande_regiao_codigo TEXT NOT NULL,
    grande_regiao_nome TEXT NOT NULL
);

CREATE TABLE territorios (
    territorio_id INTEGER PRIMARY KEY,
    chave_territorial TEXT NOT NULL UNIQUE,
    nivel TEXT NOT NULL,
    ordem_nivel INTEGER NOT NULL,
    codigo TEXT NOT NULL,
    ocorrencia_codigo INTEGER NOT NULL DEFAULT 1,
    nome TEXT NOT NULL,
    nome_normalizado TEXT NOT NULL,
    territorio_pai_id INTEGER,
    codigo_uf TEXT,
    codigo_regiao_intermediaria TEXT,
    nome_regiao_intermediaria TEXT,
    codigo_regiao_imediata TEXT,
    nome_regiao_imediata TEXT,
    codigo_municipio TEXT,
    nome_municipio TEXT,
    codigo_distrito TEXT,
    nome_distrito TEXT,
    fonte_territorio_id INTEGER,
    linha_origem INTEGER,
    codigo_duplicado INTEGER NOT NULL DEFAULT 0,
    dados_demograficos_disponiveis INTEGER NOT NULL DEFAULT 0,
    observacoes TEXT,
    FOREIGN KEY (territorio_pai_id) REFERENCES territorios(territorio_id),
    FOREIGN KEY (codigo_uf) REFERENCES ufs(codigo_uf),
    FOREIGN KEY (fonte_territorio_id) REFERENCES fontes(fonte_id),
    UNIQUE (nivel, codigo, ocorrencia_codigo)
);

CREATE TABLE demografia_agregada (
    territorio_id INTEGER PRIMARY KEY,
    fonte_id INTEGER NOT NULL,
    linha_origem INTEGER NOT NULL,
    v01006 INTEGER,
    v01007 INTEGER,
    v01008 INTEGER,
    v01009 INTEGER,
    v01010 INTEGER,
    v01011 INTEGER,
    v01012 INTEGER,
    v01013 INTEGER,
    v01014 INTEGER,
    v01015 INTEGER,
    v01016 INTEGER,
    v01017 INTEGER,
    v01018 INTEGER,
    v01019 INTEGER,
    v01020 INTEGER,
    v01021 INTEGER,
    v01022 INTEGER,
    v01023 INTEGER,
    v01024 INTEGER,
    v01025 INTEGER,
    v01026 INTEGER,
    v01027 INTEGER,
    v01028 INTEGER,
    v01029 INTEGER,
    v01030 INTEGER,
    v01031 INTEGER,
    v01032 INTEGER,
    v01033 INTEGER,
    v01034 INTEGER,
    v01035 INTEGER,
    v01036 INTEGER,
    v01037 INTEGER,
    v01038 INTEGER,
    v01039 INTEGER,
    v01040 INTEGER,
    v01041 INTEGER,
    FOREIGN KEY (territorio_id) REFERENCES territorios(territorio_id),
    FOREIGN KEY (fonte_id) REFERENCES fontes(fonte_id)
);

CREATE TABLE qualidade_demografia (
    territorio_id INTEGER PRIMARY KEY,
    celulas_ausentes INTEGER NOT NULL,
    diferenca_total_sexo INTEGER,
    diferenca_masculino_faixas INTEGER,
    diferenca_feminino_faixas INTEGER,
    diferenca_total_faixas INTEGER,
    fecha_total_sexo INTEGER,
    fecha_masculino_faixas INTEGER,
    fecha_feminino_faixas INTEGER,
    fecha_total_faixas INTEGER,
    FOREIGN KEY (territorio_id) REFERENCES territorios(territorio_id)
);

CREATE TABLE indicadores_amostragem (
    territorio_id INTEGER PRIMARY KEY,
    masculino_16_mais_municipal INTEGER,
    feminino_16_mais_municipal INTEGER,
    populacao_16_mais_municipal INTEGER,
    idade_16_19_municipal INTEGER,
    metodo_16_mais TEXT,
    renda_ate_2sm_bruta INTEGER,
    renda_2_a_5sm_bruta INTEGER,
    renda_5_a_10sm_bruta INTEGER,
    renda_mais_10sm_bruta INTEGER,
    renda_disponivel INTEGER NOT NULL DEFAULT 0,
    ano_referencia_renda INTEGER,
    fonte_idade_exata_id INTEGER,
    fonte_renda_id INTEGER,
    observacoes TEXT,
    FOREIGN KEY (territorio_id) REFERENCES territorios(territorio_id),
    FOREIGN KEY (fonte_idade_exata_id) REFERENCES fontes(fonte_id),
    FOREIGN KEY (fonte_renda_id) REFERENCES fontes(fonte_id)
);

CREATE TABLE dicionario_variaveis (
    variavel TEXT PRIMARY KEY,
    colecao TEXT NOT NULL,
    tipo TEXT,
    tema TEXT,
    descricao TEXT NOT NULL,
    utilizada_neste_banco INTEGER NOT NULL DEFAULT 0,
    nome_amigavel TEXT,
    grupo_amostragem TEXT,
    ordem_amostragem INTEGER
);

CREATE TABLE dicionario_siglas (
    sigla TEXT PRIMARY KEY,
    descricao TEXT NOT NULL,
    colecao TEXT NOT NULL
);

CREATE TABLE cobertura_hierarquia (
    territorio_pai_id INTEGER NOT NULL,
    nivel_filho TEXT NOT NULL,
    aplicavel INTEGER NOT NULL,
    quantidade_registros_filhos INTEGER NOT NULL,
    quantidade_codigos_filhos INTEGER NOT NULL,
    quantidade_filhos_com_dados INTEGER NOT NULL,
    populacao_pai INTEGER,
    populacao_filhos INTEGER,
    diferenca_populacao INTEGER,
    proporcao_cobertura REAL,
    fecha_exatamente INTEGER,
    observacoes TEXT,
    PRIMARY KEY (territorio_pai_id, nivel_filho),
    FOREIGN KEY (territorio_pai_id) REFERENCES territorios(territorio_id)
);

CREATE TABLE alertas_qualidade (
    alerta_id INTEGER PRIMARY KEY,
    severidade TEXT NOT NULL,
    categoria TEXT NOT NULL,
    territorio_id INTEGER,
    nivel TEXT,
    codigo TEXT,
    mensagem TEXT NOT NULL,
    detalhes_json TEXT,
    FOREIGN KEY (territorio_id) REFERENCES territorios(territorio_id)
);

CREATE TABLE metricas_qualidade (
    metrica TEXT NOT NULL,
    nivel TEXT NOT NULL DEFAULT '',
    valor_numerico REAL,
    valor_texto TEXT,
    observacoes TEXT,
    PRIMARY KEY (metrica, nivel)
);

CREATE INDEX idx_territorios_nivel_codigo ON territorios(nivel, codigo);
CREATE INDEX idx_territorios_nome_normalizado ON territorios(nome_normalizado);
CREATE INDEX idx_territorios_pai ON territorios(territorio_pai_id);
CREATE INDEX idx_territorios_uf ON territorios(codigo_uf, nivel);
CREATE INDEX idx_territorios_municipio ON territorios(codigo_municipio, nivel);
CREATE INDEX idx_territorios_distrito ON territorios(codigo_distrito, nivel);
CREATE INDEX idx_alertas_categoria ON alertas_qualidade(categoria, severidade);

CREATE VIEW vw_territorios AS
SELECT
    t.territorio_id,
    t.chave_territorial,
    t.nivel,
    t.ordem_nivel,
    t.codigo,
    t.ocorrencia_codigo,
    t.nome,
    t.nome_normalizado,
    t.territorio_pai_id,
    p.nivel AS nivel_pai,
    p.codigo AS codigo_pai,
    p.nome AS nome_pai,
    t.codigo_uf,
    u.sigla AS uf,
    u.nome AS nome_uf,
    u.grande_regiao_codigo,
    u.grande_regiao_nome,
    t.codigo_regiao_intermediaria,
    t.nome_regiao_intermediaria,
    t.codigo_regiao_imediata,
    t.nome_regiao_imediata,
    t.codigo_municipio,
    t.nome_municipio,
    t.codigo_distrito,
    t.nome_distrito,
    t.codigo_duplicado,
    t.dados_demograficos_disponiveis,
    t.observacoes
FROM territorios t
LEFT JOIN territorios p ON p.territorio_id = t.territorio_pai_id
LEFT JOIN ufs u ON u.codigo_uf = t.codigo_uf;

CREATE VIEW vw_demografia_amigavel AS
SELECT
    d.territorio_id,
    d.v01006 AS populacao_total,
    d.v01007 AS masculino_total,
    d.v01008 AS feminino_total,
    d.v01009 AS masculino_0_4,
    d.v01010 AS masculino_5_9,
    d.v01011 AS masculino_10_14,
    d.v01012 AS masculino_15_19,
    d.v01013 AS masculino_20_24,
    d.v01014 AS masculino_25_29,
    d.v01015 AS masculino_30_39,
    d.v01016 AS masculino_40_49,
    d.v01017 AS masculino_50_59,
    d.v01018 AS masculino_60_69,
    d.v01019 AS masculino_70_mais,
    d.v01020 AS feminino_0_4,
    d.v01021 AS feminino_5_9,
    d.v01022 AS feminino_10_14,
    d.v01023 AS feminino_15_19,
    d.v01024 AS feminino_20_24,
    d.v01025 AS feminino_25_29,
    d.v01026 AS feminino_30_39,
    d.v01027 AS feminino_40_49,
    d.v01028 AS feminino_50_59,
    d.v01029 AS feminino_60_69,
    d.v01030 AS feminino_70_mais,
    d.v01031 AS idade_0_4,
    d.v01032 AS idade_5_9,
    d.v01033 AS idade_10_14,
    d.v01034 AS idade_15_19,
    d.v01035 AS idade_20_24,
    d.v01036 AS idade_25_29,
    d.v01037 AS idade_30_39,
    d.v01038 AS idade_40_49,
    d.v01039 AS idade_50_59,
    d.v01040 AS idade_60_69,
    d.v01041 AS idade_70_mais,
    CASE WHEN d.v01034 IS NULL THEN NULL ELSE
        d.v01034 + d.v01035 + d.v01036 + d.v01037 + d.v01038 + d.v01039 + d.v01040 + d.v01041
    END AS populacao_15_mais,
    CASE WHEN d.v01034 IS NULL THEN NULL ELSE CAST(ROUND(
        d.v01034 * 0.8 + d.v01035 + d.v01036 + d.v01037 + d.v01038 + d.v01039 + d.v01040 + d.v01041
    ) AS INTEGER) END AS populacao_16_mais_estimada_uniforme,
    CASE WHEN d.v01012 IS NULL THEN NULL ELSE CAST(ROUND(
        d.v01012 * 0.8 + d.v01013 + d.v01014 + d.v01015 + d.v01016 + d.v01017 + d.v01018 + d.v01019
    ) AS INTEGER) END AS masculino_16_mais_estimado_uniforme,
    CASE WHEN d.v01023 IS NULL THEN NULL ELSE CAST(ROUND(
        d.v01023 * 0.8 + d.v01024 + d.v01025 + d.v01026 + d.v01027 + d.v01028 + d.v01029 + d.v01030
    ) AS INTEGER) END AS feminino_16_mais_estimado_uniforme,
    CASE WHEN d.v01035 IS NULL THEN NULL ELSE d.v01035 + d.v01036 END AS idade_20_29,
    d.v01037 AS idade_30_39,
    d.v01038 AS idade_40_49,
    CASE WHEN d.v01039 IS NULL THEN NULL ELSE d.v01039 + d.v01040 + d.v01041 END AS idade_50_mais,
    q.celulas_ausentes,
    q.diferenca_total_sexo,
    q.diferenca_masculino_faixas,
    q.diferenca_feminino_faixas,
    q.diferenca_total_faixas,
    q.fecha_total_sexo,
    q.fecha_total_faixas
FROM demografia_agregada d
LEFT JOIN qualidade_demografia q ON q.territorio_id = d.territorio_id;

CREATE VIEW vw_base_amostragem AS
SELECT
    t.*,
    d.populacao_total,
    d.masculino_total,
    d.feminino_total,
    d.masculino_0_4,
    d.masculino_5_9,
    d.masculino_10_14,
    d.masculino_15_19,
    d.masculino_20_24,
    d.masculino_25_29,
    d.masculino_30_39,
    d.masculino_40_49,
    d.masculino_50_59,
    d.masculino_60_69,
    d.masculino_70_mais,
    d.feminino_0_4,
    d.feminino_5_9,
    d.feminino_10_14,
    d.feminino_15_19,
    d.feminino_20_24,
    d.feminino_25_29,
    d.feminino_30_39,
    d.feminino_40_49,
    d.feminino_50_59,
    d.feminino_60_69,
    d.feminino_70_mais,
    d.idade_0_4,
    d.idade_5_9,
    d.idade_10_14,
    d.idade_15_19,
    d.idade_20_24,
    d.idade_25_29,
    d.idade_30_39,
    d.idade_40_49,
    d.idade_50_59,
    d.idade_60_69,
    d.idade_70_mais,
    d.populacao_15_mais,
    i.populacao_16_mais_municipal,
    d.populacao_16_mais_estimada_uniforme,
    COALESCE(i.populacao_16_mais_municipal, d.populacao_16_mais_estimada_uniforme) AS populacao_16_mais_referencia,
    CASE
        WHEN i.populacao_16_mais_municipal IS NOT NULL THEN 'municipal_sidra_16_19_mais_censo_20_mais'
        WHEN d.populacao_16_mais_estimada_uniforme IS NOT NULL THEN 'estimada_80_porcento_da_faixa_15_19'
        ELSE 'indisponivel'
    END AS metodo_populacao_16_mais_referencia,
    i.masculino_16_mais_municipal,
    i.feminino_16_mais_municipal,
    d.masculino_16_mais_estimado_uniforme,
    d.feminino_16_mais_estimado_uniforme,
    COALESCE(i.masculino_16_mais_municipal, d.masculino_16_mais_estimado_uniforme) AS masculino_16_mais_referencia,
    COALESCE(i.feminino_16_mais_municipal, d.feminino_16_mais_estimado_uniforme) AS feminino_16_mais_referencia,
    COALESCE(i.idade_16_19_municipal, CAST(ROUND(d.idade_15_19 * 0.8) AS INTEGER)) AS idade_16_19_referencia,
    d.idade_20_29,
    d.idade_30_39,
    d.idade_40_49,
    d.idade_50_mais,
    CASE WHEN i.renda_disponivel = 1 THEN i.renda_ate_2sm_bruta END AS renda_ate_2sm,
    CASE WHEN i.renda_disponivel = 1 THEN i.renda_2_a_5sm_bruta END AS renda_2_a_5sm,
    CASE WHEN i.renda_disponivel = 1 THEN i.renda_5_a_10sm_bruta END AS renda_5_a_10sm,
    CASE WHEN i.renda_disponivel = 1 THEN i.renda_mais_10sm_bruta END AS renda_mais_10sm,
    COALESCE(i.renda_disponivel, 0) AS renda_disponivel,
    i.ano_referencia_renda,
    d.celulas_ausentes,
    d.diferenca_total_sexo,
    d.diferenca_total_faixas,
    d.fecha_total_sexo,
    d.fecha_total_faixas
FROM vw_territorios t
JOIN vw_demografia_amigavel d ON d.territorio_id = t.territorio_id
LEFT JOIN indicadores_amostragem i ON i.territorio_id = t.territorio_id;

CREATE VIEW vw_amostra_municipios AS SELECT * FROM vw_base_amostragem WHERE nivel = 'municipio';
CREATE VIEW vw_amostra_distritos AS SELECT * FROM vw_base_amostragem WHERE nivel = 'distrito';
CREATE VIEW vw_amostra_subdistritos AS SELECT * FROM vw_base_amostragem WHERE nivel = 'subdistrito';
CREATE VIEW vw_amostra_bairros AS SELECT * FROM vw_base_amostragem WHERE nivel = 'bairro';

CREATE VIEW vw_alertas_qualidade AS
SELECT
    a.alerta_id,
    a.severidade,
    a.categoria,
    a.territorio_id,
    COALESCE(a.nivel, t.nivel) AS nivel,
    COALESCE(a.codigo, t.codigo) AS codigo,
    t.nome,
    t.codigo_uf,
    t.codigo_municipio,
    a.mensagem,
    a.detalhes_json
FROM alertas_qualidade a
LEFT JOIN territorios t ON t.territorio_id = a.territorio_id;
"""

FRIENDLY_VARIABLES = {
    "V01006": ("populacao_total", "populacao", 1),
    "V01007": ("masculino_total", "sexo", 2),
    "V01008": ("feminino_total", "sexo", 3),
    "V01009": ("masculino_0_4", "sexo_idade", 4),
    "V01010": ("masculino_5_9", "sexo_idade", 5),
    "V01011": ("masculino_10_14", "sexo_idade", 6),
    "V01012": ("masculino_15_19", "sexo_idade", 7),
    "V01013": ("masculino_20_24", "sexo_idade", 8),
    "V01014": ("masculino_25_29", "sexo_idade", 9),
    "V01015": ("masculino_30_39", "sexo_idade", 10),
    "V01016": ("masculino_40_49", "sexo_idade", 11),
    "V01017": ("masculino_50_59", "sexo_idade", 12),
    "V01018": ("masculino_60_69", "sexo_idade", 13),
    "V01019": ("masculino_70_mais", "sexo_idade", 14),
    "V01020": ("feminino_0_4", "sexo_idade", 15),
    "V01021": ("feminino_5_9", "sexo_idade", 16),
    "V01022": ("feminino_10_14", "sexo_idade", 17),
    "V01023": ("feminino_15_19", "sexo_idade", 18),
    "V01024": ("feminino_20_24", "sexo_idade", 19),
    "V01025": ("feminino_25_29", "sexo_idade", 20),
    "V01026": ("feminino_30_39", "sexo_idade", 21),
    "V01027": ("feminino_40_49", "sexo_idade", 22),
    "V01028": ("feminino_50_59", "sexo_idade", 23),
    "V01029": ("feminino_60_69", "sexo_idade", 24),
    "V01030": ("feminino_70_mais", "sexo_idade", 25),
    "V01031": ("idade_0_4", "idade", 26),
    "V01032": ("idade_5_9", "idade", 27),
    "V01033": ("idade_10_14", "idade", 28),
    "V01034": ("idade_15_19", "idade", 29),
    "V01035": ("idade_20_24", "idade", 30),
    "V01036": ("idade_25_29", "idade", 31),
    "V01037": ("idade_30_39", "idade", 32),
    "V01038": ("idade_40_49", "idade", 33),
    "V01039": ("idade_50_59", "idade", 34),
    "V01040": ("idade_60_69", "idade", 35),
    "V01041": ("idade_70_mais", "idade", 36),
}


@dataclass(frozen=True)
class InputPaths:
    municipios: Path
    distritos: Path
    subdistritos: Path
    bairros: Path
    dicionario: Path
    master_municipios: Path | None
    master_distritos: Path | None


def normalize_text(value: Any) -> str:
    text = "" if value is None else str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_code(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def safe_int(value: Any) -> int | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        if value in {"", "."}:
            return None
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def sum_if_complete(values: Iterable[Any]) -> int | None:
    converted = [safe_int(v) for v in values]
    if any(v is None for v in converted):
        return None
    return int(sum(converted))


def read_demography_excel(path: Path, code_col: str, name_col: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype={code_col: "string"}, keep_default_na=False)
    return prepare_demography(df, code_col, name_col)


def read_demography_csv(path: Path, code_col: str, name_col: str) -> pd.DataFrame:
    df = pd.read_csv(
        path,
        sep=";",
        encoding="latin1",
        dtype={code_col: "string"},
        keep_default_na=False,
    )
    return prepare_demography(df, code_col, name_col)


def prepare_demography(df: pd.DataFrame, code_col: str, name_col: str) -> pd.DataFrame:
    expected = [code_col, name_col, *VAR_COLUMNS]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas ausentes em {code_col}: {missing}")
    out = df[expected].copy()
    out[code_col] = normalize_code(out[code_col])
    out[name_col] = out[name_col].astype("string").str.strip()
    for col in VAR_COLUMNS:
        out[col] = pd.to_numeric(out[col].replace({".": pd.NA, "": pd.NA}), errors="coerce").astype("Int64")
    return out


def read_project_master(path: Path | None, code_col: str) -> pd.DataFrame:
    if path is None or not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path, dtype={code_col: "string"})
    df[code_col] = normalize_code(df[code_col])
    return df


def insert_many(conn: sqlite3.Connection, sql: str, rows: list[tuple[Any, ...]]) -> None:
    if rows:
        conn.executemany(sql, rows)


def source_record(
    source_id: int,
    logical_name: str,
    path: Path | None,
    content_type: str,
    level: str | None,
    file_format: str | None,
    sheet: str | None,
    encoding: str | None,
    total_records: int | None,
    reference_year: str | None,
    declared_origin: str,
    ingested_at: str,
    notes: str | None = None,
) -> tuple[Any, ...]:
    return (
        source_id,
        logical_name,
        path.name if path else None,
        content_type,
        level,
        file_format,
        sheet,
        encoding,
        sha256_file(path) if path and path.exists() else None,
        path.stat().st_size if path and path.exists() else None,
        total_records,
        reference_year,
        declared_origin,
        ingested_at,
        notes,
    )


def build_database(inputs: InputPaths, output: Path, reports_dir: Path | None = None) -> dict[str, Any]:
    municipios = read_demography_excel(inputs.municipios, "CD_MUN", "NM_MUN")
    distritos = read_demography_csv(inputs.distritos, "CD_DIST", "NM_DIST")
    subdistritos = read_demography_excel(inputs.subdistritos, "CD_SUBDIST", "NM_SUBDIST")
    bairros = read_demography_excel(inputs.bairros, "CD_BAIRRO", "NM_BAIRRO")
    master_mun = read_project_master(inputs.master_municipios, "cod_municipio")
    master_dist = read_project_master(inputs.master_distritos, "cod_distrito")

    dictionary_frames: dict[str, pd.DataFrame] = {}
    dictionary_record_count = 0
    dictionary_workbook = pd.ExcelFile(inputs.dicionario)
    for sheet in dictionary_workbook.sheet_names:
        dictionary_frame = pd.read_excel(inputs.dicionario, sheet_name=sheet, dtype="string")
        dictionary_frames[sheet] = dictionary_frame
        if "Variável" in dictionary_frame.columns:
            dictionary_record_count += len(dictionary_frame)

    for frame, code in [
        (distritos, "CD_DIST"),
        (subdistritos, "CD_SUBDIST"),
        (bairros, "CD_BAIRRO"),
    ]:
        frame[code] = normalize_code(frame[code])

    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    ingested_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    conn = sqlite3.connect(output)
    conn.execute("PRAGMA journal_mode = DELETE")
    conn.execute("PRAGMA synchronous = FULL")
    conn.executescript(SCHEMA_SQL)

    source_rows = [
        source_record(1, "Demografia por municípios", inputs.municipios, "agregado_demografico", "municipio", "xlsx", "Agregados_por_municipios_demogr", None, len(municipios), "Censo 2022", "IBGE — arquivo fornecido pelo usuário", ingested_at),
        source_record(2, "Demografia por distritos", inputs.distritos, "agregado_demografico", "distrito", "csv", None, "latin1; separador ponto e vírgula", len(distritos), "Censo 2022", "IBGE — arquivo fornecido pelo usuário", ingested_at, "O ponto (.) foi preservado semanticamente como valor indisponível/NULL."),
        source_record(3, "Demografia por subdistritos", inputs.subdistritos, "agregado_demografico", "subdistrito", "xlsx", "Agregados_por_subdistritos_demo", None, len(subdistritos), "Censo 2022", "IBGE — arquivo fornecido pelo usuário", ingested_at),
        source_record(4, "Demografia por bairros", inputs.bairros, "agregado_demografico", "bairro", "xlsx", "Agregados_por_bairros_demografi", None, len(bairros), "Censo 2022", "IBGE — arquivo fornecido pelo usuário", ingested_at),
        source_record(5, "Dicionário de dados de agregados", inputs.dicionario, "dicionario", None, "xlsx", "múltiplas", None, dictionary_record_count, "arquivo 2026-05-20", "IBGE — arquivo fornecido pelo usuário", ingested_at),
        source_record(6, "Base mestre municipal do projeto", inputs.master_municipios, "hierarquia_e_indicadores", "municipio", "csv", None, "utf-8", len(master_mun) if not master_mun.empty else 0, "DTB 2025; Censo 2022; renda 2010", "Projeto Mapa de Correspondência", ingested_at, "Usada para regiões intermediárias/imediatas, idade exata municipal 16–19 e renda."),
        source_record(7, "Base mestre distrital do projeto", inputs.master_distritos, "hierarquia_e_indicadores", "distrito", "csv", None, "utf-8", len(master_dist) if not master_dist.empty else 0, "DTB 2025; Censo 2022; renda 2010", "Projeto Mapa de Correspondência", ingested_at, "Usada para renda distrital; a faixa distrital originalmente rotulada 16–19 não foi tratada como exata."),
        source_record(8, "Mapa interno de UFs e grandes regiões", None, "referencia_territorial", "uf", "interno", None, None, len(UF_INFO), "vigente", "Tabela estática de códigos IBGE", ingested_at),
    ]
    insert_many(
        conn,
        """INSERT INTO fontes (
            fonte_id,nome_logico,nome_arquivo,tipo_conteudo,nivel_territorial,formato,
            planilha,encoding,checksum_sha256,tamanho_bytes,total_registros,ano_referencia,
            origem_declarada,data_ingestao_utc,observacoes
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        source_rows,
    )

    metadata = [
        ("versao_schema", "1.0.0", "Versão do esquema do banco."),
        ("data_criacao_utc", ingested_at, "Data de construção do arquivo SQLite."),
        ("finalidade", "Apoio à montagem e ao controle de amostras territoriais", None),
        ("formato", "SQLite 3", None),
        ("politica_valor_ponto", "NULL", "Valores representados por ponto (.) nas fontes foram convertidos em NULL, nunca em zero."),
        ("politica_codigo", "TEXT", "Códigos territoriais são armazenados como texto para preservar zeros e evitar notação científica."),
        ("politica_16_mais", "referencia_municipal; estimada_demais_niveis", "Municípios usam uma referência baseada na idade 16–19 do projeto; demais níveis usam estimativa explícita de 80% da faixa 15–19, somente nas visões."),
    ]
    insert_many(conn, "INSERT INTO metadados_banco(chave,valor,descricao) VALUES (?,?,?)", metadata)

    uf_rows = [(code, *UF_INFO[code][:2], UF_INFO[code][3], UF_INFO[code][2]) for code in sorted(UF_INFO)]
    insert_many(conn, "INSERT INTO ufs(codigo_uf,sigla,nome,grande_regiao_codigo,grande_regiao_nome) VALUES (?,?,?,?,?)", uf_rows)

    # Hierarquia municipal proveniente da base mestre do projeto.
    mun_hierarchy: dict[str, dict[str, str | None]] = {}
    if not master_mun.empty:
        for _, row in master_mun.iterrows():
            code = str(row["cod_municipio"])
            mun_hierarchy[code] = {
                "codigo_uf": code[:2],
                "cod_ri": str(row.get("cod_regiao_intermediaria", "")).replace(".0", "") if pd.notna(row.get("cod_regiao_intermediaria")) else None,
                "nome_ri": str(row.get("regiao_intermediaria", "")).strip() if pd.notna(row.get("regiao_intermediaria")) else None,
                "cod_rim": str(row.get("cod_regiao_imediata", "")).replace(".0", "") if pd.notna(row.get("cod_regiao_imediata")) else None,
                "nome_rim": str(row.get("regiao_imediata", "")).strip() if pd.notna(row.get("regiao_imediata")) else None,
            }

    territories: list[dict[str, Any]] = []
    key_to_id: dict[str, int] = {}
    next_id = 1

    def add_territory(
        key: str,
        level: str,
        order: int,
        code: str,
        name: str,
        parent_key: str | None,
        source_id: int,
        line_origin: int | None = None,
        occurrence: int = 1,
        duplicate: int = 0,
        data_available: int = 0,
        code_uf: str | None = None,
        code_ri: str | None = None,
        name_ri: str | None = None,
        code_rim: str | None = None,
        name_rim: str | None = None,
        code_mun: str | None = None,
        name_mun: str | None = None,
        code_dist: str | None = None,
        name_dist: str | None = None,
        notes: str | None = None,
    ) -> int:
        nonlocal next_id
        if key in key_to_id:
            return key_to_id[key]
        parent_id = key_to_id.get(parent_key) if parent_key else None
        tid = next_id
        next_id += 1
        key_to_id[key] = tid
        territories.append({
            "territorio_id": tid,
            "chave_territorial": key,
            "nivel": level,
            "ordem_nivel": order,
            "codigo": str(code),
            "ocorrencia_codigo": occurrence,
            "nome": str(name).strip(),
            "nome_normalizado": normalize_text(name),
            "territorio_pai_id": parent_id,
            "codigo_uf": code_uf,
            "codigo_regiao_intermediaria": code_ri,
            "nome_regiao_intermediaria": name_ri,
            "codigo_regiao_imediata": code_rim,
            "nome_regiao_imediata": name_rim,
            "codigo_municipio": code_mun,
            "nome_municipio": name_mun,
            "codigo_distrito": code_dist,
            "nome_distrito": name_dist,
            "fonte_territorio_id": source_id,
            "linha_origem": line_origin,
            "codigo_duplicado": duplicate,
            "dados_demograficos_disponiveis": data_available,
            "observacoes": notes,
        })
        return tid

    add_territory("brasil:BR", "brasil", 0, "BR", "Brasil", None, 8)
    for region_code, region_name in GRANDES_REGIOES.items():
        add_territory(f"grande_regiao:{region_code}", "grande_regiao", 1, region_code, region_name, "brasil:BR", 8)
    for uf_code, (sigla, uf_name, region_name, region_code) in sorted(UF_INFO.items()):
        add_territory(
            f"uf:{uf_code}", "uf", 2, uf_code, uf_name,
            f"grande_regiao:{region_code}", 8, code_uf=uf_code,
        )

    # Regiões intermediárias e imediatas.
    if not master_mun.empty:
        region_intermediate = (
            master_mun[["cod_regiao_intermediaria", "regiao_intermediaria", "uf"]]
            .drop_duplicates()
            .sort_values("cod_regiao_intermediaria")
        )
        for _, row in region_intermediate.iterrows():
            code = str(row["cod_regiao_intermediaria"]).replace(".0", "")
            uf_sigla = str(row["uf"])
            uf_code = next((c for c, info in UF_INFO.items() if info[0] == uf_sigla), code[:2])
            add_territory(
                f"regiao_intermediaria:{code}", "regiao_intermediaria", 3, code,
                str(row["regiao_intermediaria"]), f"uf:{uf_code}", 6,
                code_uf=uf_code, code_ri=code, name_ri=str(row["regiao_intermediaria"]),
            )

        region_immediate = (
            master_mun[["cod_regiao_imediata", "regiao_imediata", "cod_regiao_intermediaria", "regiao_intermediaria", "uf"]]
            .drop_duplicates()
            .sort_values("cod_regiao_imediata")
        )
        for _, row in region_immediate.iterrows():
            code = str(row["cod_regiao_imediata"]).replace(".0", "")
            ri_code = str(row["cod_regiao_intermediaria"]).replace(".0", "")
            uf_sigla = str(row["uf"])
            uf_code = next((c for c, info in UF_INFO.items() if info[0] == uf_sigla), code[:2])
            add_territory(
                f"regiao_imediata:{code}", "regiao_imediata", 4, code,
                str(row["regiao_imediata"]), f"regiao_intermediaria:{ri_code}", 6,
                code_uf=uf_code, code_ri=ri_code, name_ri=str(row["regiao_intermediaria"]),
                code_rim=code, name_rim=str(row["regiao_imediata"]),
            )

    demography_rows: list[tuple[Any, ...]] = []
    quality_rows: list[tuple[Any, ...]] = []

    def add_demography_for_frame(
        frame: pd.DataFrame,
        level: str,
        code_col: str,
        name_col: str,
        order: int,
        source_id: int,
    ) -> None:
        duplicate_counts = frame[code_col].value_counts()
        occurrence = frame.groupby(code_col).cumcount() + 1
        for pos, (_, row) in enumerate(frame.iterrows(), start=2):
            code = str(row[code_col])
            name = str(row[name_col])
            occ = int(occurrence.iloc[pos - 2])
            is_duplicate = int(duplicate_counts.get(code, 0) > 1)
            key = f"{level}:{code}" if not is_duplicate else f"{level}:{code}:{occ}"
            code_uf = code[:2]
            code_mun = code[:7] if level != "municipio" else code
            mun_name = None
            if level == "municipio":
                mun_name = name
            else:
                match = municipios.loc[municipios["CD_MUN"] == code_mun, "NM_MUN"]
                mun_name = str(match.iloc[0]) if not match.empty else None

            hierarchy = mun_hierarchy.get(code_mun, {})
            code_ri = hierarchy.get("cod_ri")
            name_ri = hierarchy.get("nome_ri")
            code_rim = hierarchy.get("cod_rim")
            name_rim = hierarchy.get("nome_rim")

            if level == "municipio":
                parent_key = f"regiao_imediata:{code_rim}" if code_rim and f"regiao_imediata:{code_rim}" in key_to_id else f"uf:{code_uf}"
                code_dist = None
                dist_name = None
            elif level == "distrito":
                parent_key = f"municipio:{code_mun}"
                code_dist = code
                dist_name = name
            elif level == "subdistrito":
                code_dist = code[:9]
                parent_key = f"distrito:{code_dist}"
                dist_match = distritos.loc[distritos["CD_DIST"] == code_dist, "NM_DIST"]
                dist_name = str(dist_match.iloc[0]) if not dist_match.empty else None
            elif level == "bairro":
                parent_key = f"municipio:{code_mun}"
                code_dist = None
                dist_name = None
            else:
                raise ValueError(level)

            values = [safe_int(row[col]) for col in VAR_COLUMNS]
            data_available = int(values[0] is not None)
            notes = None
            if is_duplicate:
                notes = "Código repetido na fonte; registro preservado com ocorrência própria."
            tid = add_territory(
                key, level, order, code, name, parent_key, source_id,
                line_origin=pos, occurrence=occ, duplicate=is_duplicate,
                data_available=data_available, code_uf=code_uf,
                code_ri=code_ri, name_ri=name_ri, code_rim=code_rim, name_rim=name_rim,
                code_mun=code_mun, name_mun=mun_name, code_dist=code_dist, name_dist=dist_name,
                notes=notes,
            )
            demography_rows.append((tid, source_id, pos, *values))
            missing_cells = sum(v is None for v in values)
            if data_available:
                total, male, female = values[0], values[1], values[2]
                male_ages = values[3:14]
                female_ages = values[14:25]
                total_ages = values[25:36]
                diff_sex = total - male - female if None not in (total, male, female) else None
                diff_male = male - sum(male_ages) if male is not None and all(v is not None for v in male_ages) else None
                diff_female = female - sum(female_ages) if female is not None and all(v is not None for v in female_ages) else None
                diff_total_age = total - sum(total_ages) if total is not None and all(v is not None for v in total_ages) else None
            else:
                diff_sex = diff_male = diff_female = diff_total_age = None
            quality_rows.append((
                tid, missing_cells, diff_sex, diff_male, diff_female, diff_total_age,
                None if diff_sex is None else int(diff_sex == 0),
                None if diff_male is None else int(diff_male == 0),
                None if diff_female is None else int(diff_female == 0),
                None if diff_total_age is None else int(diff_total_age == 0),
            ))

    add_demography_for_frame(municipios, "municipio", "CD_MUN", "NM_MUN", 5, 1)
    add_demography_for_frame(distritos, "distrito", "CD_DIST", "NM_DIST", 6, 2)
    add_demography_for_frame(subdistritos, "subdistrito", "CD_SUBDIST", "NM_SUBDIST", 7, 3)
    add_demography_for_frame(bairros, "bairro", "CD_BAIRRO", "NM_BAIRRO", 7, 4)

    territory_columns = [
        "territorio_id", "chave_territorial", "nivel", "ordem_nivel", "codigo",
        "ocorrencia_codigo", "nome", "nome_normalizado", "territorio_pai_id", "codigo_uf",
        "codigo_regiao_intermediaria", "nome_regiao_intermediaria", "codigo_regiao_imediata",
        "nome_regiao_imediata", "codigo_municipio", "nome_municipio", "codigo_distrito",
        "nome_distrito", "fonte_territorio_id", "linha_origem", "codigo_duplicado",
        "dados_demograficos_disponiveis", "observacoes",
    ]
    territory_sql = f"INSERT INTO territorios({','.join(territory_columns)}) VALUES ({','.join('?' for _ in territory_columns)})"
    insert_many(conn, territory_sql, [tuple(t[c] for c in territory_columns) for t in territories])

    demog_cols = ["territorio_id", "fonte_id", "linha_origem", *VAR_COLUMNS_SQL]
    demog_sql = f"INSERT INTO demografia_agregada({','.join(demog_cols)}) VALUES ({','.join('?' for _ in demog_cols)})"
    insert_many(conn, demog_sql, demography_rows)
    insert_many(
        conn,
        """INSERT INTO qualidade_demografia(
            territorio_id,celulas_ausentes,diferenca_total_sexo,diferenca_masculino_faixas,
            diferenca_feminino_faixas,diferenca_total_faixas,fecha_total_sexo,
            fecha_masculino_faixas,fecha_feminino_faixas,fecha_total_faixas
        ) VALUES (?,?,?,?,?,?,?,?,?,?)""",
        quality_rows,
    )

    # Indicadores adicionais do projeto: idade exata municipal e renda 2010.
    indicator_rows: dict[int, list[Any]] = {}
    income_cols = ["renda_ate_2sm", "renda_2_a_5sm", "renda_5_a_10sm", "renda_mais_10sm"]
    if not master_mun.empty:
        numeric_cols = [
            "masc_16_19", "fem_16_19", "masc_20_29", "fem_20_29", "masc_30_39", "fem_30_39",
            "masc_40_49", "fem_40_49", "masc_50mais", "fem_50mais", *income_cols,
        ]
        for col in numeric_cols:
            if col in master_mun.columns:
                master_mun[col] = pd.to_numeric(master_mun[col], errors="coerce")
        for _, row in master_mun.iterrows():
            code = str(row["cod_municipio"])
            tid = key_to_id.get(f"municipio:{code}")
            if not tid:
                continue
            male16 = sum_if_complete([row.get("masc_16_19"), row.get("masc_20_29"), row.get("masc_30_39"), row.get("masc_40_49"), row.get("masc_50mais")])
            female16 = sum_if_complete([row.get("fem_16_19"), row.get("fem_20_29"), row.get("fem_30_39"), row.get("fem_40_49"), row.get("fem_50mais")])
            pop16 = male16 + female16 if male16 is not None and female16 is not None else None
            age1619 = sum_if_complete([row.get("masc_16_19"), row.get("fem_16_19")])
            incomes = [safe_int(row.get(c)) for c in income_cols]
            income_available = int(all(v is not None for v in incomes) and sum(incomes) > 0)
            indicator_rows[tid] = [
                tid, male16, female16, pop16, age1619,
                "Total 16–19 do SIDRA 9514; divisão por sexo proporcional à faixa 15–19; faixas 20+ do Censo 2022",
                *incomes, income_available, 2010, 6, 6,
                None if income_available else "Renda não localizada na base de 2010; valores brutos permanecem zerados na fonte do projeto.",
            ]

    if not master_dist.empty:
        for col in income_cols:
            if col in master_dist.columns:
                master_dist[col] = pd.to_numeric(master_dist[col], errors="coerce")
        for _, row in master_dist.iterrows():
            code = str(row["cod_distrito"])
            tid = key_to_id.get(f"distrito:{code}")
            if not tid:
                continue
            incomes = [safe_int(row.get(c)) for c in income_cols]
            income_available = int(all(v is not None for v in incomes) and sum(incomes) > 0)
            indicator_rows[tid] = [
                tid, None, None, None, None, None,
                *incomes, income_available, 2010, None, 7,
                None if income_available else "Renda não localizada na base de 2010; valores brutos permanecem zerados na fonte do projeto.",
            ]

    # Distritos novos no agregado, ausentes da base mestre do projeto.
    for code in sorted(set(distritos["CD_DIST"]) - set(master_dist.get("cod_distrito", pd.Series(dtype="string")))):
        tid = key_to_id.get(f"distrito:{code}")
        if tid and tid not in indicator_rows:
            indicator_rows[tid] = [
                tid, None, None, None, None, None,
                None, None, None, None, 0, 2010, None, None,
                "Distrito presente no agregado demográfico, mas ausente da base mestre distrital do projeto.",
            ]

    insert_many(
        conn,
        """INSERT INTO indicadores_amostragem(
            territorio_id,masculino_16_mais_municipal,feminino_16_mais_municipal,populacao_16_mais_municipal,
            idade_16_19_municipal,metodo_16_mais,renda_ate_2sm_bruta,renda_2_a_5sm_bruta,
            renda_5_a_10sm_bruta,renda_mais_10sm_bruta,renda_disponivel,ano_referencia_renda,
            fonte_idade_exata_id,fonte_renda_id,observacoes
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        [tuple(v) for _, v in sorted(indicator_rows.items())],
    )

    # Dicionário completo.
    dictionary_rows: list[tuple[Any, ...]] = []
    sigla_rows: list[tuple[Any, ...]] = []
    for sheet, frame in dictionary_frames.items():
        if "Variável" in frame.columns:
            for _, row in frame.iterrows():
                variable = str(row["Variável"]).strip()
                friendly = FRIENDLY_VARIABLES.get(variable)
                dictionary_rows.append((
                    variable, sheet,
                    None if "Tipo" not in frame.columns or pd.isna(row.get("Tipo")) else str(row.get("Tipo")),
                    None if pd.isna(row.get("Tema")) else str(row.get("Tema")),
                    str(row.get("Descrição", "")).strip(),
                    int(variable in FRIENDLY_VARIABLES),
                    friendly[0] if friendly else None,
                    friendly[1] if friendly else None,
                    friendly[2] if friendly else None,
                ))
        elif "Sigla" in frame.columns:
            for _, row in frame.iterrows():
                sigla_rows.append((str(row["Sigla"]).strip(), str(row["Descrição"]).strip(), sheet))
    insert_many(
        conn,
        """INSERT INTO dicionario_variaveis(
            variavel,colecao,tipo,tema,descricao,utilizada_neste_banco,nome_amigavel,grupo_amostragem,ordem_amostragem
        ) VALUES (?,?,?,?,?,?,?,?,?)""",
        dictionary_rows,
    )
    insert_many(conn, "INSERT INTO dicionario_siglas(sigla,descricao,colecao) VALUES (?,?,?)", sigla_rows)

    # Cobertura de filhos por pai.
    territory_df = pd.DataFrame(territories)
    demog_by_id = {row[0]: row[3] for row in demography_rows}  # v01006 é o primeiro valor após metadados
    coverage_rows: list[tuple[Any, ...]] = []

    def add_coverage(parent_level: str, child_level: str) -> None:
        parents = territory_df[territory_df["nivel"] == parent_level]
        children = territory_df[territory_df["nivel"] == child_level]
        for _, parent in parents.iterrows():
            child_rows = children[children["territorio_pai_id"] == parent["territorio_id"]]
            q_records = len(child_rows)
            q_codes = child_rows["codigo"].nunique() if q_records else 0
            child_ids = child_rows["territorio_id"].tolist()
            values = [demog_by_id.get(int(tid)) for tid in child_ids]
            values_non_null = [v for v in values if v is not None]
            q_with_data = len(values_non_null)
            parent_pop = demog_by_id.get(int(parent["territorio_id"]))
            child_pop = sum(values_non_null) if values_non_null else (0 if q_records else None)
            applicable = int(q_records > 0)
            diff = parent_pop - child_pop if parent_pop is not None and child_pop is not None else None
            ratio = child_pop / parent_pop if parent_pop not in (None, 0) and child_pop is not None else None
            exact = int(diff == 0) if applicable and diff is not None else None
            notes = None
            if not applicable:
                notes = f"Nenhum registro de {child_level} disponível para este pai."
            elif exact == 0:
                notes = "A soma dos filhos não cobre integralmente a população do pai; usar o nível com cautela."
            coverage_rows.append((
                int(parent["territorio_id"]), child_level, applicable, q_records, q_codes, q_with_data,
                parent_pop, child_pop, diff, ratio, exact, notes,
            ))

    add_coverage("municipio", "distrito")
    add_coverage("distrito", "subdistrito")
    add_coverage("municipio", "bairro")
    insert_many(
        conn,
        """INSERT INTO cobertura_hierarquia(
            territorio_pai_id,nivel_filho,aplicavel,quantidade_registros_filhos,quantidade_codigos_filhos,
            quantidade_filhos_com_dados,populacao_pai,populacao_filhos,diferenca_populacao,
            proporcao_cobertura,fecha_exatamente,observacoes
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        coverage_rows,
    )

    alerts: list[tuple[Any, ...]] = []
    alert_id = 1
    for territory in territories:
        if territory["nivel"] in {"municipio", "distrito", "subdistrito", "bairro"} and not territory["dados_demograficos_disponiveis"]:
            alerts.append((
                alert_id, "alta", "demografia_indisponivel", territory["territorio_id"], territory["nivel"], territory["codigo"],
                "A fonte usa ponto (.) para todas as variáveis demográficas; os valores foram gravados como NULL.", None,
            ))
            alert_id += 1

    duplicate_groups = territory_df[territory_df["codigo_duplicado"] == 1].groupby(["nivel", "codigo"])
    for (level, code), group in duplicate_groups:
        details = group[["chave_territorial", "nome", "linha_origem"]].to_dict("records")
        alerts.append((
            alert_id, "alta", "codigo_territorial_duplicado", None, level, code,
            "O mesmo código aparece em mais de um registro. Todos foram preservados com chaves territoriais distintas.",
            json.dumps(details, ensure_ascii=False),
        ))
        alert_id += 1

    for tid, values in indicator_rows.items():
        income_available = int(values[10])
        if income_available == 0:
            t = territory_df.loc[territory_df["territorio_id"] == tid].iloc[0]
            alerts.append((
                alert_id, "media", "renda_indisponivel", int(tid), str(t["nivel"]), str(t["codigo"]),
                "Não há distribuição de renda utilizável para esta unidade na base de 2010.", None,
            ))
            alert_id += 1

    coverage_df = pd.DataFrame(coverage_rows, columns=[
        "territorio_pai_id", "nivel_filho", "aplicavel", "q_registros", "q_codigos", "q_com_dados",
        "pop_pai", "pop_filhos", "diferenca", "proporcao", "fecha", "observacoes",
    ])
    for _, row in coverage_df[(coverage_df["aplicavel"] == 1) & (coverage_df["fecha"] == 0)].iterrows():
        t = territory_df.loc[territory_df["territorio_id"] == row["territorio_pai_id"]].iloc[0]
        alerts.append((
            alert_id, "media", f"cobertura_{row['nivel_filho']}_parcial", int(t["territorio_id"]), str(t["nivel"]), str(t["codigo"]),
            f"A soma de {row['nivel_filho']} cobre {row['proporcao']:.2%} da população do território pai.",
            json.dumps({"populacao_pai": row["pop_pai"], "populacao_filhos": row["pop_filhos"], "diferenca": row["diferenca"]}, ensure_ascii=False),
        ))
        alert_id += 1

    new_district_codes = sorted(set(distritos["CD_DIST"]) - set(master_dist.get("cod_distrito", pd.Series(dtype="string"))))
    for code in new_district_codes:
        tid = key_to_id.get(f"distrito:{code}")
        t = territory_df.loc[territory_df["territorio_id"] == tid].iloc[0] if tid else None
        alerts.append((
            alert_id, "baixa", "distrito_ausente_master_projeto", tid,
            "distrito", code,
            "Distrito presente no novo agregado demográfico, mas ausente da base mestre distrital anterior do projeto.", None,
        ))
        alert_id += 1

    insert_many(
        conn,
        "INSERT INTO alertas_qualidade(alerta_id,severidade,categoria,territorio_id,nivel,codigo,mensagem,detalhes_json) VALUES (?,?,?,?,?,?,?,?)",
        alerts,
    )

    # Métricas de auditoria.
    metrics: list[tuple[Any, ...]] = []
    for level, frame in [("municipio", municipios), ("distrito", distritos), ("subdistrito", subdistritos), ("bairro", bairros)]:
        metrics.append(("registros_demograficos", level, float(len(frame)), None, None))
        metrics.append(("registros_sem_demografia", level, float(frame[VAR_COLUMNS[0]].isna().sum()), None, "Valores de população total indisponíveis."))
    metrics.extend([
        ("variaveis_dicionario", "", float(len(dictionary_rows)), None, None),
        ("siglas_dicionario", "", float(len(sigla_rows)), None, None),
        ("populacao_total_municipios", "municipio", float(municipios["V01006"].sum()), None, None),
        ("populacao_total_distritos", "distrito", float(distritos["V01006"].sum()), None, None),
        ("codigos_duplicados", "bairro", float(len(duplicate_groups)), None, None),
        ("distritos_novos_vs_master", "distrito", float(len(new_district_codes)), None, None),
        ("municipios_com_bairros", "municipio", float(((coverage_df["nivel_filho"] == "bairro") & (coverage_df["aplicavel"] == 1)).sum()), None, None),
        ("municipios_com_cobertura_bairro_exata", "municipio", float(((coverage_df["nivel_filho"] == "bairro") & (coverage_df["fecha"] == 1)).sum()), None, None),
        ("distritos_com_subdistritos", "distrito", float(((coverage_df["nivel_filho"] == "subdistrito") & (coverage_df["aplicavel"] == 1)).sum()), None, None),
        ("distritos_com_cobertura_subdistrito_exata", "distrito", float(((coverage_df["nivel_filho"] == "subdistrito") & (coverage_df["fecha"] == 1)).sum()), None, None),
        ("alertas_total", "", float(len(alerts)), None, None),
    ])
    insert_many(conn, "INSERT INTO metricas_qualidade(metrica,nivel,valor_numerico,valor_texto,observacoes) VALUES (?,?,?,?,?)", metrics)

    conn.commit()
    integrity = conn.execute("PRAGMA integrity_check").fetchone()[0]
    foreign_key_issues = conn.execute("PRAGMA foreign_key_check").fetchall()
    if integrity != "ok" or foreign_key_issues:
        conn.close()
        raise RuntimeError(f"Falha de integridade: {integrity}; FKs: {foreign_key_issues[:5]}")

    summary = {
        "database": str(output),
        "size_bytes": output.stat().st_size,
        "integrity_check": integrity,
        "foreign_key_issues": len(foreign_key_issues),
        "territories": conn.execute("SELECT nivel, COUNT(*) FROM territorios GROUP BY nivel ORDER BY ordem_nivel, nivel").fetchall(),
        "demography_rows": conn.execute("SELECT COUNT(*) FROM demografia_agregada").fetchone()[0],
        "dictionary_variables": conn.execute("SELECT COUNT(*) FROM dicionario_variaveis").fetchone()[0],
        "alerts": conn.execute("SELECT categoria, COUNT(*) FROM alertas_qualidade GROUP BY categoria ORDER BY categoria").fetchall(),
        "municipal_population": conn.execute("SELECT SUM(populacao_total) FROM vw_amostra_municipios").fetchone()[0],
        "district_population": conn.execute("SELECT SUM(populacao_total) FROM vw_amostra_distritos").fetchone()[0],
    }

    if reports_dir is not None:
        reports_dir.mkdir(parents=True, exist_ok=True)
        pd.read_sql_query(
            "SELECT metrica,nivel,valor_numerico,valor_texto,observacoes FROM metricas_qualidade ORDER BY metrica,nivel",
            conn,
        ).to_csv(reports_dir / "metricas_qualidade.csv", index=False, encoding="utf-8-sig")
        pd.read_sql_query(
            "SELECT * FROM vw_alertas_qualidade WHERE categoria='demografia_indisponivel' ORDER BY nivel,codigo",
            conn,
        ).to_csv(reports_dir / "territorios_sem_demografia.csv", index=False, encoding="utf-8-sig")
        pd.read_sql_query(
            "SELECT * FROM vw_alertas_qualidade WHERE categoria='codigo_territorial_duplicado' ORDER BY nivel,codigo",
            conn,
        ).to_csv(reports_dir / "codigos_duplicados.csv", index=False, encoding="utf-8-sig")
        pd.read_sql_query(
            """SELECT t.codigo,t.nome,t.uf,t.nome_uf,c.*
               FROM cobertura_hierarquia c JOIN vw_territorios t ON t.territorio_id=c.territorio_pai_id
               WHERE c.nivel_filho='bairro' AND c.aplicavel=1 ORDER BY c.proporcao_cobertura,t.codigo""",
            conn,
        ).to_csv(reports_dir / "cobertura_bairros_por_municipio.csv", index=False, encoding="utf-8-sig")
        pd.read_sql_query(
            """SELECT t.codigo,t.nome,t.uf,t.nome_uf,t.codigo_municipio,c.*
               FROM cobertura_hierarquia c JOIN vw_territorios t ON t.territorio_id=c.territorio_pai_id
               WHERE c.nivel_filho='subdistrito' AND c.aplicavel=1 ORDER BY c.proporcao_cobertura,t.codigo""",
            conn,
        ).to_csv(reports_dir / "cobertura_subdistritos_por_distrito.csv", index=False, encoding="utf-8-sig")
        pd.read_sql_query(
            "SELECT * FROM vw_alertas_qualidade WHERE categoria='renda_indisponivel' ORDER BY nivel,codigo",
            conn,
        ).to_csv(reports_dir / "renda_indisponivel.csv", index=False, encoding="utf-8-sig")
        pd.read_sql_query(
            "SELECT * FROM vw_alertas_qualidade WHERE categoria='distrito_ausente_master_projeto' ORDER BY codigo",
            conn,
        ).to_csv(reports_dir / "distritos_novos_vs_master_projeto.csv", index=False, encoding="utf-8-sig")
        (reports_dir / "resumo_validacao.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    conn.execute("VACUUM")
    conn.close()
    summary["size_bytes"] = output.stat().st_size
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Constrói o banco de dados de apoio à montagem de amostras.")
    parser.add_argument("--municipios", type=Path, required=True)
    parser.add_argument("--distritos", type=Path, required=True)
    parser.add_argument("--subdistritos", type=Path, required=True)
    parser.add_argument("--bairros", type=Path, required=True)
    parser.add_argument("--dicionario", type=Path, required=True)
    parser.add_argument("--master-municipios", type=Path)
    parser.add_argument("--master-distritos", type=Path)
    parser.add_argument("--saida", type=Path, required=True)
    parser.add_argument("--relatorios-dir", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    inputs = InputPaths(
        municipios=args.municipios,
        distritos=args.distritos,
        subdistritos=args.subdistritos,
        bairros=args.bairros,
        dicionario=args.dicionario,
        master_municipios=args.master_municipios,
        master_distritos=args.master_distritos,
    )
    summary = build_database(inputs, args.saida, args.relatorios_dir)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
