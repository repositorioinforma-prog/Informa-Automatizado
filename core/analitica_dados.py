"""
Leitura de bancos (CSV / XLSX / XLS / SAV) para as ferramentas analíticas
portadas do app desktop "Relatória Analítica" (Cruzamentos, Análise
Residual, Matriz de Correlação).

Porte fiel da leitura feita em cada `pickData()` do original — inclusive
a tentativa em várias codificações quando o .sav vem com acentuação
quebrada. A única diferença é a origem do arquivo: lá era um caminho
escolhido num QFileDialog; aqui são os bytes de um st.file_uploader.

Não depende de Streamlit (testável isoladamente).
"""
from __future__ import annotations

import io
import os
import re
import tempfile
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Dict, List

import pandas as pd

_ERROS_DE_CODIFICACAO = (
    "Unable to convert string",
    "invalid byte sequence",
    "codec can't decode",
    "invalid continuation byte",
)


@dataclass
class BancoCarregado:
    df: pd.DataFrame
    nome_arquivo: str
    var_labels: Dict[str, str] = field(default_factory=dict)          # variável -> rótulo
    value_labels: Dict[str, Dict[Any, str]] = field(default_factory=dict)  # variável -> {código: rótulo}
    missing_ranges: Dict[str, List[Any]] = field(default_factory=dict)  # missing de usuário do SPSS


def canon(s: str) -> str:
    """Normalização usada nas comparações de rótulos (sem acento, minúsculo,
    só letras/números separados por espaço) — igual ao `_canon` original."""
    s = (s or "").strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def eh_nome_de_peso(nome: str) -> bool:
    return "peso" in canon(str(nome))


def _eh_erro_de_codificacao(e: Exception) -> bool:
    msg = str(e)
    return any(trecho in msg for trecho in _ERROS_DE_CODIFICACAO)


def _ler_sav(caminho: str, user_missing: bool):
    from core.pyreadstat_seguro import pyreadstat

    kwargs = {"user_missing": True} if user_missing else {}
    df = meta = None
    last_err = None

    try:
        df, meta = pyreadstat.read_sav(caminho, **kwargs)
    except Exception as e:  # noqa: BLE001 — repete a lógica do original
        last_err = e
        if not _eh_erro_de_codificacao(e):
            raise

    if df is None or meta is None:
        file_enc = None
        try:
            _, m0 = pyreadstat.read_sav(caminho, metadataonly=True)
            file_enc = (getattr(m0, "file_encoding", None) or "").strip() or None
        except Exception:
            file_enc = None

        candidatas: List[str] = []
        for enc in [file_enc, "utf-8", "cp1252", "latin1", "iso-8859-1", "iso-8859-15", "utf-16le", "ucs-2le"]:
            if enc and enc not in candidatas:
                candidatas.append(enc)

        for enc in candidatas:
            try:
                df, meta = pyreadstat.read_sav(caminho, encoding=enc, **kwargs)
                break
            except Exception as e:  # noqa: BLE001
                last_err = e
                if _eh_erro_de_codificacao(e):
                    continue
                raise

    if df is None or meta is None:
        raise last_err or Exception("Falha ao ler o arquivo .sav.")
    return df, meta


def _labels_de_variaveis(meta) -> Dict[str, str]:
    try:
        m = getattr(meta, "column_names_to_labels", None)
        if isinstance(m, dict):
            return {str(k): str(v) for k, v in m.items() if v}
        cn = getattr(meta, "column_names", None)
        cl = getattr(meta, "column_labels", None)
        if cn and cl and len(cn) == len(cl):
            return {str(k): str(v) for k, v in zip(cn, cl) if v}
    except Exception:
        pass
    return {}


def carregar_banco(
    conteudo: bytes,
    nome_arquivo: str,
    *,
    user_missing: bool = False,
    remover_colunas_vazias: bool = False,
) -> BancoCarregado:
    """Lê o banco a partir dos bytes do upload.

    user_missing: lê o .sav preservando os códigos de missing de usuário
        (e guarda `missing_ranges`) — usado pela Análise Residual, que
        trata esses códigos como ausentes por conta própria.
    remover_colunas_vazias: descarta colunas 100% vazias (Análise Residual
        e Matriz de Correlação faziam isso; Cruzamentos não).
    """
    nome = os.path.basename(nome_arquivo or "")
    ext = os.path.splitext(nome)[1].lower()

    var_labels: Dict[str, str] = {}
    value_labels: Dict[str, Dict[Any, str]] = {}
    missing_ranges: Dict[str, List[Any]] = {}

    if ext == ".csv":
        try:
            df = pd.read_csv(io.BytesIO(conteudo), encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(conteudo), encoding="latin1")
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(io.BytesIO(conteudo))
    elif ext == ".sav":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(conteudo)
            caminho = tmp.name
        try:
            df, meta = _ler_sav(caminho, user_missing=user_missing)
        finally:
            try:
                os.unlink(caminho)
            except OSError:
                pass

        var_labels = _labels_de_variaveis(meta)
        try:
            for var, m in (meta.variable_value_labels or {}).items():
                value_labels[str(var)] = dict(m)
        except Exception:
            value_labels = {}
        if user_missing:
            try:
                bruto = getattr(meta, "missing_ranges", None) or {}
                missing_ranges = {str(var): list(specs) for var, specs in bruto.items()}
            except Exception:
                missing_ranges = {}
    else:
        raise ValueError("Formato não suportado. Use CSV, XLSX, XLS ou SAV.")

    if remover_colunas_vazias:
        df = df.dropna(axis=1, how="all")

    return BancoCarregado(
        df=df,
        nome_arquivo=nome,
        var_labels=var_labels,
        value_labels=value_labels,
        missing_ranges=missing_ranges,
    )
