from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable
import re

import pandas as pd

from controle.core.normalization import normalize_text


_COLUMN_ALIASES = {
    "uf": {"uf", "sg uf", "sigla uf", "estado", "sigla estado", "nm uf", "nome uf"},
    "municipio": {
        "municipio", "município", "municipios", "municípios", "cidade", "cidades", "nome municipio", "nome município",
        "nm mun", "nm municipio", "nm município",
    },
    "municipio_codigo": {
        "codigo municipio", "código município", "codigo do municipio", "código do município",
        "codigo ibge", "código ibge", "codigo ibge municipio", "código ibge município",
        "cd mun", "cod municipio", "cod município",
    },
    "distrito": {"distrito", "distritos", "nome distrito", "nm distrito", "nm dist"},
    "distrito_codigo": {"codigo distrito", "código distrito", "cd distrito", "cod distrito", "cd dist"},
    "bairro": {"bairro", "bairros", "nome bairro", "nm bairro", "localidade", "localidades", "bairro localidade"},
    "tipo": {"tipo", "tipo localidade", "nivel", "nível", "categoria"},
}


def _header_key(value: object) -> str:
    text = normalize_text(value)
    return re.sub(r"\s+", " ", text).strip()


_ALIAS_LOOKUP = {
    _header_key(alias): canonical
    for canonical, aliases in _COLUMN_ALIASES.items()
    for alias in aliases
}


@dataclass(frozen=True)
class GeographicRecord:
    entity_type: str
    name: str
    uf: str = ""
    municipality: str = ""
    district: str = ""
    municipality_code: str = ""
    district_code: str = ""
    source_sheet: str = ""

    @property
    def normalized_name(self) -> str:
        return normalize_text(self.name)


@dataclass
class GeographyDatabase:
    records: list[GeographicRecord]
    source_name: str = ""

    def __post_init__(self):
        self.by_name: dict[str, list[GeographicRecord]] = {}
        self.by_type: dict[str, list[GeographicRecord]] = {"municipio": [], "distrito": [], "bairro": []}
        for record in self.records:
            key = record.normalized_name
            if not key:
                continue
            self.by_name.setdefault(key, []).append(record)
            self.by_type.setdefault(record.entity_type, []).append(record)

    def summary(self) -> dict[str, int]:
        return {
            "registros": len(self.records),
            "municipios": len({(normalize_text(r.uf), normalize_text(r.name)) for r in self.by_type.get("municipio", [])}),
            "distritos": len(self.by_type.get("distrito", [])),
            "bairros": len(self.by_type.get("bairro", [])),
            "ufs": len({_record_uf_sigla(r.uf) or normalize_text(r.uf) for r in self.records if r.uf}),
        }

    def dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([
            {
                "tipo": r.entity_type,
                "nome": r.name,
                "municipio": r.municipality,
                "distrito": r.district,
                "uf": r.uf,
                "codigo_municipio": r.municipality_code,
                "codigo_distrito": r.district_code,
                "aba_origem": r.source_sheet,
            }
            for r in self.records
        ])

    def filter_by_uf(self, siglas: "Iterable[str] | None") -> "GeographyDatabase":
        """Restringe a base a um subconjunto de UFs (ex.: só RJ, ou RJ+SP).
        Registros sem UF preenchida são mantidos (não há como saber se pertencem
        ou não à abrangência, e é mais seguro não descartá-los). Passar None ou
        um conjunto vazio devolve a base inteira (Brasil todo), sem filtrar."""
        siglas_set = {s.upper() for s in siglas} if siglas else set()
        if not siglas_set:
            return self
        filtered = []
        for r in self.records:
            sigla = _record_uf_sigla(r.uf) if r.uf else None
            if sigla is None or sigla in siglas_set:
                filtered.append(r)
        return GeographyDatabase(filtered, source_name=self.source_name)


def _canonical_columns(columns: Iterable[object]) -> dict[object, str]:
    mapped = {}
    for col in columns:
        canonical = _ALIAS_LOOKUP.get(_header_key(col))
        if canonical:
            mapped[col] = canonical
    return mapped


_UF_SIGLA_NOME = {
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas", "BA": "Bahia",
    "CE": "Ceará", "DF": "Distrito Federal", "ES": "Espírito Santo", "GO": "Goiás",
    "MA": "Maranhão", "MT": "Mato Grosso", "MS": "Mato Grosso do Sul", "MG": "Minas Gerais",
    "PA": "Pará", "PB": "Paraíba", "PR": "Paraná", "PE": "Pernambuco", "PI": "Piauí",
    "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte", "RS": "Rio Grande do Sul",
    "RO": "Rondônia", "RR": "Roraima", "SC": "Santa Catarina", "SP": "São Paulo",
    "SE": "Sergipe", "TO": "Tocantins",
}
UF_OPTIONS = [f"{sigla} — {nome}" for sigla, nome in _UF_SIGLA_NOME.items()]
_UF_KEY_TO_SIGLA = {}
for _sigla, _nome in _UF_SIGLA_NOME.items():
    _UF_KEY_TO_SIGLA[normalize_text(_sigla)] = _sigla
    _UF_KEY_TO_SIGLA[normalize_text(_nome)] = _sigla


def uf_option_to_sigla(option: str) -> str:
    return option.split(" — ")[0].strip()


def _record_uf_sigla(record_uf: str) -> str | None:
    return _UF_KEY_TO_SIGLA.get(normalize_text(record_uf))


_NAME_UF_PATTERN = re.compile(r"^(.*\S)\s*\(([A-Za-zÀ-ÿ]{2})\)\s*$")


def _split_name_uf(name: str, uf: str) -> tuple[str, str]:
    """Algumas planilhas trazem o nome já com a UF embutida, ex.: 'Cordeiro (RJ)'.
    Se ainda não temos UF vinda de outra coluna, separamos o nome do sufixo —
    caso contrário, o nome inteiro ('cordeiro rj') nunca bate com respostas que
    dizem só 'Cordeiro', e pode até colidir por engano com um bairro homônimo
    de outro estado."""
    if uf:
        return name, uf
    match = _NAME_UF_PATTERN.match(name) if name else None
    if match:
        return match.group(1).strip(), match.group(2).strip().upper()
    return name, uf


def _clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def _sheet_hint(sheet_name: str) -> str | None:
    n = normalize_text(sheet_name)
    if "bairro" in n or "localidade" in n:
        return "bairro"
    if "distrito" in n:
        return "distrito"
    if "cidade" in n or "municipio" in n:
        return "municipio"
    return None


def _type_from_value(value: object) -> str | None:
    n = normalize_text(value)
    if "bairro" in n or "localidade" in n:
        return "bairro"
    if "distrito" in n:
        return "distrito"
    if "municipio" in n or "cidade" in n:
        return "municipio"
    return None


def _records_from_frame(frame: pd.DataFrame, sheet_name: str) -> list[GeographicRecord]:
    hint = _sheet_hint(sheet_name)
    rename = _canonical_columns(frame.columns)
    # Separate-sheet layouts often use a generic column such as "Nome".
    if hint and hint not in rename.values():
        for col in frame.columns:
            if _header_key(col) in {"nome", "descricao", "denominacao"}:
                rename[col] = hint
                break
    if not rename:
        return []
    data = frame.rename(columns=rename)
    records: list[GeographicRecord] = []

    for _, row in data.iterrows():
        uf = _clean(row.get("uf"))
        municipality = _clean(row.get("municipio"))
        municipality, uf = _split_name_uf(municipality, uf)
        municipality_code = _clean(row.get("municipio_codigo"))
        district = _clean(row.get("distrito"))
        district_code = _clean(row.get("distrito_codigo"))
        bairro = _clean(row.get("bairro"))
        explicit_type = _type_from_value(row.get("tipo"))

        # Wide layout: one row can describe municipality, district and bairro.
        if municipality:
            records.append(GeographicRecord(
                "municipio", municipality, uf, municipality, "", municipality_code, "", sheet_name
            ))
        if district:
            records.append(GeographicRecord(
                "distrito", district, uf, municipality, district, municipality_code, district_code, sheet_name
            ))
        if bairro:
            records.append(GeographicRecord(
                "bairro", bairro, uf, municipality, district, municipality_code, district_code, sheet_name
            ))

        # Vertical layout: a generic name column may be called municipio/bairro/distrito
        # according to the sheet. The cases above already handle the normal layouts.
        if explicit_type and explicit_type == hint:
            continue

    return records


def _deduplicate(records: list[GeographicRecord]) -> list[GeographicRecord]:
    seen = set()
    out = []
    for r in records:
        key = (
            r.entity_type,
            normalize_text(r.name), normalize_text(r.uf), normalize_text(r.municipality), normalize_text(r.district),
            r.municipality_code, r.district_code,
        )
        if not r.name or key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def load_geography_excel(source: bytes | str | Path, source_name: str = "") -> GeographyDatabase:
    """Load an XLSX/XLS geographic dictionary, accepting multiple sheets and common column names."""
    excel_source = BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    sheets = pd.read_excel(excel_source, sheet_name=None, dtype=object)
    records: list[GeographicRecord] = []
    for sheet_name, frame in sheets.items():
        records.extend(_records_from_frame(frame, str(sheet_name)))
    records = _deduplicate(records)
    if not records:
        raise ValueError(
            "Não foi possível identificar colunas geográficas. Esperava colunas equivalentes a UF, Município/Cidade, Distrito e/ou Bairro."
        )
    return GeographyDatabase(records, source_name=source_name)
