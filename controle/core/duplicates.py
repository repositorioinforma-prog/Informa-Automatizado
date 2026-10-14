from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

from controle.core.normalization import normalize_text

# Padrões usados para achar as variáveis automaticamente pelo LABEL (ou, na
# falta de label útil, pelo nome da variável) — cobre os rótulos exatos que
# você descreveu ("Nome", "Tel.") e algumas variações comuns.
NAME_LABEL_PATTERNS = ["nome", "nome completo", "nome do entrevistado"]
PHONE_LABEL_PATTERNS = ["tel", "tel.", "telefone", "celular", "whatsapp"]
DATE_LABEL_PATTERNS = [
    "data de criacao", "data criacao", "data da entrevista", "data entrevista",
    "date created", "data cadastro", "data",
]


def find_column_by_label(
    column_labels: dict[str, str],
    patterns: list[str],
    also_check_names: bool = False,
) -> str | None:
    """Acha a variável cujo LABEL bate com um dos padrões (ex.: label 'Nome',
    'Tel.'). Prioriza igualdade exata do label normalizado; se ninguém bater
    exato, aceita o label CONTER um dos padrões; só then (e se
    `also_check_names`) cai para olhar o nome da própria variável — útil para
    coisas como data de criação, que costumam não ter um label descritivo."""
    normalized_patterns = [normalize_text(p) for p in patterns]
    exact_hits, contains_hits, name_hits = [], [], []
    for col, label in column_labels.items():
        norm_label = normalize_text(label)
        if norm_label:
            if norm_label in normalized_patterns:
                exact_hits.append(col)
            elif any(p in norm_label for p in normalized_patterns):
                contains_hits.append(col)
        if also_check_names:
            norm_name = normalize_text(col)
            if norm_name and any(p in norm_name for p in normalized_patterns):
                name_hits.append(col)
    if exact_hits:
        return exact_hits[0]
    if contains_hits:
        return contains_hits[0]
    if name_hits:
        return name_hits[0]
    return None


def normalize_phone(value) -> str:
    """Mantém só os dígitos (remove parênteses, traço, espaço, +55, etc.).

    [Informa] Telefone gravado como NÚMERO no .sav (11987654321.0) não ganha
    mais um "0" no fim, e telefones de preenchimento (0, 00000000, 99999999,
    menos de 8 dígitos) contam como vazios — antes viravam duplicidade "certa"
    entre entrevistas sem relação nenhuma."""
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        if value.is_integer():
            value = int(value)
    digits = re.sub(r"\D", "", str(value))
    if len(digits) < 8 or len(set(digits)) == 1:
        return ""
    return digits


def normalize_name(value) -> str:
    return normalize_text(value)


@dataclass
class DuplicateConfig:
    id_column: str
    name_column: str | None
    phone_column: str | None
    date_column: str | None = None
    phone_suffix_len: int = 8  # nº de dígitos finais usados para "telefone parecido"


def find_duplicates(df: pd.DataFrame, config: DuplicateConfig) -> pd.DataFrame:
    """
    Detecta duplicidade de entrevistas por nome + telefone, em 3 níveis de
    confiança (mesma lógica da seção "Duplicidades" do projeto):

    - certa: mesmo telefone (normalizado, só dígitos)
    - provavel: mesmo nome + telefone com os últimos dígitos iguais
      (ex.: DDD ausente/diferente, mas o resto do número bate)
    - possivel: mesmo nome, telefones diferentes (ou um deles vazio)

    Dentro de cada grupo, recomenda manter a entrevista mais antiga (por
    `date_column`, quando disponível — senão pela ordem em que aparece no
    banco) e excluir as demais. Retorna uma linha por entrevista envolvida em
    algum grupo (entrevistas sem duplicidade não aparecem no resultado).
    """
    working = pd.DataFrame({"ID": df[config.id_column].astype(str)}, index=df.index)
    working["nome_original"] = df[config.name_column] if config.name_column else ""
    working["telefone_original"] = df[config.phone_column] if config.phone_column else ""
    working["nome_norm"] = working["nome_original"].map(normalize_name)
    working["telefone_norm"] = working["telefone_original"].map(normalize_phone)
    if config.date_column and config.date_column in df.columns:
        working["data"] = pd.to_datetime(df[config.date_column], errors="coerce", format="mixed")
    else:
        working["data"] = pd.NaT
    working["_ordem_original"] = range(len(working))

    assigned = pd.Series(False, index=working.index)
    groups: list[dict] = []
    group_counter = 0

    def _add_group(idx, tipo: str):
        nonlocal group_counter
        idx = [i for i in idx if not assigned.loc[i]]
        if len(idx) < 2:
            return
        group_counter += 1
        sub = working.loc[idx].sort_values(by=["data", "_ordem_original"], na_position="last")
        for pos, (_, row) in enumerate(sub.iterrows()):
            recomendacao = "manter" if pos == 0 else "excluir"
            if tipo == "certa":
                motivo = f"Mesmo telefone ({row['telefone_norm']})."
            elif tipo == "provavel":
                motivo = f"Mesmo nome ({row['nome_original']}) e telefone terminando igual ({row['telefone_original']})."
            else:
                motivo = f"Mesmo nome ({row['nome_original']}), telefones diferentes ou ausentes."
            if pos > 0:
                data_txt = "" if pd.isna(sub.iloc[0]["data"]) else f" ({sub.iloc[0]['data']:%d/%m/%Y %H:%M})"
                motivo += f" Mantida a entrevista mais antiga do grupo, ID {sub.iloc[0]['ID']}{data_txt}."
            groups.append({
                "grupo": group_counter, "ID": row["ID"], "nome": row["nome_original"],
                "telefone": row["telefone_original"], "data": row["data"],
                "tipo_duplicidade": tipo, "recomendacao": recomendacao, "motivo": motivo,
            })
        assigned.loc[idx] = True

    if config.phone_column:
        com_telefone = working[working["telefone_norm"] != ""]
        for _, sub in com_telefone.groupby("telefone_norm"):
            if len(sub) > 1:
                _add_group(list(sub.index), "certa")

    if config.name_column and config.phone_column:
        restante = working[~assigned & (working["nome_norm"] != "") & (working["telefone_norm"] != "")]
        sufixo = restante["telefone_norm"].str[-config.phone_suffix_len:]
        for (_, suf), sub in restante.groupby([restante["nome_norm"], sufixo]):
            if len(sub) > 1 and suf:
                _add_group(list(sub.index), "provavel")

    if config.name_column:
        restante = working[~assigned & (working["nome_norm"] != "")]
        for _, sub in restante.groupby("nome_norm"):
            if len(sub) > 1:
                _add_group(list(sub.index), "possivel")

    columns = ["grupo", "ID", "nome", "telefone", "data", "tipo_duplicidade", "recomendacao", "motivo"]
    if not groups:
        return pd.DataFrame(columns=columns)
    result = pd.DataFrame(groups)[columns]
    # [Informa] dentro de cada grupo, a entrevista mantida vem primeiro
    ordem = result["recomendacao"].map({"manter": 0, "excluir": 1})
    return result.assign(_o=ordem).sort_values(["grupo", "_o"], kind="stable").drop(columns="_o").reset_index(drop=True)
