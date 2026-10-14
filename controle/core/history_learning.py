"""
Aprendizado a partir de um banco já codificado anteriormente (ex.: a onda de
ontem da mesma pesquisa de acompanhamento contínuo).

A ideia: se alguém já escreveu "Lebron" num banco anterior e um humano (ou o
próprio motor, com confiança) resolveu isso como código X, um novo respondente
que escrever exatamente "Lebron" hoje muito provavelmente quer dizer a mesma
coisa. Isso resolve exatamente a categoria de caso mais difícil — apelidos,
erros de digitação recorrentes, nomes populares — que a Base Brasil e o fuzzy
sozinhos não cobrem bem, porque usa a decisão real já tomada no histórico do
próprio projeto, em vez de tentar adivinhar uma regra genérica.

Cuidados deliberados:
- Só aprende de respostas que JÁ tinham código preenchido no banco anterior
  (não aprende de casos que também ficaram em branco/pendentes).
- Se o mesmo texto aparece com códigos DIFERENTES no banco anterior (ou seja,
  o próprio histórico é inconsistente para aquele texto), a entrada é
  descartada — não escolhe um dos dois sozinho.
- Só é usado para preencher casos que o motor atual não resolveu com
  confiança sozinho (nunca substitui uma decisão já confirmada hoje).
"""
from __future__ import annotations

import pandas as pd

from .normalization import normalize_text
from .recoding import choose_source_text


def build_text_to_code_lookup(
    old_df: pd.DataFrame,
    source_columns: tuple[str, ...],
    code_column: str,
    bank_value_labels_old: dict,
    hierarchy=None,
    project_labels: dict | None = None,
) -> tuple[dict[str, object], dict[str, list[object]]]:
    """Constrói normalize_text(texto respondido) -> código já usado no banco anterior.

    Retorna (lookup, conflitos). 'conflitos' lista, por texto normalizado, quais
    códigos diferentes apareceram — útil para mostrar transparência ao usuário
    sobre o que foi descartado por ambiguidade histórica.

    Se 'hierarchy' e 'project_labels' forem passados (ver
    geography.hierarquia_bairros), qualquer entrada aprendida em que o texto é
    um bairro/loteamento real que pertence a um bairro oficial DIFERENTE do
    bairro oficial do código aprendido é descartada — isso existe
    especificamente para não herdar de bancos antigos uma política de
    "realocar para o bairro aprovado mais próximo/mesma região", que é uma
    decisão de negócio que não deve ser reproduzida silenciosamente.
    """
    if code_column not in old_df.columns:
        return {}, {}

    label_by_code = {}
    if hierarchy is not None and project_labels:
        label_by_code = {code: label for code, label in project_labels.items()}

    seen: dict[str, set] = {}
    for _, row in old_df.iterrows():
        code = row.get(code_column)
        if pd.isna(code):
            continue
        text, _, _ = choose_source_text(row, source_columns, bank_value_labels_old)
        if not text:
            continue
        key = normalize_text(text)

        if hierarchy is not None and label_by_code:
            found = hierarchy.find(text)
            if found is not None:
                approved_label = label_by_code.get(code)
                belongs = approved_label is not None and normalize_text(approved_label) in (
                    {normalize_text(found.bairro_oficial)}
                    | {normalize_text(p) for p in hierarchy.data.get(found.bairro_oficial, {}).get("parcelamentos", [])}
                )
                if not belongs:
                    # texto é um bairro/loteamento real, mas o código aprendido
                    # aponta para outro bairro oficial — política antiga de
                    # realocação por proximidade, descarta.
                    continue

        seen.setdefault(key, set()).add(code)

    lookup = {}
    conflicts = {}
    for key, codes in seen.items():
        if len(codes) == 1:
            lookup[key] = next(iter(codes))
        else:
            conflicts[key] = sorted(codes, key=str)
    return lookup, conflicts


UNRESOLVED_STATUSES = {"SUGESTÃO", "AMBÍGUO", "FORA DA AMOSTRA", "NÃO IDENTIFICADO"}


def apply_learned_lookup(
    audit_df: pd.DataFrame,
    lookup: dict[str, object],
    project_labels: dict,
) -> pd.DataFrame:
    """Preenche, no relatório de auditoria já gerado, os casos que o motor
    normal não resolveu com confiança, mas que batem exatamente (texto
    normalizado) com algo já codificado no banco anterior. Não altera nenhuma
    linha que já estava CONFIRMADO/AJUSTADO/ACEITO — só complementa o que
    sobrou pendente."""
    if audit_df.empty or not lookup:
        return audit_df

    result = audit_df.copy()
    label_by_code = {code: label for code, label in project_labels.items()}

    # Força as colunas que vamos alterar linha a linha para dtype 'object'.
    # Sem isso, se a coluna ficou com um dtype numérico estrito (ex.:
    # float64) e o código aprendido do banco anterior é de um tipo um pouco
    # diferente (ex.: numpy.int64 vs float, ou um código de VALUE LABELS em
    # string), o pandas mais recente recusa a atribuição com TypeError em vez
    # de converter silenciosamente.
    for col in ("codigo_sugerido", "label_sugerido", "metodo", "status", "decisao_automatica"):
        if col in result.columns:
            result[col] = result[col].astype(object)

    mask = result["status"].isin(UNRESOLVED_STATUSES)
    for idx in result.index[mask]:
        key = normalize_text(result.at[idx, "texto_interpretado"])
        if not key or key not in lookup:
            continue
        code = lookup[key]
        result.at[idx, "codigo_sugerido"] = code
        result.at[idx, "label_sugerido"] = label_by_code.get(code, "")
        result.at[idx, "metodo"] = "aprendido_banco_anterior"
        result.at[idx, "status"] = "APRENDIDO (banco anterior)"
        result.at[idx, "decisao_automatica"] = True
    return result
