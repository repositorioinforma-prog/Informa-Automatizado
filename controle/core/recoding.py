from dataclasses import dataclass
import pandas as pd
from .matching import confidence_status, match_text
from .normalization import is_generic_other_label, normalize_text
from controle.geography.resolver import resolve_to_target
from controle.geography.hierarquia_bairros import get_hierarchy, find_in_any_hierarchy


@dataclass(frozen=True)
class ControlVariableConfig:
    output_name: str
    label_set_name: str
    source_columns: tuple[str, ...]
    geographic_type: str | None = None
    fuzzy_cutoff: float = 0.82
    auto_fuzzy_threshold: float = 0.93


def value_to_text(value, column: str, bank_value_labels: dict[str, dict]) -> str:
    if pd.isna(value):
        return ""
    labels = bank_value_labels.get(column, {})
    label = labels.get(value)
    if label is None:
        try:
            number = float(value)
            label = labels.get(int(number) if number.is_integer() else number)
        except (TypeError, ValueError):
            pass
    if label is not None:
        return "" if is_generic_other_label(label) else str(label).strip()
    return str(value).strip()


def source_attempts(row, source_columns, bank_value_labels):
    attempts = []
    for column in source_columns:
        raw = row.get(column)
        text = value_to_text(raw, column, bank_value_labels)
        attempts.append((column, raw, text))
    return attempts


def choose_source_text(row, source_columns, bank_value_labels):
    attempts = source_attempts(row, source_columns, bank_value_labels)
    for column, raw, text in attempts:
        if text:
            return text, column, attempts
    return "", None, attempts


def _recode_geographic_row(row, config, project_labels, bank_value_labels, id_column, geography_db):
    attempts = source_attempts(row, config.source_columns, bank_value_labels)
    geo_attempts = []
    usable = []
    ambiguous = []  # resolução nacionalmente ambígua (ex.: bairro homônimo em vários municípios)
    for column, raw, text in attempts:
        if not text:
            continue
        resolution = resolve_to_target(text, config.geographic_type, geography_db, fuzzy_cutoff=max(config.fuzzy_cutoff, 0.75))
        geo_attempts.append((column, text, resolution))
        if resolution.value:
            usable.append((column, text, resolution))
        elif resolution.status == "AMBÍGUO" and resolution.candidates:
            ambiguous.append((column, text, resolution))

    if not usable:
        fontes_consultadas = " | ".join(f"{c}={raw!s} -> {txt}" for c, raw, txt in attempts)

        # Antes de qualquer outra tentativa: se é uma variável de bairro e a
        # Base Brasil nem sequer encontrou o texto como bairro (comum para
        # loteamentos pequenos/populares que a base nacional não tem), ainda
        # assim pode ser um loteamento conhecido de uma hierarquia municipal
        # cadastrada (ex.: Campo Grande) — checa isso antes de desistir.
        if config.geographic_type == "bairro":
            for column, raw, text in attempts:
                if not text:
                    continue
                found, hierarchy = find_in_any_hierarchy(text, geography_db)
                if not found:
                    continue
                representative = hierarchy.approved_representative(found.bairro_oficial, project_labels)
                if representative:
                    rep_code, rep_label = representative
                    return {
                        "ID": row.get(id_column), "variavel_controle": config.output_name,
                        "tipo_geografico": config.geographic_type, "texto_interpretado": text,
                        "fonte_utilizada": column, "codigo_sugerido": rep_code, "label_sugerido": rep_label,
                        "metodo": f"hierarquia_bairro_oficial:{found.metodo}", "confianca": 0.9,
                        "status": "AJUSTADO AUTOMATICAMENTE", "decisao_automatica": True,
                        "candidatos": "", "localidade_base": text, "tipo_localidade_base": "loteamento",
                        "municipio_base": "", "uf_base": "",
                        "fontes_consultadas": fontes_consultadas,
                    }

        # A Base Brasil pode não conter o texto exatamente como respondido (ex.:
        # uma resposta de lista fechada como "CAPITAL/Rio de Janeiro"). Antes de
        # desistir, tenta casar diretamente com os VALUE LABELS do projeto —
        # essa é justamente a lista de localidades válidas na amostra. Verifica
        # TODAS as fontes (não só a primeira) para não aplicar automaticamente
        # quando duas fontes da própria entrevista se contradizem.
        direct_matches = []
        for column, raw, text in attempts:
            if not text:
                continue
            direct = match_text(text, project_labels, config.fuzzy_cutoff)
            if direct.code is not None:
                direct_matches.append((column, text, direct))

        if direct_matches:
            distinct_codes = {d.code for _, _, d in direct_matches}
            if len(distinct_codes) == 1:
                column, text, direct = direct_matches[0]
                status = confidence_status(direct, config.auto_fuzzy_threshold)
                automatic = status in {"CONFIRMADO", "AJUSTADO AUTOMATICAMENTE"}
                return {
                    "ID": row.get(id_column), "variavel_controle": config.output_name,
                    "tipo_geografico": config.geographic_type, "texto_interpretado": text,
                    "fonte_utilizada": column, "codigo_sugerido": direct.code, "label_sugerido": direct.label or "",
                    "metodo": f"lista_projeto_direta:{direct.method}", "confianca": round(direct.score, 4),
                    "status": status, "decisao_automatica": automatic,
                    "candidatos": ", ".join(map(str, direct.candidates)),
                    "localidade_base": "", "tipo_localidade_base": "", "municipio_base": "", "uf_base": "",
                    "fontes_consultadas": fontes_consultadas,
                }
            # Fontes diferentes da mesma entrevista apontam para localidades do
            # projeto diferentes (ex.: q0002="CAPITAL/Rio de Janeiro" mas
            # q0003 menciona claramente outro município). Não decide sozinho.
            resumo = "; ".join(f"{c}={t!r}→{d.label}" for c, t, d in direct_matches)
            return {
                "ID": row.get(id_column), "variavel_controle": config.output_name,
                "tipo_geografico": config.geographic_type, "texto_interpretado": resumo,
                "fonte_utilizada": " + ".join(dict.fromkeys(c for c, _, _ in direct_matches)),
                "codigo_sugerido": None, "label_sugerido": "",
                "metodo": "fontes_lista_projeto_conflitantes", "confianca": 0.0,
                "status": "AMBÍGUO", "decisao_automatica": False,
                "candidatos": ", ".join(sorted({str(d.label) for _, _, d in direct_matches})),
                "localidade_base": "", "tipo_localidade_base": "", "municipio_base": "", "uf_base": "",
                "fontes_consultadas": fontes_consultadas,
            }

        if ambiguous:
            column, text, resolution = ambiguous[0]
            return {
                "ID": row.get(id_column), "variavel_controle": config.output_name,
                "tipo_geografico": config.geographic_type, "texto_interpretado": text,
                "fonte_utilizada": column, "codigo_sugerido": None, "label_sugerido": "",
                "metodo": "geografica_ambigua_nacional", "confianca": resolution.score,
                "status": "AMBÍGUO", "decisao_automatica": False,
                "candidatos": ", ".join(resolution.candidates),
                "localidade_base": "", "tipo_localidade_base": "", "municipio_base": "", "uf_base": "",
                "fontes_consultadas": fontes_consultadas,
            }

        text = next((t for _, _, t in attempts if t), "")
        source = next((c for c, _, t in attempts if t), "")
        return {
            "ID": row.get(id_column), "variavel_controle": config.output_name,
            "tipo_geografico": config.geographic_type, "texto_interpretado": text,
            "fonte_utilizada": source, "codigo_sugerido": None, "label_sugerido": "",
            "metodo": "sem_correspondencia_geografica" if text else "vazio", "confianca": 0.0,
            "status": "NÃO IDENTIFICADO" if text else "VAZIO", "decisao_automatica": False,
            "candidatos": "", "localidade_base": "", "tipo_localidade_base": "", "municipio_base": "", "uf_base": "",
            "fontes_consultadas": fontes_consultadas,
        }

    # Prefer exact geographic resolutions; among equals, honor source-column priority.
    usable.sort(key=lambda item: (item[2].method != "geografica_exata", -item[2].score, config.source_columns.index(item[0])))
    source, text, resolution = usable[0]

    # If multiple source answers imply different target localities, flag rather than silently choosing.
    target_keys = {normalize_text(r.value) for _, _, r in usable if r.value}
    if len(target_keys) > 1:
        return {
            "ID": row.get(id_column), "variavel_controle": config.output_name,
            "tipo_geografico": config.geographic_type, "texto_interpretado": text,
            "fonte_utilizada": source, "codigo_sugerido": None, "label_sugerido": "",
            "metodo": "fontes_geograficas_conflitantes", "confianca": resolution.score,
            "status": "AMBÍGUO", "decisao_automatica": False,
            "candidatos": " | ".join(sorted({r.value for _, _, r in usable if r.value})),
            "localidade_base": resolution.matched_name or "", "tipo_localidade_base": resolution.matched_type or "",
            "municipio_base": resolution.municipality or (resolution.value if config.geographic_type == "municipio" else ""),
            "uf_base": resolution.uf,
            "fontes_consultadas": " | ".join(f"{c}={raw!s} -> {txt}" for c, raw, txt in attempts),
        }

    project_result = match_text(resolution.value, project_labels, config.fuzzy_cutoff)
    if project_result.code is None:
        # Antes de aceitar "fora da amostra", verifica se o texto é um
        # loteamento que fica DE VERDADE dentro de um bairro oficial que tem
        # um representante aprovado nos VALUE LABELS — isso é uma correção
        # legítima (mesmo lugar, nome popular diferente), bem diferente de
        # realocar para outro bairro por estar "perto"/na mesma região, o que
        # o motor nunca faz sozinho.
        if config.geographic_type == "bairro" and resolution.municipality:
            hierarchy = get_hierarchy(resolution.municipality, geography_db)
            if hierarchy:
                found = hierarchy.find(resolution.value) or hierarchy.find(text)
                if found:
                    representative = hierarchy.approved_representative(found.bairro_oficial, project_labels)
                    if representative:
                        rep_code, rep_label = representative
                        return {
                            "ID": row.get(id_column), "variavel_controle": config.output_name,
                            "tipo_geografico": config.geographic_type, "texto_interpretado": text,
                            "fonte_utilizada": source, "codigo_sugerido": rep_code, "label_sugerido": rep_label,
                            "metodo": f"hierarquia_bairro_oficial:{found.metodo}", "confianca": round(resolution.score, 4),
                            "status": "AJUSTADO AUTOMATICAMENTE", "decisao_automatica": True,
                            "candidatos": "", "localidade_base": resolution.matched_name or text,
                            "tipo_localidade_base": resolution.matched_type or "",
                            "municipio_base": resolution.municipality, "uf_base": resolution.uf,
                            "fontes_consultadas": " | ".join(f"{c}={raw!s} -> {txt}" for c, raw, txt in attempts),
                        }
        return {
            "ID": row.get(id_column), "variavel_controle": config.output_name,
            "tipo_geografico": config.geographic_type, "texto_interpretado": text,
            "fonte_utilizada": source, "codigo_sugerido": None, "label_sugerido": resolution.value,
            "metodo": resolution.method, "confianca": round(resolution.score, 4),
            "status": "FORA DA AMOSTRA", "decisao_automatica": False,
            "candidatos": ", ".join(resolution.candidates), "localidade_base": resolution.matched_name or text,
            "tipo_localidade_base": resolution.matched_type or "", "municipio_base": resolution.municipality or (resolution.value if config.geographic_type == "municipio" else ""),
            "uf_base": resolution.uf,
            "fontes_consultadas": " | ".join(f"{c}={raw!s} -> {txt}" for c, raw, txt in attempts),
        }

    code, label = project_result.code, project_result.label
    # Precisa de confiança nos DOIS lados: a resposta bateu com segurança contra
    # a Base Brasil E o nome identificado bateu com segurança contra os VALUE
    # LABELS do projeto. O nome "oficial" de um bairro na Base Brasil às vezes
    # não tem o mesmo prefixo popular que o projeto usa (ex.: Base Brasil chama
    # de "Sobrinho", o projeto usa "Vila Sobrinho") — por isso o segundo lado
    # também usa correspondência por substring/fuzzy, não só igualdade exata.
    geo_automatic = resolution.method in {"geografica_exata", "geografica_exata_sufixo_uf"}
    project_status = confidence_status(project_result, config.auto_fuzzy_threshold)
    project_automatic = project_status in {"CONFIRMADO", "AJUSTADO AUTOMATICAMENTE"}
    automatic = geo_automatic and project_automatic
    if automatic and normalize_text(text) == normalize_text(label):
        status = "CONFIRMADO"
    elif automatic:
        status = "AJUSTADO AUTOMATICAMENTE"
    else:
        status = "SUGESTÃO"
    return {
        "ID": row.get(id_column), "variavel_controle": config.output_name,
        "tipo_geografico": config.geographic_type, "texto_interpretado": text,
        "fonte_utilizada": source, "codigo_sugerido": code, "label_sugerido": label,
        "metodo": f"{resolution.method}+{project_result.method}", "confianca": round(min(resolution.score, project_result.score), 4), "status": status,
        "decisao_automatica": automatic, "candidatos": ", ".join(resolution.candidates),
        "localidade_base": resolution.matched_name or text, "tipo_localidade_base": resolution.matched_type or "",
        "municipio_base": resolution.municipality or (resolution.value if config.geographic_type == "municipio" else ""),
        "uf_base": resolution.uf,
        "fontes_consultadas": " | ".join(f"{c}={raw!s} -> {txt}" for c, raw, txt in attempts),
    }


def recode_dataframe(df, config, project_labels, bank_value_labels, id_column, geography_db=None):
    records = []
    for _, row in df.iterrows():
        if config.geographic_type and geography_db is not None:
            records.append(_recode_geographic_row(
                row, config, project_labels, bank_value_labels, id_column, geography_db
            ))
            continue

        text, source_used, attempts = choose_source_text(row, config.source_columns, bank_value_labels)
        result = match_text(text, project_labels, config.fuzzy_cutoff)
        status = confidence_status(result, config.auto_fuzzy_threshold)
        records.append({
            "ID": row.get(id_column),
            "variavel_controle": config.output_name,
            "tipo_geografico": config.geographic_type or "",
            "texto_interpretado": text,
            "fonte_utilizada": source_used or "",
            "codigo_sugerido": result.code,
            "label_sugerido": result.label or "",
            "metodo": result.method,
            "confianca": round(result.score, 4),
            "status": status,
            "decisao_automatica": status in {"CONFIRMADO", "AJUSTADO AUTOMATICAMENTE"},
            "candidatos": ", ".join(map(str, result.candidates)),
            "localidade_base": "", "tipo_localidade_base": "", "municipio_base": "", "uf_base": "",
            "fontes_consultadas": " | ".join(f"{c}={raw!s} -> {txt}" for c, raw, txt in attempts),
        })
    result_df = pd.DataFrame(records)
    if not result_df.empty:
        # Ajuda a diagnosticar "por que isso não bateu, se está escrito igual no
        # VALUE LABELS?" — mostra o texto exatamente como o motor de comparação
        # o enxerga (sem acento, minúsculo, sem pontuação), lado a lado com o
        # texto original interpretado.
        result_df["texto_normalizado"] = result_df["texto_interpretado"].map(normalize_text)
    return result_df
