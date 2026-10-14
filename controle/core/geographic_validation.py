from __future__ import annotations

import pandas as pd

from controle.core.normalization import normalize_text


def project_label_lookup(labels: dict) -> dict[str, tuple[object, str]]:
    lookup = {}
    for code, label in labels.items():
        key = normalize_text(label)
        if key and key not in lookup:
            lookup[key] = (code, label)
    return lookup


def build_geographic_coherence_report(
    audit_df: pd.DataFrame,
    configs: list[dict],
    label_sets: dict[str, dict],
) -> pd.DataFrame:
    """Compare municipality control with municipality implied by bairro/distrito."""
    municipal = [c for c in configs if c.get("geographic_type") == "municipio"]
    children = [c for c in configs if c.get("geographic_type") in {"bairro", "distrito"}]
    if not municipal or not children or audit_df.empty:
        return pd.DataFrame()

    city_cfg = municipal[0]
    city_rows = audit_df[audit_df["variavel_controle"] == city_cfg["output_name"]].copy()
    city_rows = city_rows.set_index("ID", drop=False)
    city_project = project_label_lookup(label_sets[city_cfg["label_set_name"]])
    issues = []

    for child_cfg in children:
        child_rows = audit_df[audit_df["variavel_controle"] == child_cfg["output_name"]]
        for _, child in child_rows.iterrows():
            rid = child["ID"]
            implied = str(child.get("municipio_base", "") or "").strip()
            if not implied or rid not in city_rows.index:
                continue
            city = city_rows.loc[rid]
            if isinstance(city, pd.DataFrame):
                city = city.iloc[0]
            city_value = str(city.get("label_sugerido", "") or "").strip()
            if not city_value or normalize_text(city_value) == normalize_text(implied):
                continue

            valid = city_project.get(normalize_text(implied))
            if valid:
                status = "INCONSISTENTE"
                recommendation = f"Alterar {city_cfg['output_name']} para {valid[1]}"
                code = valid[0]
                reason = (
                    f"{child_cfg['output_name']} foi identificado como {child.get('localidade_base', '')}, "
                    f"pertencente a {implied}. O município está nos VALUE LABELS válidos do projeto."
                )
            else:
                status = "FORA DA AMOSTRA"
                recommendation = "Revisar; não recodificar automaticamente"
                code = None
                reason = (
                    f"{child_cfg['output_name']} indica o município {implied}, mas esse município não está "
                    f"nos VALUE LABELS de {city_cfg['output_name']} deste projeto."
                )
            issues.append({
                "ID": rid,
                "status": status,
                "cidade_controle_atual": city_value,
                "localidade_informada": child.get("texto_interpretado", ""),
                "localidade_identificada": child.get("localidade_base", ""),
                "municipio_base": implied,
                "codigo_cidade_sugerido": code,
                "recomendacao": recommendation,
                "motivo": reason,
            })
    return pd.DataFrame(issues)


def build_cidade_bairro_realocation(
    df,
    audit_df: pd.DataFrame,
    city_cfg: dict,
    child_cfg: dict,
    label_sets: dict[str, dict],
    bank_value_labels: dict,
    geography_db,
    id_column: str,
) -> pd.DataFrame:
    """Verificação cruzada de dois sentidos entre Cidade e Bairro:

    Caso 1 — o bairro respondido pertence, na Base Brasil, a um município da
    amostra diferente do que foi marcado em Cidade (reaproveita a checagem de
    coerência já existente).

    Caso 2 — o que a pessoa escreveu no campo de Bairro é, na verdade, o nome
    de outro município da amostra (ex.: escreveu "Duque de Caxias" onde se
    esperava um bairro). Para isso, resolvemos o mesmo texto do campo de
    Bairro como se fosse uma resposta de Cidade, usando o mesmo motor.

    Em ambos os casos, só sugerimos realocação quando o município alternativo
    identificado está nos VALUE LABELS válidos do projeto — nunca inventamos
    uma cidade fora da amostra. Quando o MESMO texto bate como nome de bairro
    (caso 1) e como nome de outro município (caso 2) mas os dois caminhos
    apontam para cidades DIFERENTES entre si (ex.: existe um bairro chamado
    "Tanguá" em uma cidade E também um município chamado Tanguá), isso é
    tratado como ambíguo — nenhum dos dois é aplicado sozinho.
    """
    from controle.core.recoding import ControlVariableConfig, recode_dataframe  # import local evita ciclo

    city_rows = audit_df[audit_df["variavel_controle"] == city_cfg["output_name"]].set_index("ID", drop=False)

    coherence = build_geographic_coherence_report(audit_df, [city_cfg, child_cfg], label_sets)
    if coherence.empty:
        # [Informa] sem nenhuma inconsistência o relatório vem sem colunas (o original quebrava aqui)
        coherence = pd.DataFrame(columns=["ID", "status", "codigo_cidade_sugerido", "municipio_base",
                                          "localidade_informada", "motivo"])
    coherence = coherence[coherence["status"] == "INCONSISTENTE"].set_index("ID", drop=False)

    bairro_via_cidade_cfg = ControlVariableConfig(
        output_name="_verificacao_bairro_como_cidade",
        label_set_name=city_cfg["label_set_name"],
        source_columns=child_cfg["source_columns"],
        geographic_type="municipio",
        fuzzy_cutoff=city_cfg.get("fuzzy_cutoff", 0.82),
        auto_fuzzy_threshold=city_cfg.get("auto_fuzzy_threshold", 0.93),
    )
    audit_bairro_como_cidade = recode_dataframe(
        df, bairro_via_cidade_cfg, label_sets[city_cfg["label_set_name"]],
        bank_value_labels, id_column, geography_db=geography_db,
    ).set_index("ID", drop=False)

    all_ids = set(coherence.index) | {
        rid for rid, row in audit_bairro_como_cidade.iterrows()
        if row["decisao_automatica"] and pd.notna(row["codigo_sugerido"])
    }

    resultados = []
    for rid in all_ids:
        if rid not in city_rows.index:
            continue
        atual = city_rows.loc[rid]
        if isinstance(atual, pd.DataFrame):
            atual = atual.iloc[0]
        codigo_atual = atual.get("codigo_sugerido")

        cand1 = coherence.loc[rid] if rid in coherence.index else None
        if isinstance(cand1, pd.DataFrame):
            cand1 = cand1.iloc[0]
        cod1 = cand1["codigo_cidade_sugerido"] if cand1 is not None else None

        cand2 = audit_bairro_como_cidade.loc[rid] if rid in audit_bairro_como_cidade.index else None
        if isinstance(cand2, pd.DataFrame):
            cand2 = cand2.iloc[0]
        cod2 = None
        if cand2 is not None and cand2["decisao_automatica"] and pd.notna(cand2["codigo_sugerido"]):
            cod2 = cand2["codigo_sugerido"]

        if cod1 is None and cod2 is None:
            continue
        if cod1 is not None and cod2 is not None and cod1 != cod2:
            # o mesmo texto bate como bairro de uma cidade E como nome de outra
            # cidade — os dois caminhos discordam, não decide sozinho.
            if codigo_atual in (cod1, cod2):
                continue  # a cidade atual já é uma das duas leituras possíveis; deixa como está
            resultados.append({
                "ID": rid,
                "status": "AMBÍGUO",
                "cidade_controle_atual": atual.get("label_sugerido", ""),
                "localidade_informada": (cand1["localidade_informada"] if cand1 is not None else cand2["texto_interpretado"]),
                "municipio_base": f"{cand1['municipio_base']} (como bairro) ou {cand2['label_sugerido']} (como cidade)" if cand1 is not None and cand2 is not None else "",
                "codigo_cidade_sugerido": None,
                "recomendacao": "Revisar manualmente; duas leituras possíveis e diferentes",
                "motivo": (
                    f"O texto de {child_cfg['output_name']} bate como bairro pertencente a "
                    f"'{cand1['municipio_base'] if cand1 is not None else ''}' e, separadamente, como o "
                    f"nome do município '{cand2['label_sugerido'] if cand2 is not None else ''}'. As duas "
                    f"leituras discordam entre si."
                ),
                "motivo_realocacao": "Ambíguo: bate como bairro de uma cidade e como nome de outra",
                "fonte": child_cfg["output_name"],
            })
            continue

        codigo_sugerido = cod1 if cod1 is not None else cod2
        if codigo_sugerido == codigo_atual:
            continue  # já bate com o que está em Cidade, nada a fazer

        if cod1 is not None and cod2 is not None:
            motivo_tipo = "Confirmado nos dois sentidos: bate como bairro e como nome do mesmo outro município"
            municipio_txt = cand1["municipio_base"]
            localidade_txt = cand1["localidade_informada"]
            motivo = (
                f"{child_cfg['output_name']} bate tanto como bairro pertencente a '{municipio_txt}' quanto "
                f"como o próprio nome do município '{municipio_txt}' — as duas leituras concordam."
            )
        elif cod1 is not None:
            motivo_tipo = "Bairro informado pertence a outro município da amostra"
            municipio_txt = cand1["municipio_base"]
            localidade_txt = cand1["localidade_informada"]
            motivo = cand1["motivo"]
        else:
            motivo_tipo = "Campo de Bairro contém, na verdade, o nome de outro município"
            municipio_txt = cand2["label_sugerido"]
            localidade_txt = cand2["texto_interpretado"]
            motivo = (
                f"O texto respondido em {child_cfg['output_name']} ('{localidade_txt}') corresponde ao "
                f"município '{municipio_txt}', da amostra do projeto."
            )

        resultados.append({
            "ID": rid,
            "status": "INCONSISTENTE",
            "cidade_controle_atual": atual.get("label_sugerido", ""),
            "localidade_informada": localidade_txt,
            "municipio_base": municipio_txt,
            "codigo_cidade_sugerido": codigo_sugerido,
            "recomendacao": f"Alterar {city_cfg['output_name']} para {municipio_txt}",
            "motivo": motivo,
            "motivo_realocacao": motivo_tipo,
            "fonte": child_cfg["output_name"],
        })

    return pd.DataFrame(resultados)
