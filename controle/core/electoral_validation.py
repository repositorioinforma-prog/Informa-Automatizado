"""
Validação eleitoral: verifica se o entrevistado vota dentro do(s) estado(s)
da pesquisa, em outro estado, ou se declarou que não vota / não vota mais.

Reaproveita o mesmo motor de resolução geográfica (Base Brasil) usado nas
variáveis de controle de Cidade/Bairro — mas SEMPRE contra a base nacional
completa (não filtrada por UF), porque aqui o objetivo é justamente
descobrir em qual UF a pessoa vota, não confirmar um lugar dentro de um
estado já conhecido.
"""
from __future__ import annotations

import re

import pandas as pd

from controle.core.normalization import normalize_text
from controle.core.recoding import value_to_text
from controle.geography.database import GeographyDatabase, _UF_SIGLA_NOME  # noqa: F401 (mapa de referência)
from controle.geography.resolver import resolve_to_target

# Rótulos comuns para "não vota" em pesquisas eleitorais brasileiras — tanto
# em VALUE LABELS de variáveis fechadas quanto em respostas de texto livre
# (campo "Outros"), incluindo erros de digitação recorrentes. Serve como
# sugestão automática (para códigos) e como checagem direta (para texto
# livre) — o usuário confirma/ajusta os códigos antes de aplicar.
NAO_VOTA_KEYWORDS = [
    "nao voto", "não voto", "nao vota", "não vota", "nao vou votar",
    "nao estou votando", "nao voto mais", "nao posso votar", "nao tenho titulo",
    "nao tenho título", "nao tenho titulo de eleitor", "nao possuo titulo",
    "sem titulo", "nunca tirei titulo", "nunca tirei o titulo",
    "ainda nao tirei titulo", "ainda nao tenho titulo", "titulo cancelado",
    "titulo canselado", "titulo cançelado", "titulo canceladu", "titulo suspenso",
    "titulo suspenço", "titulo irregular", "titulo irregula", "titulo bloqueado",
    "titulo bloquiado", "perdi o titulo", "perdi meu titulo",
    "nao regularizei o titulo", "nao regularizei", "nao transferi o titulo",
    "nao transferi", "nao fiz transferencia", "nao fiz tranferencia",
    "nao fiz transferensia", "nao sei onde voto", "nao sei aonde voto",
    "não sei onde voto", "n sei onde voto", "n sei aonde voto",
    "nao lembro onde voto", "nao lembro aonde voto", "n lembro onde voto",
    "nao tenho local de votacao", "nao tenho local de votação",
    "nao tenho cidade de votacao", "nao tenho cidade onde voto",
    "nao tenho domicilio eleitoral", "nao sou eleitor", "nao sou eleitora",
    "nao sou eleitro", "nao sou eleitora ainda", "nao sou obrigado a votar",
    "nao sou obrigada a votar", "nao preciso votar", "nao voto por idade",
    "menor de idade", "sou menor", "ainda sou menor", "tenho menos de 16",
    "tenho 15 anos", "tenho 14 anos", "ainda nao tenho idade",
    "nao tenho idade pra votar", "nao tenho idade para votar", "estrangeiro",
    "estrangeira", "sou estrangeiro", "sou estrangeira", "nao sou brasileiro",
    "nao sou brasileira", "nao tenho direito a voto", "nao posso votar",
    "impedido de votar", "impedida de votar", "nao voto no brasil", "moro fora",
    "moro fora do brasil", "nao voto aqui", "nao voto nessa cidade",
    "nao voto nesta cidade", "nenhuma", "nenhum", "nenhuma cidade",
    "cidade nenhuma", "lugar nenhum", "em nenhuma", "em lugar nenhum",
    "nao se aplica", "nao aplica", "n/a", "na", "nsa", "nao sabe", "nao sei",
    "n sei", "nao lembro", "n lembro", "esqueci", "esqueci onde voto",
    "esqueci a cidade", "nao recordo", "não recordo", "nao me recordo",
    "nao informado", "nao informei", "sem resposta", "prefiro nao responder",
    "nao votu", "nao vouto", "nao votoo", "nao vto", "n voto", "n vota",
    "n vou votar", "nao tenhu titulo", "nao tenho titilo", "nao tenho tituo",
    "nao tenho tilulo", "sem titilo", "titulo canceldo", "titulo canelado",
    "nao sei a cidade", "nao sei cidede", "nao sei a cidada",
    "nao sei onde vota", "nao lembro a cidade", "nao lenbro onde voto",
    "nao lembro onde vto", "nunca votei", "nunca voto", "nunca voutei",
    "nunca votey", "nunca tirei titilo", "nao tenho cadastro eleitoral",
    "sem cadastro eleitoral", "sem titulo eleitoral", "nao tenho zona eleitoral",
]

# Frases curtas/genéricas (até 4 caracteres normalizados) só valem como
# resposta INTEIRA — como substring, apareceriam dentro de nomes de lugar
# reais (ex.: "na" dentro de "Sena Madureira") e gerariam falso positivo.
_CURTAS = {normalize_text(k) for k in NAO_VOTA_KEYWORDS if len(normalize_text(k)) <= 4}
_LONGAS = [normalize_text(k) for k in NAO_VOTA_KEYWORDS if len(normalize_text(k)) > 4]


def looks_like_nao_vota_text(text: str) -> bool:
    """Confere um texto de resposta livre (ex.: campo 'Outros' de onde vota)
    contra os padrões conhecidos de 'não vota'."""
    norm = normalize_text(text)
    if not norm:
        return False
    if norm in _CURTAS:
        return True
    return any(re.search(r"(?<!\w)" + re.escape(kw) + r"(?!\w)", norm) for kw in _LONGAS)


def detect_nao_vota_codes(value_labels: dict) -> list:
    """Sugere, entre os VALUE LABELS de uma variável, quais códigos parecem
    significar 'não vota' — o usuário confirma antes de aplicar."""
    detected = []
    for code, label in value_labels.items():
        if looks_like_nao_vota_text(label):
            detected.append(code)
    return detected


def validate_electoral(
    df: pd.DataFrame,
    id_column: str,
    source_columns: tuple[str, ...],
    nao_vota_codes: set,
    bank_value_labels: dict,
    geography_db: GeographyDatabase,
    survey_ufs: set[str],
    fuzzy_cutoff: float = 0.82,
) -> pd.DataFrame:
    """Para cada entrevistado, classifica onde vota:
    - NÃO VOTA: o código respondido está em nao_vota_codes.
    - VOTA NO ESTADO DA PESQUISA: resolve para um município de UF em survey_ufs.
    - VOTA EM OUTRO ESTADO: resolve para um município de UF fora de survey_ufs.
    - AMBÍGUO: resolve para mais de um local possível, em UFs diferentes.
    - NÃO IDENTIFICADO: não deu pra resolver o texto/código.
    - VAZIO: nenhuma fonte respondida.
    """
    survey_ufs_norm = {u.upper() for u in survey_ufs}
    records = []
    for _, row in df.iterrows():
        raw_first = row.get(source_columns[0]) if source_columns else None
        if pd.notna(raw_first) and raw_first in nao_vota_codes:
            records.append({
                "ID": row.get(id_column), "status": "NÃO VOTA",
                "texto_interpretado": value_to_text(raw_first, source_columns[0], bank_value_labels) or "",
                "uf_identificada": "", "municipio_identificado": "",
            })
            continue

        text = ""
        source_used = ""
        for column in source_columns:
            raw = row.get(column)
            candidate = value_to_text(raw, column, bank_value_labels)
            if candidate:
                text, source_used = candidate, column
                break

        if not text:
            records.append({
                "ID": row.get(id_column), "status": "VAZIO",
                "texto_interpretado": "", "uf_identificada": "", "municipio_identificado": "",
            })
            continue

        if looks_like_nao_vota_text(text):
            records.append({
                "ID": row.get(id_column), "status": "NÃO VOTA",
                "texto_interpretado": text, "uf_identificada": "", "municipio_identificado": "",
                "fonte_utilizada": source_used,
            })
            continue

        resolution = resolve_to_target(
            text, "municipio", geography_db, fuzzy_cutoff=fuzzy_cutoff, preferred_ufs=frozenset(survey_ufs_norm),
        )
        if not resolution.value:
            records.append({
                "ID": row.get(id_column), "status": "NÃO IDENTIFICADO" if resolution.status != "AMBÍGUO" else "AMBÍGUO",
                "texto_interpretado": text, "uf_identificada": "", "municipio_identificado": "",
                "candidatos": ", ".join(resolution.candidates),
            })
            continue

        uf = (resolution.uf or "").upper()
        uf_sigla = next((s for s, n in _UF_SIGLA_NOME.items() if normalize_text(n) == normalize_text(uf) or s == uf), uf)
        status = "VOTA NO ESTADO DA PESQUISA" if uf_sigla in survey_ufs_norm else "VOTA EM OUTRO ESTADO"
        records.append({
            "ID": row.get(id_column), "status": status,
            "texto_interpretado": text, "uf_identificada": uf_sigla or uf,
            "municipio_identificado": resolution.value, "fonte_utilizada": source_used,
        })
    return pd.DataFrame(records)
