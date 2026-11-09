"""
Esqueleto da apresentação (Relatório Analítico) — lógica pura (sem Streamlit).

Parte do molde `assets/molde_informa.pptx` e, a partir da configuração do estudo,
gera:
  1. um .pptx "esqueleto": capa, metodologia, rodapés e bases já preenchidos;
     módulos escolhidos; tudo que é de outro estudo limpo; e um aviso visível de
     "conteúdo pendente" em cada slide de resultado (nome do shape: INFORMA_PENDENTE);
  2. um manifesto (dict/JSON) que diz, slide a slide, o que já está pronto, o que
     está pendente, de qual arquivo vem cada dado e quais divergências o verificador
     achou — para a IA (ou a pessoa) completar os dados nos lugares certos.

Regras do projeto respeitadas: nunca inventar valor, nunca usar zero como
preenchimento, nunca deixar dado de outro estudo. Campo sem informação vira
"(pendente)" e entra na lista `campos_pendentes` do manifesto.
"""
from __future__ import annotations

import copy
import io
import math
import re
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import date, datetime

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.parts.chart import ChartPart
from pptx.util import Emu, Pt

from core.verificador_relatorio_math import chave

VERSAO_MANIFESTO = 1
PENDENTE = "(pendente)"
ID_PENDENTE = "INFORMA_PENDENTE"

ESTADOS = {
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas", "BA": "Bahia", "CE": "Ceará",
    "DF": "Distrito Federal", "ES": "Espírito Santo", "GO": "Goiás", "MA": "Maranhão",
    "MT": "Mato Grosso", "MS": "Mato Grosso do Sul", "MG": "Minas Gerais", "PA": "Pará",
    "PB": "Paraíba", "PR": "Paraná", "PE": "Pernambuco", "PI": "Piauí", "RJ": "Rio de Janeiro",
    "RN": "Rio Grande do Norte", "RS": "Rio Grande do Sul", "RO": "Rondônia", "RR": "Roraima",
    "SC": "Santa Catarina", "SP": "São Paulo", "SE": "Sergipe", "TO": "Tocantins",
}
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
         "setembro", "outubro", "novembro", "dezembro"]

CARGOS = ["Presidente", "Governador", "Senador", "Deputado Federal", "Deputado Estadual", "Prefeito"]
ANALISES = {
    "cruzamentos": "Cruzamentos",
    "residuais": "Análises residuais",
    "correlacao": "Correlação",
    "mapas_calor": "Mapas de calor",
    "correspondencia": "Mapas de correspondência",
    "analise_perguntas": "Análise das perguntas",
}
GOVERNOS = {"federal": "Federal", "estadual": "Estadual", "municipal": "Municipal"}

# Nome do arquivo-fonte (como aparece no aviso de pendência e no manifesto)
FONTES = {
    "descritivo": "Relatório Descritivo",
    "cruzamentos": "Cruzamentos",
    "residuais": "Análises residuais (A.R.)",
    "correlacao": "Correlação",
    "mapas_calor": "Mapas de calor",
    "correspondencia": "Mapas de correspondência",
    "analise_perguntas": "Análise das perguntas",
    "perfil": "Perfil",
    "briefing": "briefing (informar na configuração)",
}

# ---------------------------------------------------------------------------
# Catálogo de módulos. `slides`: lista de (índice no molde, análise exigida, fonte, papel).
# A ordem do catálogo é a ordem final da apresentação (encerramento sempre por último).
# ---------------------------------------------------------------------------
def _s(idx, requer=None, fonte="descritivo", papel=None, subst=None):
    return {"idx": idx, "requer": requer, "fonte": fonte, "papel": papel, "subst": subst or []}


CATALOGO = [
    {"id": "abertura", "grupo": "Fixo", "rotulo": "Abertura institucional (capa, quem somos, como fazemos)",
     "fixo": True, "slides": [_s(i, fonte=None) for i in (1, 2, 3, 4, 5)]},
    {"id": "metodologia", "grupo": "Fixo", "rotulo": "Metodologia (objetivos, coleta, amostra, abrangência, auditoria)",
     "fixo": True, "slides": [_s(6, fonte=None), _s(7, fonte=None), _s(8, fonte=None), _s(9, fonte=None)]},
    {"id": "perfil", "grupo": "Fixo", "rotulo": "Perfil dos respondentes",
     "fixo": True, "slides": [_s(10, fonte=None), _s(11, fonte=None), _s(12, fonte="perfil", papel="perfil")]},
    {"id": "avaliacao_governos", "grupo": "Avaliação dos governos", "rotulo": "Avaliação dos governos (um slide por governo)",
     "default": True, "slides": [_s(13, fonte=None)]},  # + slides por governo (tratados em código)
    {"id": "comparativo_governos", "grupo": "Avaliação dos governos", "rotulo": "Comparativo Federal, Estadual e Municipal",
     "default": True, "slides": [_s(15, papel="comparativo_governos")]},
    {"id": "reflexao_60x30", "grupo": "Avaliação dos governos", "rotulo": "Reflexão sobre a relação 60×30",
     "default": False, "slides": [_s(16, papel="reflexao_60x30")]},
    {"id": "presidente", "grupo": "Cargos", "rotulo": "Presidente da República", "cargo": "Presidente", "default": True,
     "slides": [_s(17, fonte=None), _s(18, papel="intencao"), _s(19, "cruzamentos", "cruzamentos", "cruzamento"),
                _s(20, "residuais", "residuais", "residual")]},
    {"id": "governador", "grupo": "Cargos", "rotulo": "Governador (clonado do módulo Presidente)", "cargo": "Governador",
     "default": True, "clona": "presidente"},
    {"id": "senador", "grupo": "Cargos", "rotulo": "Senador", "cargo": "Senador", "default": True,
     "slides": [_s(21, papel="intencao"), _s(22, papel="soma"), _s(23, "cruzamentos", "cruzamentos", "cruzamento"),
                _s(24, papel="rejeicao")]},
    {"id": "deputados", "grupo": "Cargos", "rotulo": "Deputado Federal / Estadual", "cargo": "Deputado", "default": True,
     "slides": [_s(25, papel="intencao"), _s(26, "cruzamentos", "cruzamentos", "cruzamento")]},
    {"id": "prefeito", "grupo": "Cargos", "rotulo": "Prefeito (cenários 2028)", "cargo": "Prefeito", "default": False,
     "slides": [_s(27, "cruzamentos", "cruzamentos", "segmento"), _s(28, "cruzamentos", "cruzamentos", "cruzamento")]},
    {"id": "analise_perguntas", "grupo": "Análise das perguntas", "rotulo": "Análise das perguntas (divisória + slide-modelo por pergunta)",
     "default": True,
     "slides": [_s(17, "analise_perguntas", None, "divisoria_perguntas",
                   [("ELEIÇÕES 2026", "ANÁLISE DAS PERGUNTAS"), ("Presidente da República", "Perguntas × segmentos")]),
                _s(19, "analise_perguntas", "analise_perguntas", "analise_perguntas",
                   [("Cruzamento", "Análise das perguntas"), ("Voto estimulado para Presidente", "Pergunta × segmentos (repetir para cada pergunta)")])]},
    {"id": "analises_estrategicas", "grupo": "Análises estratégicas", "rotulo": "Análises estratégicas (divisória + explicações)",
     "default": True, "slides": [_s(30, fonte=None), _s(31, "correlacao", None), _s(32, "residuais", None)]},
    {"id": "candidato_gestao", "grupo": "Candidato foco (CAV)", "rotulo": "Percepção da gestão anterior e comparativo de avaliação",
     "default": False, "slides": [_s(33, papel="gestao_anterior"), _s(34, papel="comparativo_medias")]},
    {"id": "candidato_perfil_ipv", "grupo": "Candidato foco (CAV)", "rotulo": "Perfil de candidato e IPV do candidato",
     "default": False, "slides": [_s(35, papel="perfil_candidato"), _s(36, papel="ipv"),
                                  _s(37, "cruzamentos", "cruzamentos", "segmento"), _s(38, papel="ipv_apoio"),
                                  _s(39, "cruzamentos", "cruzamentos", "cruzamento")]},
    {"id": "cav", "grupo": "Candidato foco (CAV)", "rotulo": "CAV completo (níveis 1, 2 e 3)", "default": False,
     "slides": [_s(40, fonte=None), _s(41, fonte=None), _s(42, papel="nivel1"), _s(43, papel="nivel2", fonte="descritivo"),
                _s(44, "correspondencia", "correspondencia", "correspondencia"),
                _s(45, papel="comparativo_ipv"), _s(46, "cruzamentos", "cruzamentos", "cruzamento"),
                _s(47, "correlacao", "correlacao", "correlacao")]},
    {"id": "cav_enxuto", "grupo": "Candidato foco (CAV)", "rotulo": "CAV enxuto (só Descritivo)", "default": False,
     "slides": [_s(48, papel="avaliacao_lado_a_lado"), _s(49, papel="posicionamento"), _s(50, papel="ipv_tabela"),
                _s(51, papel="percepcao_atuacao")]},
    {"id": "encerramento", "grupo": "Fixo", "rotulo": "Encerramento", "fixo": True, "slides": [_s(29, fonte=None)]},
]
CATALOGO_POR_ID = {m["id"]: m for m in CATALOGO}

# slide do molde usado como base de cada governo (layout com mapas de calor)
SLIDE_AVALIACAO = 14
SLIDE_AVALIACAO_SEM_MAPAS = 48

# ---------------------------------------------------------------------------
# Limpeza de dados de outros estudos presentes no molde
# ---------------------------------------------------------------------------
OUTROS_LOCAIS = ["Paty do Alferes", "Santana de Parnaíba", "Santana do Parnaíba", "Osasco", "Alphaville"]

# Textos de tabela/gráfico que são estrutura (ficam); todo o resto vira "(pendente)".
_GENERICOS = {chave(t) for t in [
    "Mín.:", "Entre:", "Máx.:", "Região", "Governo", "Avaliação governamental",
    "% de aproveitamento da avaliação governamental", "Aproveitamento", "Aprovação", "Regular",
    "Reprovação", "Potencial eleitoral do Governo", "Total", "Voto no segundo turno em 2022", "Lula",
    "Jair Bolsonaro", "Não votou", "Base", "Base reduzida", "Voto para Governador", "Voto para Presidente",
    "Cenário 1", "Cenário 2", "Faixa Etária (em anos)", "Renda Média Mensal Domiciliar (em Salário Mínimo)",
    "Escolaridade", "Regiões", "50 ou mais", "Até 02", "Mais de 05", "Ensino Fundamental", "Ensino Médio",
    "Ensino Superior", "Médias", "Religiões", "02 e 05", "05 e 10", "Mais de 10", "Católicos", "Evangélicos",
    "Voto no segundo turno para Presidente da República em 2022", "Certamente voto", "Provavelmente voto",
    "Provavelmente não voto", "Certamente não voto", "Não conheço suficiente para votar",
    "Não conhece suficiente para votar", "Depende do candidato", "Quem é? (Espontânea)", "Não lembra o nome",
    "Antes do acréscimo de informações", "Após o acréscimo de informações", "Aval. Gov. Municipal",
    "Aval. Gov. Federal Lula", "Federal Lula", "Avaliação do Governo Federal do Lula", "Aprova", "Desaprova",
    "Masculino", "Feminino", "16 a 29", "30 a 39", "40 a 49", "Conhece muito", "Conhece de nome",
    "Conhece pela foto", "Nunca viu antes", "Ajuda muito", "Ajuda um pouco", "Não ajuda", "Não sei dizer",
    "Categoria 1", "CERTAMENTE VOTO", "PROVAVELMENTE VOTO", "PROVAVELMENTE NÃO VOTO", "CERTAMENTE NÃO VOTO",
    "NÃO CONHEÇO O SUFICIENTE PARA VOTAR", "Depende do candidato",
    "Deputado Estadual apoiado por Bolsonaro", "Deputado Estadual apoiado por Lula",
    "Não conheço suficiente para votar", "Provavelmente não voto",
]}


def _generico(texto: str) -> bool:
    t = chave(texto)
    if not t:
        return True
    if re.fullmatch(r"[\d\s.,%-]+", t):  # números, percentuais
        return True
    return t in _GENERICOS


# Correções de erros conhecidos do molde (registradas no manifesto)
CORRECOES_MOLDE = [
    ("Após do acréscimo de informações", "Após o acréscimo de informações"),
    ("míltipla", "múltipla"),
]

# Termos que denunciam resto de outro estudo (usados na verificação final)
TERMOS_RESIDUAIS = [
    "paty do alferes", "alphaville", "osasco", "santana de parna", "santana do parna", "julinho", "ricardo couto",
    "tarcisio", "lindoso", "pedro paulo", "benedita", "crivella", "emidio de souza", "wesley", "pavanato",
    "renato bolsonaro", "gil diniz", "jorge felipe", "douglas ruas", "eduardo paes", "garotinho",
    "juninho bernardes", "batata", "{{", "}}",
]


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------
@dataclass
class ConfigEstudo:
    local: str = ""
    uf: str = "RJ"
    mes_ano: str = ""                      # ex.: "Setembro de 2026" (vazio → derivado do campo)
    titulo_capa: str = "DIAGNÓSTICO POLÍTICO ELEITORAL - PESQUISA QUANTITATIVA"
    modo: str = "rua"                      # "rua" | "online"
    campo_inicio: date | None = None
    campo_fim: date | None = None
    tempo_medio: int | None = None         # minutos
    base: int | None = None                # nº de entrevistas
    margem: str = ""                       # ex.: "5%" ou "1,64%"
    nivel_confianca: str = "95%"
    idade_minima: int | None = None        # vazio → 16 (rua) / 18 (online)
    regioes: str = ""                      # texto livre: uma região por linha ("Região 1 (28,49%): bairros...")
    n_regioes: int | None = None
    governador: str = ""
    prefeito: str = ""
    objetivo_especifico: str = ""
    projeto_rodape: str = ""               # vazio → "Projeto <local>, <estado>"
    cargos: list[str] = field(default_factory=lambda: ["Presidente", "Governador", "Senador", "Deputado Federal", "Deputado Estadual"])
    governos: list[str] = field(default_factory=lambda: ["federal", "estadual", "municipal"])
    analises: dict[str, bool] = field(default_factory=lambda: {k: True for k in ANALISES})
    modulos_extras: list[str] = field(default_factory=lambda: ["analises_estrategicas"])   # ids opcionais: candidato_*, cav, cav_enxuto, reflexao_60x30
    candidato_foco: str = ""
    cargo_foco: str = "Deputado Estadual"
    gestao_anterior: str = ""
    apresentacao_candidato: str = ""

    @property
    def estado(self) -> str:
        return ESTADOS.get(self.uf.upper(), self.uf)

    @property
    def rodape(self) -> str:
        return self.projeto_rodape.strip() or f"Projeto {self.local}, {self.estado}"

    @classmethod
    def de_dict(cls, d: dict) -> "ConfigEstudo":
        validos = {k: v for k, v in d.items() if k in cls.__dataclass_fields__}
        cfg = cls(**validos)
        base = {k: True for k in ANALISES}
        base.update(cfg.analises or {})
        cfg.analises = base
        return cfg

    def para_dict(self) -> dict:
        d = asdict(self)
        for k in ("campo_inicio", "campo_fim"):
            d[k] = d[k].isoformat() if d[k] else None
        return d


def sugerir_margem_erro(n: int | None) -> str:
    """Margem (95%, amostra aleatória simples) = 98/√n, no padrão dos decks: inteiro, ou 2 casas se < 2,5."""
    if not n or n <= 0:
        return ""
    v = 98 / math.sqrt(n)
    return f"{round(v)}%" if v >= 2.5 else f"{v:.2f}".replace(".", ",") + "%"


def texto_periodo_campo(ini: date | None, fim: date | None) -> str:
    if not ini and not fim:
        return ""
    ini = ini or fim
    fim = fim or ini
    if ini == fim:
        return f"no dia {ini.day:02d} de {MESES[ini.month - 1]} de {ini.year}"
    if (ini.month, ini.year) == (fim.month, fim.year):
        return f"entre os dias {ini.day:02d} e {fim.day:02d} de {MESES[fim.month - 1]} de {fim.year}"
    return (f"entre os dias {ini.day:02d} de {MESES[ini.month - 1]} e "
            f"{fim.day:02d} de {MESES[fim.month - 1]} de {fim.year}")


def _milhar(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def mes_ano_padrao(cfg: ConfigEstudo) -> str:
    if cfg.mes_ano.strip():
        return cfg.mes_ano.strip()
    ref = cfg.campo_fim or cfg.campo_inicio or date.today()
    return f"{MESES[ref.month - 1].capitalize()} de {ref.year}"


# ---------------------------------------------------------------------------
# Sugestões a partir do inventário do Descritivo
# ---------------------------------------------------------------------------
_PALAVRAS_CARGO = {
    "Presidente": "presidente", "Governador": "governador", "Senador": "senador",
    "Deputado Federal": "deputado federal", "Deputado Estadual": "deputado estadual", "Prefeito": "prefeito",
}


def detectar_cargos(inv: dict | None) -> list[str]:
    if not inv:
        return []
    textos = [chave(t["titulo"]) for t in inv.get("totais", []) + inv.get("segmentos", [])]
    achados = []
    for cargo, pal in _PALAVRAS_CARGO.items():
        if any(pal in t and "intencao" in t for t in textos):
            achados.append(cargo)
    return achados


def sugerir_nomes(inv: dict | None) -> dict:
    """Prefeito e governador citados nos títulos de avaliação do Descritivo (sugestão para confirmar)."""
    nomes = {"prefeito": "", "governador": ""}
    for t in (inv or {}).get("totais", []):
        tit = " ".join(t["titulo"].split())
        if not nomes["prefeito"]:
            m = re.search(r"Avalia[çc][ãa]o do Governo do Prefeito (.+?) \(", tit)
            if m:
                nomes["prefeito"] = m.group(1).strip()
        if not nomes["governador"]:
            m = re.search(r"Avalia[çc][ãa]o do Governador(?: Interino)?(?: *[-–] *)?(.+?) \(", tit)
            if m and m.group(1).strip():
                nomes["governador"] = m.group(1).strip()
    return nomes


def sugerir_base(inv: dict | None) -> int | None:
    if not inv:
        return None
    bases = [int(t["base_total"]) for t in inv.get("totais", []) if t.get("base_total")]
    if not bases:
        return None
    contagem = Counter(bases)
    topo = max(contagem.values())
    return max(b for b, c in contagem.items() if c == topo)


def _tem_titulo(inv: dict | None, todos: list[str], algum: list[str] = ()) -> bool:
    """Alguma tabela (títulos + pergunta) do Descritivo casa com as palavras (normalizadas)?"""
    for t in (inv or {}).get("totais", []) + (inv or {}).get("segmentos", []):
        k = chave(f"{t.get('titulo', '')} {t.get('pergunta', '')}")
        if all(x in k for x in todos) and (not algum or any(x in k for x in algum)):
            return True
    return False


def sugerir_objetivos(inv: dict | None, cfg: "ConfigEstudo") -> list[dict]:
    """Objetivos específicos (objetivo 2 da metodologia) sugeridos a partir do que os arquivos e a configuração têm.

    Cada item: {"id", "texto", "origem"}. Só entra o que tem base concreta (título de tabela do Descritivo,
    cargo/governo/análise marcados ou candidato foco informado) — nada é inventado: a pessoa escolhe e edita.
    """
    local = cfg.local.strip() or PENDENTE
    out: list[dict] = []

    def add(id_, texto, origem):
        out.append({"id": id_, "texto": texto, "origem": origem})

    extras = {x for x in cfg.modulos_extras if not x.startswith("-")}
    cand = cfg.candidato_foco.strip()
    cargo_foco = (cfg.cargo_foco or "").strip()
    cav = extras & {"cav", "cav_enxuto", "candidato_gestao", "candidato_perfil_ipv"}
    if cand and cav:
        add("candidato_foco",
            f"Avaliar o potencial eleitoral de {cand} na disputa para {cargo_foco or PENDENTE} em {local}, "
            "mensurando conhecimento, intenção de voto e a percepção do eleitorado sobre sua atuação.",
            "Candidato foco + módulo CAV marcado")
    elif cav:
        add("candidato_foco",
            f"Avaliar o potencial eleitoral do candidato foco ({PENDENTE}) na disputa para {cargo_foco or PENDENTE} em {local}.",
            "Módulo CAV marcado (informar o candidato foco)")
    if _tem_titulo(inv, ["transferencia"]):
        add("transferencia_votos",
            "Simular a transferência de votos entre os candidatos a Deputado em cenário no qual um dos nomes não é candidato.",
            "Descritivo: pergunta de transferência de votos")
    if _tem_titulo(inv, ["senador"], ["segundo", "voto 2", "2o voto"]) or _tem_titulo(inv, ["senador", "soma"]):
        add("senado_dois_votos",
            f"Mensurar o desempenho dos pré-candidatos ao Senado em {local}, considerando o primeiro e o segundo voto.",
            "Descritivo: intenção para Senador com 1º e 2º voto")
    if _tem_titulo(inv, ["rejeicao"]):
        add("rejeicao",
            "Identificar o nível de rejeição aos nomes apresentados, apontando os candidatos com maior resistência do eleitorado.",
            "Descritivo: perguntas de rejeição")
    if _tem_titulo(inv, ["espontanea"]):
        add("lembranca_espontanea",
            "Medir a lembrança espontânea de candidatos, antes da apresentação dos nomes, como indicador de conhecimento.",
            "Descritivo: pergunta espontânea")
    if _tem_titulo(inv, ["ipv"]) or _tem_titulo(inv, ["chance", "votar"]):
        add("ipv",
            "Calcular o índice de probabilidade de voto (IPV) dos candidatos, separando quem tem voto garantido, "
            "voto possível e rejeição.",
            "Descritivo: perguntas de probabilidade de voto")
    if _tem_titulo(inv, ["melhor", "acao"]) or _tem_titulo(inv, ["pior", "acao"]) or _tem_titulo(inv, ["principal", "problema"]):
        add("gestao_acoes",
            "Identificar as ações e os problemas de maior destaque na percepção da população sobre a gestão, "
            "para orientar prioridades.",
            "Descritivo: perguntas sobre ações/problemas da gestão")
    if _tem_titulo(inv, ["midia"]) or _tem_titulo(inv, ["rede", "social"]) or _tem_titulo(inv, ["habito"]):
        add("habitos_comunicacao",
            "Mapear os hábitos de comunicação e consumo de mídia do eleitorado, para apoiar a escolha de canais de diálogo.",
            "Descritivo: perguntas de mídia/redes sociais")
    if _tem_titulo(inv, ["cenario", "2"]) and _tem_titulo(inv, ["deputado"]):
        add("cenarios_deputado",
            "Comparar cenários de disputa para Deputado, observando como a presença ou ausência de candidatos altera o desempenho de cada nome.",
            "Descritivo: mais de um cenário para Deputado")
    if "Prefeito" in cfg.cargos:
        add("prefeito_2028",
            "Avaliar o cenário antecipado da eleição municipal de 2028, identificando a intenção de voto para prefeito "
            "em diferentes configurações de disputa.",
            "Cargo Prefeito marcado")
    if "municipal" in cfg.governos:
        pref = cfg.prefeito.strip()
        add("gestao_municipal",
            f"Avaliar a aprovação da gestão municipal{(' de ' + pref) if pref else ''} em {local} e seu efeito sobre o cenário eleitoral.",
            "Governo municipal marcado")
    if cfg.analises.get("analise_perguntas"):
        add("analise_perguntas",
            "Identificar diferenças de percepção entre os segmentos (sexo, faixa etária, renda, escolaridade, religião, "
            "região e estado civil) em cada pergunta do questionário.",
            "Análise das perguntas marcada")
    if cfg.analises.get("residuais") or cfg.analises.get("correlacao"):
        add("associacoes",
            "Identificar associações entre a avaliação dos governos e a intenção de voto, por meio de análises de correlação e de resíduos.",
            "Correlação / Análises residuais marcadas")
    return out


def compor_objetivo(textos: list[str]) -> str:
    """Junta os objetivos escolhidos em um único parágrafo (ignora vazios)."""
    return " ".join(" ".join(t.split()) for t in textos if t and t.strip())


def _titulo_casa(k: str, todos: list[str], algum: list[str] = (), excluir: list[str] = ()) -> bool:
    return all(x in k for x in todos) and (not algum or any(x in k for x in algum)) and not any(x in k for x in excluir)


_INTENCAO_EXCLUI = ["rejeicao", "ipv", "chance", "espontanea", "segundo candidato", "transferencia"]


def _regra_pergunta(papel: str, pal: str | None):
    """(todos, algum, excluir) sobre o título normalizado, ou None se o papel não tem pergunta própria."""
    if papel in ("intencao", "soma") and pal == "senador":
        return ["intencao", "senador"], [], ["rejeicao"]
    if papel in ("intencao", "cruzamento", "residual") and pal:
        return ["intencao", pal], [], _INTENCAO_EXCLUI
    if papel == "rejeicao" and pal:
        return ["rejeicao", pal], [], []
    if papel == "avaliacao_federal":
        return ["avaliacao"], ["lula", "presidente", "federal"], ["melhor", "pior", "aprovacao", "camara", "intencao"]
    if papel == "avaliacao_estadual":
        return ["avaliacao"], ["governador", "estadual"], ["melhor", "pior", "aprovacao", "camara", "intencao"]
    if papel == "avaliacao_municipal":
        return ["avaliacao"], ["prefeito", "municipal"], ["melhor", "pior", "aprovacao", "camara", "intencao"]
    if papel in ("intencao", "soma") and pal == "senador":
        return ["intencao", "senador"], [], ["rejeicao"]
    return None


def _sugerir_pergunta(inv: dict | None, cargo: str | None, papel: str | None) -> dict | None:
    """Tabela dos totais do Descritivo que provavelmente alimenta o slide (a IA confirma)."""
    if not inv or not papel or cargo == "Deputado":
        return None
    pal = _PALAVRAS_CARGO.get(cargo, chave(cargo)) if cargo else None
    regra = _regra_pergunta(papel, pal)
    if not regra:
        return None
    for t in inv.get("totais", []):
        if _titulo_casa(chave(t["titulo"]), *regra):
            return {"titulo": t["titulo"], "linha_titulo": t["linha_titulo"], "base_total": t.get("base_total"),
                    "pergunta": t.get("pergunta", "")}
    return None


# ---------------------------------------------------------------------------
# Utilidades de pptx
# ---------------------------------------------------------------------------
def _walk(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from _walk(sh.shapes)


def _paragrafos(slide):
    """Todos os parágrafos de caixas de texto e células de tabela do slide."""
    for sh in _walk(slide.shapes):
        if sh.has_text_frame:
            yield from sh.text_frame.paragraphs
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    yield from cell.text_frame.paragraphs


def _substituir_no_paragrafo(par, antigo: str, novo: str) -> bool:
    runs = par.runs
    texto = "".join(r.text for r in runs)
    inicio = texto.find(antigo)
    if inicio == -1:
        return False
    while inicio != -1:
        fim = inicio + len(antigo)
        pos, primeiro = 0, True
        for r in runs:
            r_ini, r_fim = pos, pos + len(r.text)
            pos = r_fim
            if r_fim <= inicio or r_ini >= fim:
                continue
            a, b = max(inicio, r_ini) - r_ini, min(fim, r_fim) - r_ini
            r.text = r.text[:a] + (novo if primeiro else "") + r.text[b:]
            primeiro = False
        texto = "".join(r.text for r in runs)
        inicio = texto.find(antigo, inicio + len(novo))
    return True


def _substituir_slide(slide, pares: list[tuple[str, str]]):
    for par in _paragrafos(slide):
        for antigo, novo in pares:
            _substituir_no_paragrafo(par, antigo, novo)


def _reescrever_paragrafo(par, segmentos: list[tuple[str, bool]]):
    """Reconstrói o parágrafo com (texto, negrito) copiando a formatação dos runs existentes."""
    runs = par.runs
    if not runs:
        return
    normal = next((r for r in runs if not r.font.bold), runs[0])
    negrito = next((r for r in runs if r.font.bold), runs[0])
    p = par._p
    fim = p.find(qn("a:endParaRPr"))
    modelos = {False: copy.deepcopy(normal._r), True: copy.deepcopy(negrito._r)}
    for r in runs:
        p.remove(r._r)
    for texto, neg in segmentos:
        novo = copy.deepcopy(modelos[neg])
        novo.find(qn("a:t")).text = texto
        rpr = novo.find(qn("a:rPr"))
        if rpr is None:
            rpr = etree.SubElement(novo, qn("a:rPr"))
            novo.insert(0, rpr)
        rpr.set("b", "1" if neg else "0")
        if fim is not None:
            fim.addprevious(novo)
        else:
            p.append(novo)


def _forma(slide, nome: str):
    for sh in _walk(slide.shapes):
        if sh.name == nome:
            return sh
    return None


def _remover_forma(sh):
    sh._element.getparent().remove(sh._element)


# ----- clonagem de slides (com gráficos e imagens) --------------------------
def _copiar_chart_part(prs, chart_part):
    pacote = prs.part.package
    nome = pacote.next_partname("/ppt/charts/chart%d.xml")
    xml = copy.deepcopy(chart_part._element)
    for ed in xml.findall(qn("c:externalData")):
        xml.remove(ed)
    novo = ChartPart(nome, chart_part.content_type, pacote, xml)
    xlsx = chart_part.chart_workbook.xlsx_part
    if xlsx is not None:
        novo.chart_workbook.update_from_xlsx_blob(xlsx.blob)
    return novo


def _clonar_slide(prs, origem):
    novo = prs.slides.add_slide(origem.slide_layout)
    for sh in list(novo.shapes):
        sh._element.getparent().remove(sh._element)
    mapa = {}
    for rid, rel in origem.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            mapa[rid] = novo.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.CHART:
            mapa[rid] = novo.part.relate_to(_copiar_chart_part(prs, rel.target_part), RT.CHART)
        else:
            mapa[rid] = novo.part.relate_to(rel.target_part, rel.reltype)
    arvore_nova = novo.shapes._spTree
    for el in origem.shapes._spTree:
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        arvore_nova.append(copy.deepcopy(el))
    ns_r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for el in arvore_nova.iter():
        for atr, val in list(el.attrib.items()):
            if atr.startswith(ns_r) and val in mapa:
                el.set(atr, mapa[val])
    bg = origem._element.cSld.find(qn("p:bg"))
    if bg is not None and novo._element.cSld.find(qn("p:bg")) is None:
        novo._element.cSld.insert(0, copy.deepcopy(bg))
    return novo


def _aplicar_ordem(prs, slides_em_ordem):
    lista = prs.slides._sldIdLst
    por_part = {id(prs.part.related_part(s.rId)): s for s in lista}
    manter = [por_part[id(sl.part)] for sl in slides_em_ordem]
    for sld_id in list(lista):
        lista.remove(sld_id)
        if sld_id not in manter:
            prs.part.drop_rel(sld_id.rId)
    for sld_id in manter:
        lista.append(sld_id)


# ----- limpeza de gráficos e tabelas ---------------------------------------
_RE_SERIE_ONDA = re.compile(r"^[a-zç]{3}/\d{2}$", re.I)   # "abr/26", "jun/26": rótulo de onda de outro estudo


def _serie_de_outro_estudo(nome: str) -> bool:
    n = nome.strip()
    return bool(_RE_SERIE_ONDA.match(n)) or chave(n) in {"capital", "interior", "total"}


def _limpar_graficos(slide, local: str, trocas: list[tuple[str, str]]):
    """Categorias que não são estrutura (nomes de candidatos de outro estudo) viram '(pendente)'."""
    for sh in _walk(slide.shapes):
        if not (getattr(sh, "has_chart", False) and sh.has_chart):
            continue
        chart = sh.chart
        plot = chart.plots[0]
        cats = [" ".join(str(c).split()) for c in plot.categories]
        novas = []
        for c in cats:
            for a, b in trocas:
                c = c.replace(a, b)
            novas.append(c)
        todas_genericas = all(_generico(c) or "candidato" in chave(c) for c in novas)
        if not todas_genericas:
            novas = [PENDENTE]
        dados = CategoryChartData(number_format="0.0")
        dados.categories = novas
        for serie in plot.series:
            nome = serie.name or ""
            for a, b in trocas:
                nome = nome.replace(a, b)
            if _serie_de_outro_estudo(nome):
                nome = PENDENTE
            dados.add_series(" ".join(nome.split()), [None] * len(novas))
        chart.replace_data(dados)


def _limpar_tabelas(slide, trocas: list[tuple[str, str]]):
    for sh in _walk(slide.shapes):
        if not (getattr(sh, "has_table", False) and sh.has_table):
            continue
        for row in sh.table.rows:
            for cell in row.cells:
                if not cell.text_frame.text.strip():
                    continue
                for par in cell.text_frame.paragraphs:
                    for a, b in trocas:
                        _substituir_no_paragrafo(par, a, b)
                texto = " ".join(cell.text_frame.text.split())
                if _generico(texto) or PENDENTE in texto or "candidato" in chave(texto) and "apoiado" in chave(texto):
                    continue
                if "avaliacao do governo" in chave(texto) or chave(texto).startswith(("aval ", "estadual ", "municipal ", "ipv para")):
                    continue
                par0 = cell.text_frame.paragraphs[0]
                if par0.runs:
                    par0.runs[0].text = PENDENTE
                    for r in par0.runs[1:]:
                        r.text = ""
                    for extra in cell.text_frame.paragraphs[1:]:
                        for r in extra.runs:
                            r.text = ""


def _marcar_pendente(slide, texto: str, pos=None, nome: str = ID_PENDENTE):
    left, top, w, h = pos or (547666, 110000, 7500000, 260000)
    caixa = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(w), Emu(h))
    caixa.name = nome
    tf = caixa.text_frame
    tf.word_wrap = True
    tf.text = texto
    par = tf.paragraphs[0]
    par.alignment = PP_ALIGN.LEFT
    r = par.runs[0]
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)


def _titulo_slide(slide) -> str:
    candidatos = []
    for sh in _walk(slide.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip():
            t = " ".join(sh.text_frame.text.split())
            if len(t) > 6 and not t.startswith(("Projeto", "Base:", "(Em", "Pergunta")):
                candidatos.append((sh.top or 0, t))
    candidatos.sort()
    return candidatos[0][1][:120] if candidatos else ""


def _descrever_formas(slide) -> list[dict]:
    itens = []
    for sh in _walk(slide.shapes):
        if getattr(sh, "has_chart", False) and sh.has_chart:
            itens.append({"forma": sh.name, "tipo": "grafico",
                          "categorias": [" ".join(str(c).split()) for c in sh.chart.plots[0].categories],
                          "series": [s.name for s in sh.chart.plots[0].series]})
        elif getattr(sh, "has_table", False) and sh.has_table:
            itens.append({"forma": sh.name, "tipo": "tabela",
                          "linhas": len(sh.table.rows), "colunas": len(sh.table.columns),
                          "cabecalho": [" ".join(c.text.split())[:40] for c in sh.table.rows[0].cells]})
        elif sh.name.startswith(ID_PENDENTE):
            itens.append({"forma": sh.name, "tipo": "aviso_pendente"})
    return itens


# ---------------------------------------------------------------------------
# Montagem do esqueleto
# ---------------------------------------------------------------------------
_ESTADO_DE = {  # "no estado <prep> <Nome>"
    "RJ": "do", "SP": "de", "MG": "de", "BA": "da", "CE": "do", "PR": "do", "PA": "do", "AM": "do", "MA": "do",
    "MT": "de", "MS": "de", "RS": "do", "RN": "do", "PI": "do", "TO": "do", "AC": "do", "AP": "do", "RO": "de",
    "RR": "de", "DF": "do", "ES": "do", "GO": "de", "AL": "de", "PB": "da", "PE": "de", "SE": "de", "SC": "de",
}


def caminho_molde_padrao() -> str:
    import os
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "molde_informa.pptx")


def _lista_pt(itens: list[str]) -> str:
    if len(itens) <= 1:
        return "".join(itens)
    return ", ".join(itens[:-1]) + " e " + itens[-1]


def _modulo_ativo(m: dict, cfg: ConfigEstudo, extras: set) -> bool:
    if m.get("fixo"):
        return True
    mid = m["id"]
    if m.get("cargo"):
        cargo = m["cargo"]
        if cargo == "Deputado":
            return "Deputado Federal" in cfg.cargos or "Deputado Estadual" in cfg.cargos
        return cargo in cfg.cargos
    if mid == "avaliacao_governos":
        return bool(cfg.governos)
    if mid == "comparativo_governos":
        return len(cfg.governos) >= 2 and mid not in extras_negados(cfg)
    if mid == "analise_perguntas":
        return bool(cfg.analises.get("analise_perguntas")) and mid not in extras_negados(cfg)
    return mid in extras


def extras_negados(cfg: ConfigEstudo) -> set:
    return {x[1:] for x in cfg.modulos_extras if x.startswith("-")}


def modulos_padrao() -> list[str]:
    """Ids opcionais marcados por padrão (para a tela)."""
    return [m["id"] for m in CATALOGO if m.get("default") and not m.get("cargo")
            and not m.get("fixo") and m["id"] not in ("avaliacao_governos", "comparativo_governos", "analise_perguntas")]


def _definir_paragrafos(shape, lista_segmentos: list[list[tuple[str, bool]]]):
    tf = shape.text_frame
    pars = tf.paragraphs
    if not lista_segmentos:
        return
    modelo = copy.deepcopy(pars[0]._p)
    for p in pars[1:]:
        p._p.getparent().remove(p._p)
    ultimo = pars[0]._p
    for _ in range(1, len(lista_segmentos)):
        novo = copy.deepcopy(modelo)
        ultimo.addnext(novo)
        ultimo = novo
    for par, segs in zip(tf.paragraphs, lista_segmentos):
        _reescrever_paragrafo(par, segs)


def _pend(valor, rotulo: str, campos: list, slide_id: str | None = None) -> str:
    if valor in (None, "", 0):
        campos.append(rotulo)
        return PENDENTE
    return str(valor)


def _ajustar_fonte(shape, capacidade: int, extra_por_paragrafo: int = 0, base_pt: float = 12.0, minimo_pt: float = 7.0):
    """Reduz a fonte (≈ 1/√ do excesso de texto) para o texto caber na caixa: `capacidade` = nº de caracteres
    que cabem em `base_pt`. Lista longa de bairros (ex.: Angra dos Reis) estourava o slide."""
    pars = shape.text_frame.paragraphs
    n = len(shape.text_frame.text) + extra_por_paragrafo * len(pars)
    if n <= capacidade:
        return
    pt = max(minimo_pt, math.floor(base_pt * math.sqrt(capacidade / n) * 2) / 2)
    for par in pars:
        for r in par.runs:
            r.font.size = Pt(pt)


def _preencher_capa_e_metodologia(slides: dict, cfg: ConfigEstudo, campos: list, avisos: list, mapa_regioes: bytes | None = None):
    local = cfg.local.strip()
    estado = cfg.estado
    s1 = slides.get(1)
    if s1:
        caixa = _forma(s1, "CaixaDeTexto 1")
        if caixa:
            _definir_paragrafos(caixa, [[(estado, True)], [(local, True)]])
        caixa = _forma(s1, "CaixaDeTexto 2")
        if caixa:
            _definir_paragrafos(caixa, [[(cfg.titulo_capa.strip() or PENDENTE, True)]])
        caixa = _forma(s1, "CaixaDeTexto 9")
        if caixa:
            _definir_paragrafos(caixa, [[(mes_ano_padrao(cfg), True)]])

    s6 = slides.get(6)
    if s6 and _forma(s6, "CaixaDeTexto 7"):
        cargos = [c for c in CARGOS if c in cfg.cargos and c != "Prefeito"]
        rejeicao = ", além da rejeição aos nomes apresentados para o Senado. " if "Senador" in cfg.cargos else ". "
        p0 = [(f"Medir as preferências eleitorais dos moradores de {local} para as eleições de 2026", True)]
        p0 += [(", contemplando as disputas para ", False), (_lista_pt(cargos) or PENDENTE, True), (rejeicao, False)]
        paragrafos = [p0]
        if cfg.objetivo_especifico.strip():
            paragrafos.append([(cfg.objetivo_especifico.strip(), False)])
        elif "Prefeito" in cfg.cargos:
            campos.append("Candidatos de Prefeito (objetivo 2)")
            paragrafos.append([("Avaliar o cenário antecipado da eleição municipal de 2028", True),
                               (", identificando a intenção de voto para prefeito entre ", False), (PENDENTE, True),
                               (" em diferentes configurações de disputa. ", False)])
        else:
            campos.append("Objetivo específico do estudo")
            paragrafos.append([(f"{PENDENTE}: objetivo específico do estudo (informar no briefing). ", False)])
        if cfg.governos:
            ordem = [g for g in ("municipal", "estadual", "federal") if g in cfg.governos]
            nomes = []
            for g in ordem:
                if g == "federal":
                    nomes.append("Lula")
                elif g == "estadual":
                    nomes.append(cfg.governador.strip() or PENDENTE)
                else:
                    nomes.append(cfg.prefeito.strip() or PENDENTE)
            paragrafos.append([(f"Compreender a percepção da população sobre os governos {_lista_pt([GOVERNOS[g].lower() for g in ordem])}", True),
                               (", por meio da avaliação das administrações de ", False), (_lista_pt(nomes), True), (".", False)])
        _definir_paragrafos(_forma(s6, "CaixaDeTexto 7"), paragrafos)
        _ajustar_fonte(_forma(s6, "CaixaDeTexto 7"), capacidade=720, extra_por_paragrafo=60, base_pt=14.0)

    s7 = slides.get(7)
    if s7:
        online = cfg.modo == "online"
        idade = cfg.idade_minima or (18 if online else 16)
        if online:
            caixa = _forma(s7, "Retângulo 5")
            if caixa:
                _definir_paragrafos(caixa, [[
                    ("O estudo utilizou ", False), ("metodologia quantitativa ", True), ("por meio de ", False),
                    ("entrevistas em ambiente digital ", True),
                    ("com sofisticado processo de coleta dos dados e questionário de autopreenchimento, realizado com os residentes de ", False),
                    (local, True), (". ", False), ("A auditoria externa ", True), ("ligou para ", False), ("10%", True),
                    (" dos respondentes, checando algumas informações e comprovando a veracidade da pesquisa.", False)]])
            caixa = _forma(s7, "CaixaDeTexto 9")
            if caixa:
                _definir_paragrafos(caixa, [[
                    ("Os entrevistados foram convidados a participarem da pesquisa pelas principais redes sociais. E, a amostra considerou os moradores de ", False),
                    (local, True), (f", com {idade} anos ou mais", True), (", segmentadas conforme dados censitários de ", False),
                    ("sexo e idade ", True), ("do IBGE 2022 e ", False), ("renda média mensal familiar ", True), ("do IBGE 2010.", False)]])
        else:
            caixa = _forma(s7, "CaixaDeTexto 9")
            if caixa:
                _definir_paragrafos(caixa, [[
                    ("A pesquisa foi realizada com ", False), ("eleitores de ", True), (local, True),
                    (f", no estado {_ESTADO_DE.get(cfg.uf.upper(), 'de')} {estado}", True),
                    (f", com {idade} anos ou mais e com entrevistas segmentadas conforme dados censitários (IBGE). "
                     "O desenho amostral considerou a representatividade dos moradores do município de ", False),
                    (local + " ", True), ("por gênero, idade, renda média mensal domiciliar e regiões da cidade.", False)]])
        caixa = _forma(s7, "CaixaDeTexto 7")
        if caixa:
            periodo = texto_periodo_campo(cfg.campo_inicio, cfg.campo_fim)
            if periodo:
                m = re.match(r"^(entre os dias |no dia )(.+)$", periodo)
                seg_periodo = [("As entrevistas foram realizadas ", False), (m.group(1), False), (m.group(2), True), (". ", False)]
            else:
                campos.append("Período de campo")
                seg_periodo = [("As entrevistas foram realizadas ", False), (f"{PENDENTE} (período de campo)", True), (". ", False)]
            tempo = f"{cfg.tempo_medio} minutos" if cfg.tempo_medio else None
            if not tempo:
                campos.append("Tempo médio de aplicação")
            seg_tempo = [("O tempo médio de aplicação do questionário foi de cerca de ", False),
                         (tempo or PENDENTE, True), (".", False)]
            _definir_paragrafos(caixa, [seg_periodo, seg_tempo])
        caixa = _forma(s7, "CaixaDeTexto 15")
        if caixa:
            base = _milhar(cfg.base) if cfg.base else None
            if not base:
                campos.append("Base (nº de entrevistas)")
            margem = cfg.margem.strip()
            if not margem:
                campos.append("Margem de erro")
            _definir_paragrafos(caixa, [
                [("Foram realizadas ", False), (base or PENDENTE, True), (" entrevistas. ", True)],
                [("O erro esperado para a amostra aplicada é de ", False), (margem or PENDENTE, True),
                 (", para o nível de confiança de ", False), (cfg.nivel_confianca or "95%", True), (".", False)]])

    s8 = slides.get(8)
    if s8:
        mapa = _forma(s8, "Imagem 3")
        if mapa:
            pos = (mapa.left, mapa.top, mapa.width, 600000)
            caixa_img = (mapa.left, mapa.top, mapa.width, mapa.height)
            _remover_forma(mapa)
            if mapa_regioes and _inserir_imagem_ajustada(s8, mapa_regioes, caixa_img, "Mapa das regiões"):
                avisos.append("Mapa das regiões (slide 8) vem da capa do Relatório Descritivo — conferir se corresponde ao estudo.")
            else:
                campos.append("Imagem do mapa das regiões")
                _marcar_pendente(s8, "IMAGEM PENDENTE – mapa das regiões do estudo", pos)
        n_reg = f"{cfg.n_regioes:02d}" if cfg.n_regioes else None
        if not n_reg:
            campos.append("Número de regiões")
        caixa = _forma(s8, "CaixaDeTexto 22")
        if caixa:
            if cfg.modo == "online":
                seg = [("Os questionários foram impulsionados em redes sociais para moradores das ", False),
                       (f"{n_reg or PENDENTE} regiões de {local}", True),
                       (", contemplando na amostra respondentes de ", False), (f"{PENDENTE} cidades. ", True)]
                campos.append("Número de cidades (online)")
            else:
                seg = [("A amostra foi desenhada para representar o munícipio de ", False), (local + " ", True),
                       ("e com esse intuito, a cidade foi dividida em ", False), (n_reg or PENDENTE, True), (" regiões. ", True),
                       ("As entrevistas foram distribuídas nos bairros, entre as regiões de ", False), (local, True),
                       (", de acordo com os dados do IBGE/2022 de gênero e idade e, IBGE/2010 de renda média domiciliar dos munícipes da cidade.", False)]
            _definir_paragrafos(caixa, [seg])
        caixa = _forma(s8, "CaixaDeTexto 6")
        if caixa:
            linhas = [l.strip() for l in cfg.regioes.splitlines() if l.strip()]
            if linhas:
                pars = []
                for l in linhas:
                    if ":" in l:
                        a, b = l.split(":", 1)
                        pars.append([(a + ": ", True), (b.strip(), False)])
                    else:
                        pars.append([(l, False)])
                _definir_paragrafos(caixa, pars)
                _ajustar_fonte(caixa, capacidade=720, extra_por_paragrafo=80)
            else:
                campos.append("Regiões e pesos (bairros por região)")
                _definir_paragrafos(caixa, [[(f"{PENDENTE}: regiões, pesos e bairros do desenho amostral.", False)]])


def _inserir_imagem_ajustada(slide, dados: bytes, caixa, nome: str) -> bool:
    """Insere a imagem dentro da caixa (esq, topo, larg, alt) mantendo a proporção, centralizada. False se não abrir."""
    try:
        from PIL import Image
        with Image.open(io.BytesIO(dados)) as im:
            w, h = im.size
    except Exception:
        return False
    if not w or not h:
        return False
    esq, topo, larg, alt = caixa
    escala = min(larg / w, alt / h)
    lw, la = int(w * escala), int(h * escala)
    pic = slide.shapes.add_picture(io.BytesIO(dados), Emu(esq + (larg - lw) // 2), Emu(topo + (alt - la) // 2), Emu(lw), Emu(la))
    pic.name = nome
    return True


_LEGENDA_CAIXA = (547666, 5772139, 11100000, 590000)   # faixa do rodapé dos slides de mapa de calor (acima do rodapé)


def _fonte_legenda(pars: list[str], larg_emu: int, alt_emu: int, colunas: int = 2) -> tuple[float, bool]:
    """Maior fonte (8 → 5,5 pt) em que a legenda cabe na faixa, em `colunas` colunas. Devolve (pt, coube)."""
    larg_col = (larg_emu / 914400) / colunas - 0.1
    alt = alt_emu / 914400
    for pt in (8, 7.5, 7, 6.5, 6, 5.5):
        por_linha = max(10, larg_col / (pt * 0.0076))            # largura média de um caractere ≈ 0,52 × corpo
        linhas = sum(max(1, math.ceil(len(t) / por_linha)) for t in pars)
        if linhas / colunas * pt * 1.22 / 72 <= alt:
            return pt, True
    return 5.5, False


def _legenda_regioes(slide, regioes: str, avisos: list | None = None) -> bool:
    """Escreve a legenda «Região N (xx%): bairros» na faixa do rodapé (no lugar da imagem do molde)."""
    linhas = [l.strip() for l in regioes.splitlines() if l.strip()]
    if not linhas:
        return False
    esq, topo, larg, alt = _LEGENDA_CAIXA
    caixa = slide.shapes.add_textbox(Emu(esq), Emu(topo), Emu(larg), Emu(alt))
    caixa.name = "Legenda das regiões"
    tf = caixa.text_frame
    tf.word_wrap = True
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    bp = tf._txBody.find(qn("a:bodyPr"))
    bp.set("numCol", "2")
    bp.set("spcCol", "182880")
    pt, coube = _fonte_legenda(linhas, larg, alt)
    for i, l in enumerate(linhas):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        cab, sep, resto = l.partition(":")
        partes = [(cab + (": " if sep else ""), True), (resto.strip(), False)] if sep else [(l, False)]
        for texto, negrito in partes:
            if not texto:
                continue
            r = par.add_run()
            r.text = texto
            r.font.size = Pt(pt)
            r.font.bold = negrito
            r.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
    if not coube and avisos is not None and "Legenda das regiões" not in " ".join(avisos):
        avisos.append("Legenda das regiões (slides de mapa de calor) muito longa: fonte mínima de 5,5 pt — conferir o encaixe.")
    return True


def _remover_imagem_regioes(slide, regioes: str = "", avisos: list | None = None, campos: list | None = None):
    """O molde traz uma imagem (texto rasterizado) com bairros/regiões de outro estudo no rodapé dos slides de mapa.
    Troca pela legenda lida da capa do Descritivo; sem ela, deixa o aviso de pendência."""
    for sh in list(slide.shapes):
        if sh.shape_type == 13 and sh.left == 0 and sh.top == 5772139 and sh.width == 12192000 and sh.height == 577473:
            _remover_forma(sh)
            if _legenda_regioes(slide, regioes, avisos):
                continue
            if campos is not None:
                campos.append("Legenda das regiões (slides de mapa de calor)")
            _marcar_pendente(slide, "IMAGEM PENDENTE – legenda de bairros por região (fonte: capa do Relatório Descritivo)",
                             (547666, 5800000, 9000000, 260000), nome=ID_PENDENTE + "_REGIOES")


def _definir_base(slide, valor: str):
    for par in _paragrafos(slide):
        t = par.text.strip()
        if t.startswith("Base:") and "entrevistas" in t and par.runs:
            _reescrever_paragrafo(par, [(f"Base: {valor} entrevistas", False)])


def _bio_candidato(slide, cfg: ConfigEstudo, campos: list):
    for sh in _walk(slide.shapes):
        if sh.has_text_frame and "Esse é" in sh.text_frame.text and "{{CANDIDATO}}" in sh.text_frame.text:
            texto = cfg.apresentacao_candidato.strip()
            if not texto:
                campos.append("Texto de apresentação do candidato (nível 2)")
                texto = f"{PENDENTE}: texto de apresentação do candidato (vem do briefing)."
            _definir_paragrafos(sh, [[(texto, False)]])


def montar_esqueleto(molde, cfg: ConfigEstudo, inv: dict | None = None, achados: list | None = None,
                     mapa_regioes: bytes | None = None):
    """Gera (bytes do .pptx, manifesto dict). `molde`: caminho, bytes ou None (usa assets/)."""
    if not cfg.local.strip():
        raise ValueError("Informe o local do estudo.")
    if molde is None:
        molde = caminho_molde_padrao()
    prs = Presentation(io.BytesIO(molde) if isinstance(molde, (bytes, bytearray)) else molde)
    originais = list(prs.slides)
    campos: list[str] = []
    avisos: list[str] = []
    extras = {x for x in cfg.modulos_extras if not x.startswith("-")}

    usados: set[int] = set()

    def pegar(idx: int):
        if idx not in usados:
            usados.add(idx)
            return originais[idx - 1]
        return _clonar_slide(prs, originais[idx - 1])

    plano: list[dict] = []
    contagem: Counter = Counter()

    def adicionar(slide, idx, mod_id, papel=None, fonte=None, cargo=None, subst=None, tabelas_tambem=True, requer=None):
        contagem[mod_id] += 1
        plano.append({"slide": slide, "idx": idx, "modulo": mod_id, "papel": papel, "fonte": fonte, "cargo": cargo,
                      "subst": subst or [], "tabelas": tabelas_tambem, "id": f"{mod_id}-{contagem[mod_id]:02d}"})

    for m in CATALOGO:
        if not _modulo_ativo(m, cfg, extras):
            continue
        mid = m["id"]
        if m.get("clona"):  # Governador = clone do módulo Presidente
            base_mod = CATALOGO_POR_ID[m["clona"]]
            for s in base_mod["slides"]:
                if s["requer"] and not cfg.analises.get(s["requer"]):
                    continue
                slide = pegar(s["idx"]) if False else _clonar_slide(prs, originais[s["idx"] - 1])
                adicionar(slide, s["idx"], mid, s["papel"], s["fonte"], "Governador",
                          subst=[("Presidente da República", "Governador"), ("Presidente", "Governador")], tabelas_tambem=False)
            continue
        if mid == "analises_estrategicas" and not (cfg.analises.get("correlacao") or cfg.analises.get("residuais")):
            continue
        if mid == "metodologia" and cfg.modo == "online":
            itens = [s for s in m["slides"] if s["idx"] != 9]
        else:
            itens = m["slides"]
        for s in itens:
            if s["requer"] and not cfg.analises.get(s["requer"]):
                continue
            if m.get("cargo") == "Deputado" and s["idx"] == 25 and False:
                continue
            adicionar(pegar(s["idx"]), s["idx"], mid, s["papel"], s["fonte"], m.get("cargo"), subst=s.get("subst"))
        if mid == "avaliacao_governos":
            if cfg.analises.get("mapas_calor"):
                for g in [g for g in ("federal", "estadual", "municipal") if g in cfg.governos]:
                    slide = pegar(SLIDE_AVALIACAO)
                    nome = {"federal": "Federal Lula",
                            "estadual": f"Estadual {cfg.governador.strip() or PENDENTE}",
                            "municipal": f"Municipal {cfg.prefeito.strip() or PENDENTE}"}[g]
                    if g == "estadual" and not cfg.governador.strip():
                        campos.append("Nome do governador")
                    if g == "municipal" and not cfg.prefeito.strip():
                        campos.append("Nome do prefeito")
                    adicionar(slide, SLIDE_AVALIACAO, mid, f"avaliacao_{g}", "descritivo",
                              subst=[("Federal Lula", nome)] if g != "federal" else [])
            else:
                if "federal" in cfg.governos and "estadual" in cfg.governos:
                    adicionar(pegar(SLIDE_AVALIACAO_SEM_MAPAS), SLIDE_AVALIACAO_SEM_MAPAS, mid, "avaliacao_lado_a_lado", "descritivo")
                    if not cfg.governador.strip():
                        campos.append("Nome do governador")
                else:
                    avisos.append("Sem 'Mapas de calor' o layout de avaliação só cobre Federal + Estadual lado a lado; "
                                  "marque Mapas de calor para gerar um slide por governo.")
                if "municipal" in cfg.governos:
                    avisos.append("Governo municipal sem mapas de calor: layout lado a lado não cobre este governo — adicionar manualmente.")

    # encerramento sempre por último (já é o último do catálogo)
    # --- remove o que não foi planejado e aplica a ordem
    _aplicar_ordem(prs, [e["slide"] for e in plano])

    # --- preenchimentos específicos (capa, metodologia, bases, bio)
    por_idx = {}
    for e in plano:
        por_idx.setdefault(e["idx"], e["slide"])  # primeira ocorrência (original) para capa/metodologia
    _preencher_capa_e_metodologia(por_idx, cfg, campos, avisos, mapa_regioes)

    local = cfg.local.strip()
    gov = cfg.governador.strip() or PENDENTE
    pref = cfg.prefeito.strip() or PENDENTE
    candidato = cfg.candidato_foco.strip()
    gestao = cfg.gestao_anterior.strip()
    cav_ativos = any(e["modulo"] in ("candidato_gestao", "candidato_perfil_ipv", "cav", "cav_enxuto") for e in plano)
    if cav_ativos:
        if not candidato:
            campos.append("Candidato foco")
        if any(e["idx"] in (33, 47) for e in plano) and not gestao:
            campos.append("Gestão anterior")
    cand_txt = candidato or PENDENTE
    gestao_txt = gestao or PENDENTE

    trocas_globais = [
        ("Governo Municipal Osasco", f"Governo Municipal {pref}"), ("Gestão Gerson Pessoa", f"Gestão {gestao_txt}"),
        ("Gestão Prefeito Rogério Lins", f"Gestão {cand_txt}"), ("Dr. Julinho Juju", pref),
        ("Projeto {{LOCAL}}, {{UF}}", cfg.rodape), ("Projeto Paty do Alferes, Rio de Janeiro", cfg.rodape),
        ("Projeto Alphaville, {{LOCAL}}", cfg.rodape),
        ("{{CANDIDATO}}", cand_txt), ("{{GESTAO_ANTERIOR}}", gestao_txt), ("{{LOCAL}}", local), ("{{UF}}", cfg.estado),
        ("Estadual Desembargador Ricardo Couto", f"Estadual {gov}"), ("Estadual do Desembargador Ricardo Couto", f"Estadual do {gov}"),
        ("Desembargador Ricardo Couto", gov), ("Municipal Dr. Julinho Juju", f"Municipal {pref}"),
        ("Governo Municipal de Dr. Julinho Juju", f"Governo Municipal de {pref}"), ("Tarcísio de Freitas", gov),
    ] + [(o, local) for o in OUTROS_LOCAIS]
    trocas_tabela = [(a, b) for a, b in trocas_globais]
    trocas_grafico = [t for t in trocas_globais if not t[0].startswith("Projeto")]
    corrigidas = Counter()

    for e in plano:
        sl, idx = e["slide"], e["idx"]
        _bio_candidato(sl, cfg, campos)
        pares = list(e["subst"])
        if e["modulo"] in ("candidato_gestao", "candidato_perfil_ipv", "cav") and cfg.cargo_foco != "Deputado Estadual":
            pares += [("Deputado Estadual", cfg.cargo_foco), ("deputado estadual", cfg.cargo_foco.lower())]
        if e["modulo"] == "cav_enxuto" and cfg.cargo_foco != "Deputado Federal":
            pares += [("Deputado Federal", cfg.cargo_foco), ("deputado federal", cfg.cargo_foco.lower())]
        if e["subst"] and not e["tabelas"]:
            for sh in _walk(sl.shapes):
                if sh.has_text_frame:
                    for par in sh.text_frame.paragraphs:
                        for a, b in pares:
                            _substituir_no_paragrafo(par, a, b)
        else:
            _substituir_slide(sl, pares)
        _substituir_slide(sl, trocas_globais)
        for antigo, novo in CORRECOES_MOLDE:
            for par in _paragrafos(sl):
                if _substituir_no_paragrafo(par, antigo, novo):
                    corrigidas[(antigo, novo)] += 1
        _limpar_graficos(sl, local, trocas_grafico)
        _remover_imagem_regioes(sl, cfg.regioes, avisos, campos)
        _limpar_tabelas(sl, [])
        if idx == 12:
            _definir_base(sl, _milhar(cfg.base) if cfg.base else "__")
        elif idx not in (1, 6, 7, 8):
            _definir_base(sl, "__")
        if e["fonte"]:
            _marcar_pendente(sl, f"CONTEÚDO PENDENTE – fonte: {FONTES[e['fonte']]}")

    # --- notas com ID e manifesto
    slides_manifesto = []
    for pos, e in enumerate(plano, 1):
        sl = e["slide"]
        sl.notes_slide.notes_text_frame.text = f"INFORMA-ID: {e['id']} | módulo: {e['modulo']} | papel: {e['papel'] or '-'}"
        cargo = e["cargo"] if e["cargo"] not in (None, "Deputado") else None
        sug = _sugerir_pergunta(inv, cargo, e["papel"])
        if e["cargo"] == "Deputado" and inv:
            lista = []
            for c in ("Deputado Federal", "Deputado Estadual"):
                pal = _PALAVRAS_CARGO.get(c, chave(c))
                for t in inv.get("totais", []):
                    k_ = chave(t["titulo"])
                    if _titulo_casa(k_, ["intencao", pal], [], _INTENCAO_EXCLUI):
                        lista.append({"cargo": c, "titulo": t["titulo"], "linha_titulo": t["linha_titulo"],
                                      "base_total": t.get("base_total"), "pergunta": t.get("pergunta", "")})
            deputados_lista = lista
        item = {
            "posicao": pos, "id": e["id"], "modulo": e["modulo"], "papel": e["papel"], "slide_no_molde": e["idx"],
            "titulo": _titulo_slide(sl),
            "status": "pendente" if e["fonte"] else "pronto",
            "fonte_esperada": FONTES[e["fonte"]] if e["fonte"] else None,
            "formas": _descrever_formas(sl),
        }
        if sug:
            item["pergunta_sugerida_no_descritivo"] = sug
        if e["cargo"] == "Deputado" and inv and deputados_lista and e["papel"] in ("intencao", "cruzamento"):
            item["perguntas_candidatas_no_descritivo"] = deputados_lista
        slides_manifesto.append(item)

    # --- salva e verifica
    saida = io.BytesIO()
    prs.save(saida)
    dados = saida.getvalue()
    residuos = verificar_residuos(dados, [cfg.local, cfg.governador, cfg.prefeito, cfg.candidato_foco, cfg.gestao_anterior])

    if inv and inv.get("tem_secoes") is not None:
        faltando = [c for c in cfg.cargos if c in _PALAVRAS_CARGO and c not in detectar_cargos(inv)]
        if faltando:
            avisos.append("Cargos selecionados sem pergunta de intenção de voto no Descritivo: " + ", ".join(faltando) + ".")
    avisos.append("Slide de perfil: conferir a nota do salário mínimo (texto fixo do molde).")
    if cfg.modo == "rua" and not cfg.regioes.strip():
        avisos.append("Estudo de rua sem regiões informadas: slide de abrangência fica com texto pendente.")

    manifesto = {
        "versao": VERSAO_MANIFESTO,
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "estudo": {"local": local, "uf": cfg.uf, "estado": cfg.estado, "modo": cfg.modo, "mes_ano": mes_ano_padrao(cfg)},
        "configuracao": cfg.para_dict(),
        "campos_pendentes": sorted(set(campos), key=campos.index),
        "avisos": avisos,
        "correcoes_molde": [f"'{a}' → '{b}' ({n}x)" for (a, b), n in corrigidas.items()],
        "residuos_de_outros_estudos": residuos,
        "slides": slides_manifesto,
        "verificacao_descritivo": None,
        "instrucoes_para_preenchimento": INSTRUCOES_IA,
    }
    if inv is not None:
        manifesto["verificacao_descritivo"] = {
            "tabelas_totais": len(inv.get("totais", [])),
            "tabelas_segmentos": len(inv.get("segmentos", [])),
            "divergencias": [{k: a[k] for k in ("tipo", "titulo_total", "titulo_segmento", "detalhe", "linhas")}
                             for a in (achados or [])],
            "perguntas_nos_totais": [{"titulo": t["titulo"], "linha": t["linha_titulo"], "base": t.get("base_total"),
                                      "pergunta": t.get("pergunta", "")} for t in inv.get("totais", [])],
        }
    return dados, manifesto


INSTRUCOES_IA = [
    "Para cada slide com status 'pendente': leia o arquivo indicado em 'fonte_esperada', preencha os gráficos "
    "(chart.replace_data) e as tabelas listadas em 'formas' e, ao terminar, remova as formas cujo nome começa com 'INFORMA_PENDENTE'.",
    "Se o arquivo-fonte não foi enviado, mantenha o slide e troque o texto do aviso por "
    "'Arquivo de <fonte> não enviado – conteúdo pendente'. Nunca remova o slide nem estime valores.",
    "Textos '(pendente)' em campos do slide: preencher com dados do Descritivo/briefing; se não houver, manter e listar em Pendências.",
    "Gráficos e tabelas foram esvaziados de propósito (nomes de candidatos de outro estudo foram removidos): "
    "monte categorias e linhas a partir dos dados do estudo, casando rótulos por texto normalizado.",
    "Base: o rodapé 'Base: __ entrevistas' deve receber a base da pergunta do slide (veja 'pergunta_sugerida_no_descritivo' "
    "e 'perguntas_nos_totais'); só o slide de perfil já traz a base total.",
    "Use as divergências de 'verificacao_descritivo' para escolher a fonte correta e registre as que afetarem o slide na conferência.",
    "Slides com papel 'analise_perguntas' (fonte: Análise das perguntas, crosstabs do SPSS): o slide é um modelo — clone-o por pergunta "
    "(título da tabela '<Pergunta> * <Segmento> Crosstabulation'; casar a pergunta com a tabela do Descritivo para título, pergunta do rodapé e base). "
    "Monte tabelas de % por segmento (sexo, faixa etária, renda, escolaridade, estado civil | religiões, regiões) com o resíduo ajustado na 2ª linha da célula: "
    "uma pergunta com até 6 alternativas cabe em 1 slide com 2 tabelas; com mais alternativas use 2 slides (um por grupo de segmentos) e, acima de 10 alternativas, "
    "mostre só o % (resíduo apenas como destaque). Destaque Adjusted Residual ≥ +1,96 (azul) e ≤ −1,96 (laranja). Os valores do arquivo são não ponderados: "
    "avise isso no slide e não os misture com o Descritivo ponderado. Sem o arquivo, mantenha o aviso 'Arquivo de Análise das perguntas não enviado – conteúdo pendente'.",
    "Gráficos de barras: ao preencher, defina c:tickLblSkip=1 no eixo de categorias (senão o PowerPoint omite rótulos alternados), "
    "quebre rótulos longos antes de '(' e reduza a fonte do eixo (≈9 pt) quando houver 11+ categorias.",
    "Rodapé: a pergunta deve ficar em uma linha (fonte menor se necessário) sem encostar na linha 'Projeto…'.",
    "Regras de dados: nunca inventar valor, nunca usar zero como preenchimento, nunca deixar dado de outro estudo; "
    "resultado ponderado é a referência principal.",
]


def verificar_residuos(pptx_bytes: bytes, permitidos: list[str] | None = None) -> list[dict]:
    """Procura nomes/termos de outros estudos e marcadores soltos no pptx gerado.
    `permitidos`: nomes do próprio estudo (local, governador...) que não contam como resíduo."""
    perm = " ".join(chave(x) for x in (permitidos or []) if x)
    prs = Presentation(io.BytesIO(pptx_bytes))
    achados = []
    for pos, sl in enumerate(prs.slides, 1):
        textos = [p.text for p in _paragrafos(sl)]
        for sh in _walk(sl.shapes):
            if getattr(sh, "has_chart", False) and sh.has_chart:
                textos += [str(c) for c in sh.chart.plots[0].categories]
                textos += [str(x.name) for x in sh.chart.plots[0].series]
                for x in sh.chart.plots[0].series:
                    if _serie_de_outro_estudo(str(x.name)):
                        achados.append({"posicao": pos, "termo": f"série '{x.name}'"})
        todo = chave(" ".join(textos))
        bruto = " ".join(textos)
        for termo in TERMOS_RESIDUAIS:
            if termo in ("{{", "}}"):
                if termo in bruto:
                    achados.append({"posicao": pos, "termo": termo})
            elif chave(termo) and chave(termo) in todo and chave(termo) not in perm:
                achados.append({"posicao": pos, "termo": termo})
    return achados
