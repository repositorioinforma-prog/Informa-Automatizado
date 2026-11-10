"""
Padronização de Títulos — catálogo dos títulos padrão por tipo de pergunta (lógica pura, sem Streamlit).

Levantado a partir dos Relatórios Descritivos de 6 projetos (Alphaville, Angra dos Reis, Nilópolis, Osasco,
São Paulo Online e Rio de Janeiro Benedita 2 — 8 arquivos, ~150 títulos únicos) e do Total do RJ Deputados.

Cada padrão descreve UM tipo de pergunta e traz:
    * o modelo do título, com campos entre chaves — «Intenção de voto para {cargo}/2026 - Cenário Único»;
    * o modelo da pergunta (texto do `CAPTION 'Pergunta: ...'` do .sps);
    * o sufixo do tipo de resposta — «(Estimulada e Única)», «(Espontânea e Única)», «(Estimulada e Múltipla)» —,
      que `montar_titulo` acrescenta sozinho;
    * as outras formas que o mesmo título já apareceu nos relatórios (`variantes`) e outras redações da pergunta
      (`perguntas_vistas`), para a pessoa reconhecer o que está procurando.

Onde os relatórios divergiram, valeu a forma mais frequente: «/2026» (e não «/26»), «Cenário Único», «Índice de
Probabilidade de Voto (IPV)», «Rejeição aos candidatos para {cargo}» e «Presidente da República». Nos empates
valeu a forma mais explícita; a outra fica em `variantes`.

Por ora o módulo é só de CONSULTA: dado o tipo de pergunta, devolve o título padrão (e a pergunta). A aplicação
automática nos arquivos fica para uma etapa seguinte.

Convenção do projeto: este arquivo só tem lógica pura (sem `import streamlit`); a tela fica em
`analises/padronizacao_titulos.py`.
"""
from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from string import Formatter
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

# ---------------------------------------------------------------- tipos de resposta (sufixo do título)
ESTIMULADA = "Estimulada"
ESPONTANEA = "Espontânea"
UNICA = "Única"
MULTIPLA = "Múltipla"

# ---------------------------------------------------------------- grupos
G_VOTO = "Intenção de voto"
G_REJEICAO = "Rejeição"
G_IPV = "IPV, chance de voto e perfil"
G_CONHECIMENTO = "Conhecimento e imagem do candidato"
G_AVALIACAO = "Avaliação e aprovação"
G_POSICIONAMENTO = "Posicionamento político"
G_MIDIA = "Hábitos de mídia e comunicação"
G_OUTROS = "Outras tabelas"
GRUPOS: Tuple[str, ...] = (
    G_VOTO, G_REJEICAO, G_IPV, G_CONHECIMENTO, G_AVALIACAO, G_POSICIONAMENTO, G_MIDIA, G_OUTROS,
)

CARGOS: Tuple[str, ...] = (
    "Presidente da República", "Governador", "Senador", "Deputado Federal", "Deputado Estadual", "Senadora",
)
CARGOS_DEPUTADO: Tuple[str, ...] = ("Deputado Federal", "Deputado Estadual")
ANO_PADRAO = "2026"


# ---------------------------------------------------------------- estrutura
@dataclass(frozen=True)
class Campo:
    """Parte variável do título/pergunta (o que muda de projeto para projeto)."""
    nome: str
    rotulo: str
    exemplo: str
    opcoes: Tuple[str, ...] = ()      # quando há uma lista fechada usual (a pessoa ainda pode digitar outra)


@dataclass(frozen=True)
class PadraoTitulo:
    id: str
    grupo: str
    nome: str                          # como a pessoa reconhece o tipo de pergunta
    modelo_titulo: str                 # com {campos}; SEM o sufixo do tipo de resposta
    modelo_pergunta: str               # texto da pergunta, com {campos}
    campos: Tuple[Campo, ...] = ()
    estimulada: Optional[str] = ESTIMULADA   # ESTIMULADA | ESPONTANEA | None (sem sufixo)
    resposta: Optional[str] = UNICA          # UNICA | MULTIPLA | None
    palavras_chave: Tuple[str, ...] = ()
    observacao: str = ""
    variantes: Tuple[Tuple[str, str], ...] = ()          # (onde apareceu, título como estava)
    perguntas_vistas: Tuple[Tuple[str, str], ...] = ()   # (onde apareceu, outra redação da pergunta)

    @property
    def sufixo(self) -> str:
        if not self.estimulada or not self.resposta:
            return ""
        return f"({self.estimulada} e {self.resposta})"

    @property
    def nomes_campos(self) -> Tuple[str, ...]:
        return tuple(c.nome for c in self.campos)


_C_CARGO = Campo("cargo", "Cargo", "Governador", CARGOS)
_C_CARGO_DEP = Campo("cargo", "Cargo", "Deputado Estadual", CARGOS_DEPUTADO)
_C_CAND = Campo("candidato", "Candidato", "Rogério Lins")
_C_CIDADE = Campo("cidade", "Cidade", "Osasco")
_C_PREFEITO = Campo("prefeito", "Prefeito", "Gerson Pessoa")


# ---------------------------------------------------------------- catálogo
PADROES: Tuple[PadraoTitulo, ...] = (
    # ============================================================ intenção de voto
    PadraoTitulo(
        id="voto_espontanea", grupo=G_VOTO, nome="Voto espontâneo",
        modelo_titulo="Intenção de voto para {cargo}/2026",
        modelo_pergunta="Se as eleições fossem hoje, em quem votaria para {cargo}?",
        campos=(_C_CARGO,), estimulada=ESPONTANEA, resposta=UNICA,
        palavras_chave=("espontanea", "espontâneo", "sem lista", "primeira pergunta"),
        variantes=(
            ("Alphaville", "Intenção de voto para Presidente da República/26 (Espontânea e Única)"),
            ("Angra dos Reis", "Intenção de voto para Presidente/2026 (Espontânea e Única)"),
            ("Nilópolis", "Intenção de voto para Deputado Estadual/2026 - Cenário Único (Espontânea e Única)"),
        ),
        perguntas_vistas=(
            ("RJ Deputados", "Se as eleições para Deputado Estadual fossem hoje, em quem você votaria?"),
            ("Alphaville", "Se as eleições fossem hoje, em quem votaria para Governador do Estado de São Paulo?"),
        ),
    ),
    PadraoTitulo(
        id="voto_cenario_unico", grupo=G_VOTO, nome="Voto estimulado — Cenário Único",
        modelo_titulo="Intenção de voto para {cargo}/2026 - Cenário Único",
        modelo_pergunta="Sendo estes os candidatos para {cargo}, em qual votaria?",
        campos=(_C_CARGO,), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("estimulada", "estimulado", "cenario unico", "lista de candidatos"),
        observacao="Para Deputado Federal/Estadual a pergunta costuma levar «da sua cidade» (ver outras redações).",
        variantes=(
            ("Alphaville", "Intenção de voto para Governador/26 - Cenário único (Estimulada e Única)"),
            ("Angra dos Reis", "Intenção de voto para Presidente/2026 - Cenário Único (Estimulada e Única)"),
            ("Nilópolis", "Intenção de voto para Deputado Federal/26 (Estimulada e Única)"),
        ),
        perguntas_vistas=(
            ("Osasco", "Sendo estes os candidatos da sua cidade para Deputado Federal, em qual votaria?"),
            ("Angra dos Reis", "Se as eleições para Governador fossem hoje e os candidatos fossem esses, em qual votaria?"),
        ),
    ),
    PadraoTitulo(
        id="voto_cenario_numerado", grupo=G_VOTO, nome="Voto estimulado — Cenário 1, 2, ...",
        modelo_titulo="Intenção de voto para {cargo}/2026 - Cenário {numero}",
        modelo_pergunta="Se as eleições para {cargo} fossem hoje e alguns dos candidatos fossem esses, em qual deles você votaria?",
        campos=(_C_CARGO_DEP, Campo("numero", "Número do cenário", "1", ("1", "2", "3"))),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("cenario 1", "cenario 2", "cenarios", "estimulada"),
        variantes=(("Angra dos Reis", "Intenção de voto para Deputado Estadual/2026 - Cenário 1 (Estimulada e Única)"),),
        perguntas_vistas=(
            ("Angra dos Reis", "Sendo esses os candidatos a Deputado Estadual de Angra, em qual deles você votaria?"),
        ),
    ),
    PadraoTitulo(
        id="voto_transferencia", grupo=G_VOTO, nome="Voto estimulado — Transferência de voto",
        modelo_titulo="Intenção de voto para {cargo}/2026 - Transferência de voto sem {candidato}",
        modelo_pergunta="Se {candidato} não for candidato e se as eleições para {cargo} fossem hoje e alguns dos candidatos fossem esses, em quem você votaria?",
        campos=(_C_CARGO_DEP, Campo("candidato", "Candidato que sai", "Greg Duarte")),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("transferencia", "sem o candidato", "candidato fora"),
        variantes=(("Angra dos Reis", "Intenção de voto para Deputado Estadual/2026 - Transferência de voto sem Greg Duarte (Estimulada e Única)"),),
    ),
    PadraoTitulo(
        id="voto_senador_primeiro", grupo=G_VOTO, nome="Senador — Primeiro Candidato",
        modelo_titulo="Intenção de voto para Senador/2026 - Primeiro Candidato - Cenário Único",
        modelo_pergunta="Para senador, votamos em 2 nomes. Qual o primeiro candidato que você votaria para senador?",
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("senador", "1º voto", "primeiro voto", "primeira opcao"),
        variantes=(
            ("Osasco", "Intenção de voto para Senador/2026 - Primeiro candidato - Cenário Único (Estimulada e Única)"),
            ("Alphaville", "Primeira opção de voto para Senador por São Paulo/26 - Cenário único (Estimulada e Única)"),
        ),
        perguntas_vistas=(("Nilópolis", "Neste cenário, qual seria o primeiro candidato que você votaria para senador?"),),
    ),
    PadraoTitulo(
        id="voto_senador_segundo", grupo=G_VOTO, nome="Senador — Segundo Candidato",
        modelo_titulo="Intenção de voto para Senador/2026 - Segundo Candidato - Cenário Único",
        modelo_pergunta="E qual seria o segundo candidato que você votaria para senador?",
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("senador", "2º voto", "segundo voto", "segunda opcao"),
        variantes=(
            ("Osasco", "Intenção de voto para Senador/2026 - Segundo candidato - Cenário Único (Estimulada e Única)"),
            ("Alphaville", "Segunda opção de voto para Senador por São Paulo/26 - Cenário único (Estimulada e Única)"),
        ),
        perguntas_vistas=(("Angra dos Reis", "E, qual seria o segundo candidato que você votaria para senador?"),),
    ),
    PadraoTitulo(
        id="voto_senador_base100", grupo=G_VOTO, nome="Senador — Base 100% (1º + 2º Voto)",
        modelo_titulo="Intenção de voto para Senador/2026 - Base 100% (1º + 2º Voto)",
        modelo_pergunta="Considerando o primeiro e o segundo voto para Senador, em quem você votaria?",
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("senador", "base 100", "1º + 2º", "soma dos votos"),
        observacao="Em alguns relatórios esta tabela aparece só como «Candidatos»; o padrão é o título completo.",
        variantes=(("vários", "Candidatos"),),
    ),
    PadraoTitulo(
        id="voto_segundo_turno", grupo=G_VOTO, nome="Voto no 2º Turno",
        modelo_titulo="Intenção de voto para {cargo}/2026 no 2º Turno",
        modelo_pergunta="Sendo estes os candidatos para {cargo} no segundo turno das eleições, em qual votaria?",
        campos=(Campo("cargo", "Cargo", "Presidente da República", ("Presidente da República", "Governador")),),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("segundo turno", "2o turno", "turno"),
        variantes=(("Benedita 2", "Intenção de voto para Presidente/2026 no 2º Turno (Estimulada e Única)"),),
    ),

    # ============================================================ rejeição
    PadraoTitulo(
        id="rejeicao", grupo=G_REJEICAO, nome="Rejeição aos candidatos",
        modelo_titulo="Rejeição aos candidatos para {cargo}",
        modelo_pergunta="E se os candidatos a {cargo} fossem esses em quem você NÃO votaria de jeito nenhum?",
        campos=(_C_CARGO,), estimulada=ESTIMULADA, resposta=MULTIPLA,
        palavras_chave=("rejeicao", "nao votaria", "de jeito nenhum", "rejeita"),
        variantes=(
            ("Alphaville", "Rejeição aos candidatos a Presidente da República/26 (Estimulada e Múltipla)"),
            ("Osasco", "Rejeição de voto para Governador/2026 (Estimulada e Múltipla)"),
        ),
        perguntas_vistas=(
            ("Alphaville", "Em quais destes candidatos a Governador você não votaria de jeito nenhum?"),
        ),
    ),

    # ============================================================ IPV, chance, perfil
    PadraoTitulo(
        id="ipv_candidato", grupo=G_IPV, nome="IPV — candidato",
        modelo_titulo="Índice de Probabilidade de Voto (IPV) para {cargo}/2026 - {candidato}",
        modelo_pergunta="Qual a chance de você votar em {candidato} para {cargo}?",
        campos=(_C_CARGO, _C_CAND), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("ipv", "indice de probabilidade", "chance de voto", "probabilidade"),
        observacao="«Senadora» vale quando a candidata é mulher (ex.: Benedita da Silva).",
        variantes=(
            ("São Paulo Online", "Índice de probabilidade de voto (IPV) para Governador/2026 - Tarcísio de Freitas (Estimulada e Única)"),
            ("Alphaville", "Índice de Probabilidade de Voto (IPV) para Deputado Federal/2026 - Wesley Cezar"),
        ),
        perguntas_vistas=(
            ("RJ Deputados", "Qual a chance de você votar no seguinte candidato para Deputado Estadual:"),
            ("Nilópolis", "Para o seguinte candidato a deputado estadual, diga se:"),
            ("São Paulo Online", "E no Fernando Haddad para Governador de São Paulo?"),
        ),
    ),
    PadraoTitulo(
        id="ipv_apoiado_por", grupo=G_IPV, nome="IPV — candidato apoiado por",
        modelo_titulo="Índice de Probabilidade de Voto (IPV) para candidato a {cargo}/2026 apoiado por {apoiador}",
        modelo_pergunta="Qual a chance de você votar em um candidato a {cargo} apoiado por {apoiador}?",
        campos=(_C_CARGO_DEP, Campo("apoiador", "Apoiador", "Lula")),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("ipv", "apoiado", "apoio", "bolsonaro", "lula", "tarcisio"),
        variantes=(("Osasco", "Índice de Probabilidade de Voto (IPV) para candidato a Deputado Estadual/2026 apoiado por Bolsonaro (Estimulada e Única)"),),
        perguntas_vistas=(("Osasco", "E em um candidato a Deputado Estadual apoiado pelo Presidente Lula?"),),
    ),
    PadraoTitulo(
        id="ipv_apos_informacoes", grupo=G_IPV, nome="IPV — após acréscimo de informações",
        modelo_titulo="Índice de Probabilidade de Voto (IPV) para {cargo}/2026 - {candidato}, após acréscimo de informações",
        modelo_pergunta="Qual a chance de você votar em {candidato} para {cargo}?",
        campos=(_C_CARGO_DEP, _C_CAND), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("ipv", "apos informacoes", "depois da apresentacao", "acrescimo"),
        variantes=(("Osasco", "Índice de Probabilidade de Voto (IPV) para Deputado Estadual/2026 - Rogério Lins, após acréscimo de informações (Estimulada e Única)"),),
    ),
    PadraoTitulo(
        id="chance_chapa", grupo=G_IPV, nome="Chance de voto — Chapa",
        modelo_titulo="Chance de voto (Chapa) {chapa}/2026",
        modelo_pergunta="Qual a chance de você votar na dupla: {chapa}?",
        campos=(Campo("chapa", "Chapa", "Lula para Presidente e Benedita da Silva para Senadora"),),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("chapa", "dupla", "tres candidaturas", "chance"),
        observacao="Para três candidaturas a pergunta vira «nessas três candidaturas: ...».",
        variantes=(("Benedita 2", "Chance de voto (Chapa) Lula para Presidente, Eduardo Paes para Governador e Benedita da Silva para Senadora/2026 (Estimulada e Única)"),),
        perguntas_vistas=(("Benedita 2", "Qual a chance de você votar nessas três candidaturas: Lula para Presidente, Eduardo Paes para governador e Benedita da Silva para Senadora?"),),
    ),
    PadraoTitulo(
        id="perfil_candidato", grupo=G_IPV, nome="Perfil de candidato de interesse",
        modelo_titulo="Perfil de candidato de interesse para {cargo}/2026",
        modelo_pergunta="Qual desses é o perfil de candidato a {cargo} que você tem mais interesse em votar?",
        campos=(_C_CARGO_DEP,), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("perfil", "interesse", "tipo de candidato"),
        variantes=(("Alphaville", "Perfil de candidato a Deputado Federal de maior interesse (Estimulada e Única)"),),
    ),

    # ============================================================ conhecimento e imagem
    PadraoTitulo(
        id="conhecimento", grupo=G_CONHECIMENTO, nome="Nível de conhecimento do candidato",
        modelo_titulo="Nível de conhecimento de {candidato}",
        modelo_pergunta="Você conhece muito, conhece de nome, conhece pela foto ou nunca ouviu falar sobre {candidato}?",
        campos=(_C_CAND,), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("conhecimento", "conhece", "notoriedade", "conhecimento inicial"),
        observacao="Empate entre «Nível de conhecimento de» (Osasco) e «Conhecimento inicial sobre» (Alphaville); vale a primeira.",
        variantes=(("Alphaville", "Conhecimento inicial sobre Wesley Cezar (Estimulada e Única)"),),
        perguntas_vistas=(("Osasco", "Você conhece muito, conhece de nome, conhece pela foto ou nunca ouviu falar sobre essa pessoa?"),),
    ),
    PadraoTitulo(
        id="conhecimento_apos_info", grupo=G_CONHECIMENTO, nome="Conhecimento — após acréscimo de informações",
        modelo_titulo="Nível de conhecimento de {candidato}, após acréscimo de informações",
        modelo_pergunta="Após estas informações, você conhece muito, conhece de nome, conhece pela foto ou nunca ouviu falar de {candidato}?",
        campos=(_C_CAND,), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("conhecimento", "apos informacoes", "acrescimo", "apresentacao"),
        variantes=(("Alphaville", "Conhecimento sobre Wesley Cezar após apresentação de informações (Estimulada e Única)"),),
    ),
    PadraoTitulo(
        id="imagem_candidato", grupo=G_CONHECIMENTO, nome="Imagem do candidato (atributo)",
        modelo_titulo="Imagem de {candidato} - {atributo}",
        modelo_pergunta="Pelo que você sabe ou imagina, {candidato} {afirmacao}?",
        campos=(
            _C_CAND,
            Campo("atributo", "Atributo", "Preparo para exercer o cargo"),
            Campo("afirmacao", "Afirmação da pergunta", "é preparado para exercer o cargo de Deputado Federal"),
        ),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("imagem", "atributo", "preparo", "influencia"),
        variantes=(("Alphaville", "Imagem de Wesley Cezar - Vontade de trazer melhorias para a cidade (Estimulada e Única)"),),
    ),
    PadraoTitulo(
        id="conhece_quem_e", grupo=G_CONHECIMENTO, nome="Para quem conhece o candidato (quem é)",
        modelo_titulo="Para quem conhece o candidato em algum nível",
        modelo_pergunta="Quem é?",
        estimulada=None, resposta=None,
        palavras_chave=("quem e", "conhece", "aberta"),
        observacao="Aparece sem sufixo de tipo de resposta no relatório de Osasco; mantido assim.",
    ),

    # ============================================================ avaliação e aprovação
    PadraoTitulo(
        id="aval_gov_municipal", grupo=G_AVALIACAO, nome="Avaliação do Governo Municipal (nota)",
        modelo_titulo="Avaliação do Governo Municipal de {cidade}",
        modelo_pergunta="Avalie o Governo Municipal de {cidade}, com uma nota entre zero e dez.",
        campos=(_C_CIDADE,), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("avaliacao", "governo", "prefeitura", "nota", "municipal"),
        variantes=(
            ("Osasco", "Avaliação do Governo Municipal de Osasco (SP) (Estimulada e Única)"),
            ("Nilópolis", "Avaliação do Governo da Prefeitura de Nilópolis"),
        ),
        perguntas_vistas=(
            ("Nilópolis", "Em uma escala de 0 a 10 em que 0 é péssima e 10 excelente, como você avalia a Prefeitura de Nilópolis?"),
        ),
    ),
    PadraoTitulo(
        id="aval_gov_estadual", grupo=G_AVALIACAO, nome="Avaliação do Governo Estadual (nota)",
        modelo_titulo="Avaliação do Governo Estadual de {governador}",
        modelo_pergunta="Avalie o Governo Estadual de {governador}, com uma nota entre zero e dez.",
        campos=(Campo("governador", "Governador", "Tarcísio de Freitas"),),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("avaliacao", "governo", "governador", "nota", "estadual"),
        variantes=(
            ("Osasco", "Avaliação do Governo Estadual Tarcísio de Freitas em 2025 (Estimulada e Única)"),
            ("Osasco (Análise das Perguntas)", "Avaliação do Governo Tarcísio de Freitas (Estimulada e Única)"),
        ),
    ),
    PadraoTitulo(
        id="aval_gov_federal", grupo=G_AVALIACAO, nome="Avaliação do Governo Federal (nota)",
        modelo_titulo="Avaliação do Governo Federal de {presidente}",
        modelo_pergunta="Avalie o Governo Federal de {presidente}, com uma nota entre zero e dez.",
        campos=(Campo("presidente", "Presidente", "Lula"),),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("avaliacao", "governo", "presidente", "nota", "federal", "lula"),
        variantes=(
            ("Angra dos Reis", "Avaliação do Governo do Presidente Lula (Estimulada e Única)"),
            ("Osasco", "Avaliação do Governo Federal Lula (Estimulada e Única)"),
        ),
        perguntas_vistas=(("Angra dos Reis", "Como o(a) Sr.(a) avalia o Governo do Presidente Lula?"),),
    ),
    PadraoTitulo(
        id="aval_gestao_prefeito", grupo=G_AVALIACAO, nome="Avaliação da gestão do prefeito (nota)",
        modelo_titulo="Avaliação da Gestão de {prefeito}",
        modelo_pergunta="Avalie a gestão de {prefeito}, com uma nota entre zero e dez.",
        campos=(_C_PREFEITO,), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("avaliacao", "gestao", "prefeito", "administracao", "nota"),
        variantes=(
            ("Alphaville", "Avaliação da gestão de Elvis Cezar (Estimulada e Única)"),
            ("Osasco", "Avaliação da administração do Prefeito Rogério Lins (Estimulada e Única)"),
        ),
        perguntas_vistas=(("Osasco", "Avalie a gestão de Rogério Lins como prefeito de Osasco, com uma nota entre zero e dez."),),
    ),
    PadraoTitulo(
        id="aval_governo_prefeito", grupo=G_AVALIACAO, nome="Avaliação do Governo do Prefeito (conceito)",
        modelo_titulo="Avaliação do Governo do Prefeito {prefeito}",
        modelo_pergunta="Como o(a) Sr.(a) avalia o Governo do Prefeito {prefeito}?",
        campos=(Campo("prefeito", "Prefeito", "Ferreti"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("avaliacao", "governo", "prefeito", "conceito", "otimo", "bom", "ruim"),
        observacao="Versão com escala de conceito (Ótimo, Bom...); a de nota de 0 a 10 é «Avaliação da Gestão de ...».",
    ),
    PadraoTitulo(
        id="aprovacao_prefeito", grupo=G_AVALIACAO, nome="Aprovação do Prefeito",
        modelo_titulo="Aprovação do Prefeito {prefeito}",
        modelo_pergunta="O(a) Sr.(a) aprova ou desaprova o modo como o Prefeito {prefeito} vem governando a cidade?",
        campos=(Campo("prefeito", "Prefeito", "Ferreti"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("aprovacao", "aprova", "desaprova", "prefeito"),
        observacao="Empate entre «Aprovação do Prefeito» (Angra) e «Percepção sobre a Gestão de» (Osasco); vale a primeira.",
        variantes=(("Osasco", "Percepção sobre a Gestão de Gerson Pessoa (Estimulada e Única)"),),
        perguntas_vistas=(("Osasco", "Você aprova ou desaprova a gestão de Gerson Pessoa?"),),
    ),
    PadraoTitulo(
        id="melhor_acao", grupo=G_AVALIACAO, nome="Melhor ação do Governo do Prefeito",
        modelo_titulo="Melhor ação do Governo do Prefeito {prefeito}",
        modelo_pergunta="O que o(a) Sr.(a) acredita que o Governo do Prefeito {prefeito} tem feito de MELHOR?",
        campos=(Campo("prefeito", "Prefeito", "Ferreti"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("melhor acao", "melhor", "acao do governo"),
    ),
    PadraoTitulo(
        id="pior_acao", grupo=G_AVALIACAO, nome="Pior ação do Governo do Prefeito",
        modelo_titulo="Pior ação do Governo do Prefeito {prefeito}",
        modelo_pergunta="O que o(a) Sr.(a) acredita que o Governo do Prefeito {prefeito} tem feito de PIOR?",
        campos=(Campo("prefeito", "Prefeito", "Ferreti"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("pior acao", "pior", "acao do governo"),
    ),
    PadraoTitulo(
        id="aval_camara", grupo=G_AVALIACAO, nome="Avaliação da Câmara de Vereadores",
        modelo_titulo="Avaliação da Câmara de Vereadores de {cidade}",
        modelo_pergunta="Como o(a) Sr.(a) avalia o trabalho da Câmara de Vereadores de {cidade}?",
        campos=(Campo("cidade", "Cidade", "Angra dos Reis"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("avaliacao", "camara", "vereadores", "legislativo"),
    ),
    PadraoTitulo(
        id="aval_autoridade", grupo=G_AVALIACAO, nome="Avaliação de uma autoridade (cargo e nome)",
        modelo_titulo="Avaliação do {autoridade} - {nome}",
        modelo_pergunta="Como o(a) Sr.(a) avalia o trabalho do {autoridade} - {nome}?",
        campos=(
            Campo("autoridade", "Cargo da autoridade", "Governador Interino"),
            Campo("nome", "Nome", "Desembargador Ricardo Couto"),
        ),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("avaliacao", "autoridade", "governador interino", "trabalho"),
    ),

    # ============================================================ posicionamento político
    PadraoTitulo(
        id="posicionamento", grupo=G_POSICIONAMENTO, nome="Posicionamento político",
        modelo_titulo="Posicionamento político",
        modelo_pergunta="Qual dessas opções o(a) Sr.(a) ou você acredita que melhor o(a) define politicamente?",
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("posicionamento", "esquerda", "direita", "centro", "ideologia"),
        observacao="Empate entre «político» e «Político»; vale a minúscula, como nos demais títulos.",
        variantes=(("São Paulo Online", "Posicionamento Político (Estimulada e Única)"),),
        perguntas_vistas=(("São Paulo Online", "Qual desses é o seu posicionamento político?"),),
    ),

    # ============================================================ hábitos de mídia
    PadraoTitulo(
        id="meios_comunicacao", grupo=G_MIDIA, nome="Meios de comunicação para saber as notícias",
        modelo_titulo="Meios de comunicação para saber as notícias de {local}",
        modelo_pergunta="Por qual meio de comunicação você fica sabendo das notícias de {local}?",
        campos=(Campo("local", "Cidade ou Estado", "Angra dos Reis"),), estimulada=ESTIMULADA, resposta=MULTIPLA,
        palavras_chave=("meios", "noticias", "comunicacao", "midia", "acompanhar"),
        variantes=(("São Paulo Online", "Principais meios para acompanhar notícias sobre o Estado (Estimulada e Múltipla)"),),
        perguntas_vistas=(("São Paulo Online", "Considerando estes meios, quais os três principais que você costuma acompanhar notícias sobre o Estado de São Paulo?"),),
    ),
    PadraoTitulo(
        id="interesse_informar", grupo=G_MIDIA, nome="Grau de interesse em se informar sobre a Prefeitura",
        modelo_titulo="Grau de interesse em se informar sobre as ações da Prefeitura",
        modelo_pergunta="Qual é o grau de seu interesse em se informar sobre as ações da Prefeitura de {cidade}?",
        campos=(Campo("cidade", "Cidade", "Angra dos Reis"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("interesse", "informar", "prefeitura", "acoes"),
    ),
    PadraoTitulo(
        id="prioridade_divulgacao", grupo=G_MIDIA, nome="Prioridades de divulgação da Prefeitura",
        modelo_titulo="Prioridades de divulgação da Prefeitura de {cidade}",
        modelo_pergunta="Para você, o que seria mais importante a Prefeitura de {cidade} divulgar?",
        campos=(Campo("cidade", "Cidade", "Angra dos Reis"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("divulgacao", "prioridades", "prefeitura", "comunicacao"),
    ),
    PadraoTitulo(
        id="lembranca_acao", grupo=G_MIDIA, nome="Lembrança de ação ou projeto da Prefeitura",
        modelo_titulo="Lembrança de ação ou projeto da Prefeitura neste ano",
        modelo_pergunta="Você se lembra de alguma ação ou projeto importante já realizado pela Prefeitura de {cidade} esse ano?",
        campos=(Campo("cidade", "Cidade", "Angra dos Reis"),), estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("lembranca", "acao", "projeto", "prefeitura", "recall"),
    ),
    PadraoTitulo(
        id="tv_aberta", grupo=G_MIDIA, nome="Emissora de TV aberta mais assistida",
        modelo_titulo="Emissora de TV aberta mais assistida",
        modelo_pergunta="Em se falando em TV aberta, qual o(a) Sr.(a) ou você mais assiste?",
        estimulada=ESTIMULADA, resposta=UNICA, palavras_chave=("tv", "televisao", "emissora", "assiste"),
    ),
    PadraoTitulo(
        id="radio", grupo=G_MIDIA, nome="Emissora de rádio mais ouvida",
        modelo_titulo="Emissora de rádio mais ouvida",
        modelo_pergunta="Em se falando em rádio, qual a emissora você mais ouve?",
        estimulada=ESTIMULADA, resposta=UNICA, palavras_chave=("radio", "emissora", "ouve"),
    ),
    PadraoTitulo(
        id="jornal_diario", grupo=G_MIDIA, nome="Jornal diário mais lido",
        modelo_titulo="Jornal diário mais lido",
        modelo_pergunta="Em se falando em jornal diário, qual você mais lê?",
        estimulada=ESTIMULADA, resposta=UNICA, palavras_chave=("jornal", "diario", "le"),
    ),
    PadraoTitulo(
        id="jornal_periodico", grupo=G_MIDIA, nome="Jornal periódico mais lido",
        modelo_titulo="Jornal periódico mais lido",
        modelo_pergunta="Em se falando em jornal periódico, qual você mais lê?",
        estimulada=ESTIMULADA, resposta=UNICA, palavras_chave=("jornal", "periodico", "le"),
    ),
    PadraoTitulo(
        id="redes_participacao", grupo=G_MIDIA, nome="Participação nas redes sociais",
        modelo_titulo="Participação nas redes sociais",
        modelo_pergunta="Qual (ou quais) redes sociais você participa?",
        estimulada=ESTIMULADA, resposta=MULTIPLA,
        palavras_chave=("redes sociais", "instagram", "facebook", "participa"),
    ),
    PadraoTitulo(
        id="redes_confianca", grupo=G_MIDIA, nome="Confiança nas notícias das redes sociais",
        modelo_titulo="Confiança nas notícias veiculadas nas redes sociais",
        modelo_pergunta="Mas você acredita e confia nas notícias veiculadas nas redes sociais?",
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("redes sociais", "confianca", "noticias", "acredita"),
    ),

    # ============================================================ outras
    PadraoTitulo(
        id="importancia_aliado", grupo=G_OUTROS, nome="Importância de um parlamentar aliado ao Prefeito",
        modelo_titulo="Percepção sobre a importância de um {cargo} aliado ao Prefeito",
        modelo_pergunta="O quanto ter um {cargo} como aliado ajuda a gestão do Prefeito de {cidade}?",
        campos=(_C_CARGO_DEP, Campo("cidade", "Cidade", "Santana de Parnaíba")),
        estimulada=ESTIMULADA, resposta=UNICA,
        palavras_chave=("percepcao", "aliado", "prefeito", "importancia"),
    ),
    PadraoTitulo(
        id="partidos_deputados", grupo=G_OUTROS, nome="Partidos/Federações dos Deputados",
        modelo_titulo="Partido/Federação dos {cargo_plural}",
        modelo_pergunta="",
        campos=(Campo("cargo_plural", "Cargo (plural)", "Deputados Estaduais", ("Deputados Estaduais", "Deputados Federais")),),
        estimulada=None, resposta=None,
        palavras_chave=("partido", "federacao", "federacoes", "deputados"),
        observacao="Tabela informativa, sem pergunta nem sufixo. O Total do RJ Deputados usa «Federações/Partidos dos ...».",
        variantes=(("RJ Deputados", "Federações/Partidos dos Deputados Estaduais"),),
    ),
)


# ---------------------------------------------------------------- acesso
def por_id(padrao_id: str) -> PadraoTitulo:
    for p in PADROES:
        if p.id == padrao_id:
            return p
    raise KeyError(f"Padrão de título desconhecido: {padrao_id}")


def padroes_do_grupo(grupo: Optional[str] = None) -> List[PadraoTitulo]:
    return [p for p in PADROES if grupo in (None, "", p.grupo)]


def _campos_do_modelo(modelo: str) -> List[str]:
    return [nome for _, nome, _, _ in Formatter().parse(modelo) if nome]


def _limpa(texto: str) -> str:
    return " ".join(str(texto).split())


def montar_titulo(padrao: PadraoTitulo, valores: Optional[Dict[str, str]] = None, parcial: bool = False) -> str:
    """Título padrão: modelo preenchido + sufixo do tipo de resposta.

    `parcial=True` deixa «‹campo›» no lugar do que falta (para mostrar na tela enquanto a pessoa preenche);
    sem isso, campo faltando é erro (ValueError), porque nunca se inventa o nome do candidato/cidade."""
    return _preenche(padrao.modelo_titulo, padrao, valores, parcial, com_sufixo=True)


def montar_pergunta(padrao: PadraoTitulo, valores: Optional[Dict[str, str]] = None, parcial: bool = False) -> str:
    """Texto da pergunta (o que vai depois de «Pergunta: » no CAPTION do .sps)."""
    return _preenche(padrao.modelo_pergunta, padrao, valores, parcial, com_sufixo=False)


def _preenche(modelo: str, padrao: PadraoTitulo, valores: Optional[Dict[str, str]], parcial: bool, com_sufixo: bool) -> str:
    valores = {k: _limpa(v) for k, v in (valores or {}).items() if v is not None and _limpa(v)}
    faltam = [c for c in _campos_do_modelo(modelo) if c not in valores]
    if faltam and not parcial:
        raise ValueError("Faltou preencher: " + ", ".join(faltam))
    cheio = {c: valores.get(c, f"‹{c}›") for c in _campos_do_modelo(modelo)}
    texto = _limpa(modelo.format(**cheio))
    if com_sufixo and padrao.sufixo:
        texto = f"{texto} {padrao.sufixo}"
    return texto


def exemplo_de(padrao: PadraoTitulo) -> Dict[str, str]:
    return {c.nome: c.exemplo for c in padrao.campos}


def titulo_exemplo(padrao: PadraoTitulo) -> str:
    return montar_titulo(padrao, exemplo_de(padrao))


def linha_title_sps(titulo: str) -> str:
    """O título como entra no .sps: «/TITLE 'Titulo: ...'» (apóstrofo dobrado, como o SPSS exige)."""
    return "/TITLE 'Titulo: " + titulo.replace("'", "''") + "'"


def linha_caption_sps(pergunta: str) -> str:
    return "/CAPTION 'Pergunta: " + pergunta.replace("'", "''") + "'"


# ---------------------------------------------------------------- busca
def _sem_acento(texto: str) -> str:
    t = unicodedata.normalize("NFD", str(texto or ""))
    return "".join(c for c in t if unicodedata.category(c) != "Mn").lower()


def _palavras(texto: str) -> List[str]:
    return [w for w in re.split(r"[^a-z0-9º]+", _sem_acento(texto)) if w]


def _texto_busca(p: PadraoTitulo) -> str:
    partes = [p.nome, p.grupo, p.modelo_titulo, p.modelo_pergunta, p.observacao, " ".join(p.palavras_chave)]
    partes += [t for _, t in p.variantes] + [t for _, t in p.perguntas_vistas]
    partes += [titulo_exemplo(p)]
    return _sem_acento(" ".join(partes))


def buscar(consulta: str, grupo: Optional[str] = None) -> List[PadraoTitulo]:
    """Padrões que contêm TODAS as palavras da consulta (sem acento, sem maiúscula). Sem consulta devolve
    todos do grupo. Os que casam no nome vêm primeiro."""
    base = padroes_do_grupo(grupo)
    palavras = _palavras(consulta)
    if not palavras:
        return base
    achados: List[Tuple[int, int, PadraoTitulo]] = []
    for ordem, p in enumerate(base):
        corpo = _texto_busca(p)
        if all(w in corpo for w in palavras):
            no_nome = _sem_acento(p.nome + " " + " ".join(p.palavras_chave))
            achados.append((0 if all(w in no_nome for w in palavras) else 1, ordem, p))
    achados.sort(key=lambda t: (t[0], t[1]))
    return [p for _, _, p in achados]


# ---------------------------------------------------------------- catálogo em tabela / Excel
COLUNAS_CATALOGO = ("Grupo", "Tipo de pergunta", "Título padrão (modelo)", "Pergunta (modelo)", "Tipo de resposta",
                    "Exemplo de título", "Outras formas já vistas", "Observação")


def catalogo_como_linhas(padroes: Optional[Iterable[PadraoTitulo]] = None) -> List[Dict[str, str]]:
    linhas = []
    for p in (padroes if padroes is not None else PADROES):
        linhas.append({
            "Grupo": p.grupo,
            "Tipo de pergunta": p.nome,
            "Título padrão (modelo)": _limpa(f"{p.modelo_titulo} {p.sufixo}"),
            "Pergunta (modelo)": p.modelo_pergunta,
            "Tipo de resposta": p.sufixo.strip("()") or "—",
            "Exemplo de título": titulo_exemplo(p),
            "Outras formas já vistas": "\n".join(f"{o}: {t}" for o, t in p.variantes),
            "Observação": p.observacao,
        })
    return linhas


def catalogo_excel(padroes: Optional[Iterable[PadraoTitulo]] = None) -> bytes:
    """Catálogo completo em Excel (uma aba), para consulta fora do app."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "Padrões de título"
    ws.append(list(COLUNAS_CATALOGO))
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFD9E1F2")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for linha in catalogo_como_linhas(padroes):
        ws.append([linha[c] for c in COLUNAS_CATALOGO])
    for col, larg in zip("ABCDEFGH", (26, 38, 58, 58, 22, 58, 70, 50)):
        ws.column_dimensions[col].width = larg
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- etapa seguinte (ainda não implementada)
def aplicar_padronizacao_titulos(ws):
    """
    Reservado para a etapa seguinte: aplicar os padrões acima aos títulos de uma planilha.

    Por ora o módulo é só de consulta (veja `buscar`, `montar_titulo` e `montar_pergunta`).

    Args:
        ws: worksheet do openpyxl a ser processada.

    Returns:
        Por enquanto, sempre 0 (nenhuma alteração aplicada).
    """
    return 0
