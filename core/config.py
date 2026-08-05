"""
Constantes de configuração compartilhadas entre os módulos de análise.
Nenhuma dependência do Streamlit aqui de propósito — só dados/lógica pura,
para poder ser testado e reutilizado sem precisar rodar `streamlit run`.
"""

# Variáveis normalmente usadas nos cruzamentos de Mapas de Correspondência.
# Quando um arquivo é carregado, o app verifica quais delas existem e já
# sugere/pré-seleciona no módulo "Múltiplas Variáveis". Edite esta lista
# livremente conforme o questionário mudar.
VARIAVEIS_INTERESSE_PADRAO = [
    "P1", "P2", "P3", "P4", "P5", "juncaoreligiao", "P7", "P7_C",
]

# Variáveis de cota comumente usadas no módulo de Exclusões (equilíbrio de
# amostra). Auto-detectadas quando presentes no arquivo carregado.
VARIAVEIS_EXCLUSAO_PADRAO = ["P1", "P2", "P3", "P7"]

# Nomes amigáveis para exibir no lugar do código da variável (ex.: nos
# títulos dos slides do PowerPoint exportado). Variáveis sem entrada aqui
# continuam aparecendo com o próprio código (ex.: "juncaoreligiao").
ROTULOS_VARIAVEIS = {
    "P1": "SEXO",
    "P2": "IDADE",
    "P3": "RENDA",
    "P4": "ESCOLARIDADE",
    "P5": "RAÇA",
    "P7": "REGIÕES",
    "P7_C": "CAPITAL",
}


def _rotulo_amigavel(variavel):
    """Retorna o nome amigável da variável, se houver; senão, o próprio código."""
    return ROTULOS_VARIAVEIS.get(variavel, variavel)


# Escala de intenção de voto comumente usada como VARIÁVEL PRINCIPAL nos
# Mapas de Correspondência. Quando a variável escolhida como principal tem
# categorias que batem com esta escala (a comparação usa o texto das
# categorias, não o nome/código da variável — então funciona não importa
# se a coluna se chama P7, Q7 etc.), a configuração de categorias já vem
# pré-marcada mostrando só "Certamente não voto" e "Certamente voto"; as
# demais já aparecem marcadas para remover (o usuário pode desmarcar se
# quiser mantê-las). Edite estes conjuntos se a escala mudar.
CATEGORIAS_ESCALA_INTENCAO_VOTO = {
    "certamente não voto",
    "certamente voto",
    "ns/nr",
    "não conheço suficiente para votar",
    "provavelmente não voto",
    "provavelmente voto",
    "não sabe",
}

CATEGORIAS_MANTER_PADRAO_INTENCAO_VOTO = {
    "certamente não voto",
    "certamente voto",
}


def _normalizar_categoria(valor):
    """Normaliza um valor de categoria para comparação (minúsculas, sem espaços nas pontas)."""
    return str(valor).strip().lower()
