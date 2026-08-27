# Recorte territorial direto no Gerador de Amostra — v3

## Objetivo

O Gerador agora permite restringir o universo da amostra em qualquer ponto da hierarquia territorial, sem obrigar o usuário a selecionar primeiro uma Região Intermediária.

Os filtros disponíveis são:

```text
Brasil ou uma UF
  → Regiões Intermediárias específicas
  → Regiões Imediatas específicas
  → Municípios específicos
  → Distritos específicos
```

Todos os filtros são opcionais. Uma seleção inferior pode ser feita diretamente. Por exemplo, é possível selecionar apenas os municípios São Paulo e Campinas deixando as regiões vazias.

## Regras de combinação

Os filtros são combinados por interseção:

- vazio significa **todas as unidades disponíveis** naquele nível;
- uma UF limita as opções dos níveis inferiores;
- uma Região Intermediária limita Regiões Imediatas, municípios e distritos;
- municípios selecionados limitam a lista de distritos;
- quando um filtro pai é alterado, seleções filhas que ficaram fora do novo recorte são removidas automaticamente.

Exemplo:

```text
UF: SP
Região Intermediária: Campinas
Municípios: Campinas e Valinhos
```

A distribuição usa somente Campinas e Valinhos.

## Municípios específicos

O campo **Municípios específicos** está sempre disponível. Os valores internos são os códigos IBGE, enquanto a interface exibe UF, nome e código para evitar ambiguidades.

Exemplos de uso:

```text
Saída por município
Recorte: São Paulo + Campinas
```

ou:

```text
Saída por Região Intermediária
Recorte: apenas municípios escolhidos
```

No segundo caso, o motor agrega a população dos municípios escolhidos dentro de cada região antes de distribuir a amostra.

## Distritos específicos

A base distrital é ativada automaticamente quando:

- o nível de quebra clássico é **Distrito**; ou
- a camada **Distrito** está selecionada.

Também pode ser ativada manualmente pela opção:

```text
Habilitar seleção de distritos específicos
```

Isso permite usar um recorte distrital mesmo quando a saída final é agregada em município ou região.

Exemplo:

```text
Distritos escolhidos: Água Rasa e Alto de Pinheiros
Saída final: Município
```

O resultado terá São Paulo como unidade final, mas sua população-base será formada somente pelos dois distritos escolhidos.

## Projetos JSON

O esquema dos projetos passou para `1.1` e guarda:

```json
{
  "territorio": {
    "uf": "SP",
    "regioes_intermediarias": [],
    "regioes_imediatas": [],
    "municipios": [
      {"codigo": "3550308", "nome": "São Paulo"}
    ],
    "distritos": [
      {"codigo": "355030801", "nome": "Água Rasa"}
    ]
  }
}
```

Projetos `1.0` continuam aceitos e são migrados em memória para `1.2` ao serem abertos.

## Contrato do motor

Os dois motores agora aceitam o parâmetro:

```python
codigos_distritos=["355030801", "355030802"]
```

Além dos parâmetros já existentes:

```python
codigos_regioes_intermediarias
codigos_regioes_imediatas
codigos_municipios
```

Quando `codigos_distritos` é informado, a base precisa possuir a coluna `cod_distrito`. Caso contrário, o motor retorna uma mensagem explícita em vez de ignorar o recorte.

## Limites atuais

- a seleção é múltipla, mas ainda não possui grupos nomeados de territórios;
- bairros e subdistritos continuam fora do motor de cálculo até existir uma política explícita para cobertura parcial;
- a distribuição pode ser proporcional ou manual no nível final selecionado;
- travamentos parciais com redistribuição automática do restante permanecem como evolução futura.
