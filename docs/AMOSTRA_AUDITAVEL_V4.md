# Amostra auditável, distribuição manual e exportação configurável — v4

## Objetivo

Esta versão amplia o Gerador de Amostra em três pontos:

1. permitir escolher entre distribuição proporcional e distribuição manual;
2. registrar no Excel a memória de cálculo utilizada;
3. permitir escolher como as cotas de gênero, idade e renda aparecem no arquivo final.

A amostra total e o percentual da amostra permanecem visíveis em qualquer formato de exportação.

## Modos de distribuição

### Proporcional à população

O usuário informa a amostra total e o sistema distribui as entrevistas de acordo com a população-base das unidades selecionadas.

O arredondamento usa o método do Maior Resto (Hamilton):

```text
População-base da unidade / população-base do recorte
  → amostra ideal decimal
  → parte inteira
  → entrevistas restantes distribuídas pelas maiores frações decimais
```

Exemplo:

```text
Amostra total: 500
Unidades: municípios selecionados
Método: proporcional
```

A soma das quantidades aplicadas fecha exatamente em 500.

### Metas manuais por unidade

O sistema calcula primeiro a distribuição proporcional e a mantém como referência. Em seguida, a interface apresenta uma tabela editável com:

- hierarquia territorial;
- código da unidade;
- população total;
- população-base;
- percentual real do recorte;
- amostra ideal decimal;
- sugestão proporcional arredondada;
- meta da unidade.

O usuário altera somente a coluna **Meta da unidade (n)**. O botão de aplicação fica bloqueado enquanto houver:

- valores vazios;
- números negativos;
- números não inteiros;
- códigos faltantes ou extras;
- soma diferente da amostra total.

Exemplo:

```text
Meta total: 500
Soma preenchida: 493
Situação: faltam 7 entrevistas
```

Quando a soma fecha em 500, o sistema aplica os valores, recalcula mínimo/máximo e refaz as cotas internas de gênero, idade e renda para cada unidade.

## Distribuição manual no modo por camadas

A edição manual acontece no nível mais detalhado entre a última camada configurada e as unidades territoriais selecionadas explicitamente. Se municípios específicos forem escolhidos, Município é acrescentado como nível final do cálculo manual quando necessário.

Exemplo:

```text
UF → Região Intermediária → Município
```

O usuário informa os valores municipais. Depois disso, o sistema reconstrói:

```text
Municípios
  → soma por Região Intermediária
  → soma por UF
```

Assim, os controles superiores refletem a alocação final e continuam fechando com seus filhos.

## Memória de cálculo

O resultado mantém separadamente:

- `populacao_total`: quantidade total de habitantes;
- `populacao_base`: população efetivamente usada no cálculo;
- `percentual`: participação populacional real no recorte;
- `percentual_no_pai`: participação populacional dentro do nível pai, quando há camadas;
- `amostra_pai`: quantidade proporcional recebida pelo nível pai;
- `amostra_ideal`: valor proporcional antes do arredondamento, calculado dentro do pai no modo por camadas;
- `amostra_proporcional`: sugestão arredondada pelo Maior Resto;
- `amostra`: quantidade efetivamente aplicada;
- `percentual_amostra`: participação da unidade na amostra final;
- `diferenca_vs_proporcional`: amostra aplicada menos sugestão proporcional;
- `metodo_distribuicao`: proporcional ou manual.

Em uma distribuição proporcional, a diferença é zero. Em uma distribuição manual, a coluna deixa explícito o desvio adotado pelo usuário.

No modo por camadas, o Excel refaz a memória com fórmulas: soma a sugestão proporcional dos filhos para obter a amostra do pai, calcula o percentual populacional dentro desse pai e multiplica os dois valores para chegar à amostra ideal hierárquica.

## Aba Informações do Excel

A primeira aba do arquivo exportado é **Informações**.

Ela contém:

### Configuração do projeto

- nome;
- data e hora de geração;
- fonte dos dados;
- modo clássico ou por camadas;
- método proporcional ou manual;
- base populacional;
- nível final;
- recorte territorial;
- amostra total informada;
- preferência de exibição das cotas.

### Explicação do método

A aba registra se o arquivo foi produzido pelo método proporcional ou por valores manuais e explica o papel da sugestão proporcional.

### Resumo das campanhas

Quando existem campanhas, a aba mostra para cada uma:

- quantidade de unidades;
- amostra atribuída;
- percentual da amostra completa.

### Memória por unidade territorial

Para cada unidade final, a aba apresenta:

- UF;
- Região Intermediária;
- Região Imediata;
- município;
- distrito;
- código;
- população total;
- população-base;
- percentual populacional real no recorte;
- percentual populacional dentro do nível pai;
- amostra proporcional recebida pelo nível pai;
- amostra ideal;
- sugestão proporcional;
- amostra aplicada;
- percentual da amostra;
- diferença versus proporcional;
- mínimo;
- máximo;
- método.

### Validações

O próprio Excel inclui fórmulas para conferir:

- fechamento da amostra total;
- fechamento percentual em 100%;
- diferença total versus distribuição proporcional;
- quantidade de unidades com amostra zero.

## Preferências de exibição das cotas

A interface possui três opções:

```text
Somente totais (n)
Somente percentuais (%)
Totais e percentuais
```

A preferência se aplica apenas aos campos que compõem:

- gênero;
- idade;
- renda.

### Somente totais

Exemplo:

```text
Masculino: 48
Feminino: 52
```

### Somente percentuais

Exemplo:

```text
Masculino: 48,00%
Feminino: 52,00%
```

As contagens necessárias às fórmulas permanecem em colunas auxiliares ocultas, evitando perda de auditabilidade.

### Totais e percentuais

Exemplo:

```text
Masculino (n): 48
Masculino (%): 48,00%
Feminino (n): 52
Feminino (%): 52,00%
```

Independentemente da opção, o arquivo sempre mantém:

```text
Amostra — Total
Amostra — % da amostra
TOTAL DA AMOSTRA
% DA AMOSTRA
```

## Renda ausente

Quando uma unidade não possui cobertura completa de renda:

- as cotas de renda permanecem vazias;
- o sistema não converte ausência em zero;
- totais e percentuais de renda também ficam vazios quando a soma seria parcial.

## Projetos JSON

O esquema dos projetos passou para `1.2` e pode guardar:

```json
{
  "distribuicao": {
    "metodo": "manual",
    "alocacoes_manuais": [
      {"codigo": "3550308", "amostra": 320},
      {"codigo": "3509502", "amostra": 180}
    ]
  },
  "exportacao": {
    "exibicao_cotas": "ambos"
  }
}
```

Projetos `1.0` e `1.1` continuam aceitos. Na abertura, recebem os padrões:

```text
Método: proporcional
Exibição das cotas: totais e percentuais
```

## Arquivos principais

```text
core/amostra_alocacao.py       validação e aplicação manual
core/amostra_math.py           memória de cálculo no motor clássico
core/amostra_camadas.py        reconstrução de pais após edição manual
core/amostra_exportador.py     aba Informações e layouts de cotas
core/projetos_amostra.py       esquema JSON 1.2
analises/gerador_amostra.py    editor e preferências no Streamlit
```

## Limites atuais

- no modo por camadas, a edição manual é feita no último nível; os níveis pais são derivados da soma dos filhos;
- recortes com centenas ou milhares de unidades podem deixar o editor manual mais pesado no navegador;
- bairros e subdistritos continuam fora do cálculo até existir uma política explícita para cobertura parcial;
- a edição manual ainda não possui travamentos parciais com redistribuição automática do restante.

> Atualização v5: quando há municípios, distritos ou regiões específicas selecionadas, o editor manual passa a usar automaticamente o nível mais detalhado do recorte. Consulte `METAS_MANUAIS_V5.md`.
