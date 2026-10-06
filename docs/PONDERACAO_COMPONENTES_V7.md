# Ponderação — componentes com variáveis alternativas (v7)

## Correção

A versão anterior permitia associar várias variáveis ao mesmo `pesoN`, mas o
motor matemático continuava exigindo que todas as variáveis selecionadas
estivessem preenchidas em cada caso.

Esse comportamento tornava todos os pesos ausentes quando o componente reunia
variáveis mutuamente exclusivas. Um exemplo é o bloco:

- `VN_sud`
- `VN_nord`
- `VN_sul`
- `VN_norte`
- `VN_centro`

Cada entrevistado responde somente à variável correspondente à sua região. Logo,
nenhum caso possui as cinco variáveis preenchidas ao mesmo tempo.

## Nova regra por componente

Para cada caso e para cada componente:

1. fatores das variáveis preenchidas e reconhecidas são considerados;
2. se apenas uma variável se aplica, o componente usa o fator dessa variável;
3. se mais de uma se aplica, os fatores disponíveis são multiplicados;
4. variáveis ausentes no grupo são tratadas como não aplicáveis;
5. se nenhuma variável do componente se aplica, o caso fica sem peso;
6. uma categoria preenchida sem fator configurado invalida o componente.

Depois, o peso final continua sendo o produto entre componentes:

```text
peso = peso1 × peso2 × peso3 × ...
```

## Exemplo regional

```text
P1        -> peso1
P2        -> peso2
P3        -> peso3
P7        -> peso4
VN_sud    -> peso5
VN_nord   -> peso5
VN_sul    -> peso5
VN_norte  -> peso5
VN_centro -> peso5
```

Para um entrevistado do Sudeste, `peso5` usa `VN_sud`. Para um entrevistado do
Nordeste, usa `VN_nord`, e assim por diante.

## Método iterativo

O raking também passou a aceitar variáveis alternativas agrupadas. Cada margem é
ajustada dentro dos casos em que a respectiva variável está preenchida, enquanto
a elegibilidade geral exige ao menos uma variável aplicável em cada componente.

## Validação na interface

A tela não apresenta mais `nan` como valor de validação. Quando existem casos sem
peso, o aplicativo mostra a quantidade e disponibiliza uma tabela de cobertura
por componente com:

- casos válidos;
- casos sem nenhuma variável aplicável;
- casos com categoria não mapeada;
- casos com mais de uma contribuição no mesmo componente.

## Syntax SPSS

A syntax passou a reproduzir a mesma regra do Python: valores ausentes dentro de
um componente agrupado não anulam automaticamente o peso. O componente utiliza
os fatores disponíveis e só fica ausente se nenhuma variável se aplicar ou se
houver uma categoria preenchida sem fator configurado.
