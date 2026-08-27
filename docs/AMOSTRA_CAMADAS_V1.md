# Montagem de Amostra por Camadas — v1

Esta versão introduz um segundo modo no **Gerador de Amostra** sem remover o modo clássico.

## Objetivo

Permitir que uma amostra seja distribuída de forma hierárquica, garantindo fechamento entre níveis territoriais. Exemplo:

Brasil → UF → Região Intermediária → Região Imediata → Município

Se uma UF recebe 200 entrevistas, a soma de todas as unidades da camada seguinte dentro dessa UF também será exatamente 200. O arredondamento é resolvido dentro de cada pai usando o método do Maior Resto.

## Como usar

No Gerador de Amostra, selecione **Por camadas (Beta)**.

1. Escolha o território inicial (Brasil ou uma UF).
2. Se necessário, filtre Regiões Intermediárias.
3. Marque as camadas territoriais que deseja controlar.
4. Informe o total de entrevistas e a base populacional.
5. Clique em **Gerar amostra**.

A tela apresenta um resumo de fechamento de cada camada e a amostra final no nível mais detalhado selecionado.

## Exportação

No modo por camadas, o Excel passa a incluir:

- `Amostra Completa` — distribuição no nível final;
- `Controle - UF` — quando UF foi selecionada;
- `Controle - Reg Interm` — quando Região Intermediária foi selecionada;
- `Controle - Reg Imed` — quando Região Imediata foi selecionada;
- `Controle - Município` — quando Município foi selecionado;
- `Controle - Distrito` — quando Distrito foi selecionado.

As abas de controle mostram população-base, percentual no total, percentual dentro do pai e amostra atribuída.

## Escopo desta v1

Implementado:

- distribuição proporcional;
- hierarquia recursiva;
- possibilidade de pular níveis;
- fechamento exato pai → filhos;
- validações automáticas;
- exportação das camadas;
- manutenção do modo clássico.

Ainda não implementado nesta versão:

- mínimo por unidade + proporcional;
- quantidade fixa;
- edição manual e unidades travadas;
- cotas hard/soft;
- territórios personalizados/bairros;
- regras diferentes em ramos diferentes da árvore.

## Pontos metodológicos já identificados e ainda pendentes

1. Na base distrital, a faixa chamada `16_19` usa 15–19 na fonte atual. Portanto, a base 16+ distrital ainda precisa de decisão metodológica/correção de fonte.
2. A cota de sexo continua usando `masc_total` / `fem_total`, mesmo quando o peso territorial usa população 16+.
3. Unidades sem dados de renda ainda podem produzir cotas de renda zeradas.

Esses pontos foram mantidos para que a v1 não altere silenciosamente regras demográficas já existentes enquanto a arquitetura territorial é introduzida.
