# Ponderação — fluxo do Universo v9

## Objetivo

Organizar o preenchimento do Universo na sequência em que o usuário trabalha:

1. carregar a base SPSS;
2. selecionar as variáveis da ponderação;
3. baixar o modelo Excel criado para essa seleção;
4. preencher a coluna `% Universo`;
5. importar o mesmo arquivo imediatamente abaixo do botão de download;
6. configurar os componentes de peso e executar a ponderação.

## Comportamento da importação

O arquivo é pareado por `Variável SPSS + Código SPSS`. Valores referentes a
variáveis não selecionadas permanecem guardados na sessão, mas não participam do
cálculo enquanto essas variáveis não forem selecionadas. A interface informa:

- quantas variáveis selecionadas foram encontradas no arquivo;
- quais variáveis selecionadas estão sem Universo;
- quais variáveis extras existem no arquivo importado.

## Compatibilidade

O formato do modelo, o registro histórico, os editores manuais, os componentes
de peso e o motor matemático não foram alterados. A mudança é de organização da
interface e validação do fluxo.
