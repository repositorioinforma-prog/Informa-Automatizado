# Ponderação — componentes de peso dinâmicos (v6)

## Objetivo

A quantidade de variáveis usadas na ponderação não possui limite fixo de cinco.
Por padrão, cada variável selecionada recebe um componente SPSS próprio:

- 1ª variável -> `peso1`
- 2ª variável -> `peso2`
- ...
- 16ª variável -> `peso16`

A associação pode ser editada na interface. Várias variáveis podem compartilhar o
mesmo componente. Exemplo: seis variáveis atribuídas ao número `1` compõem juntas
`peso1`.

## Regra matemática

> **Atualização v7:** o componente passou a ser usado também pelo motor matemático,
> especialmente para variáveis alternativas e mutuamente exclusivas. Consulte
> `PONDERACAO_COMPONENTES_V7.md`.

Se sexo e idade estiverem preenchidos em `peso1`, seus fatores são multiplicados.
Se várias variáveis regionais estiverem em `peso5` e somente uma delas estiver
preenchida por entrevistado, o fator disponível é utilizado sem exigir as demais.

## Registro histórico

A aba `Registro` agora possui a coluna `Componente`, registra todas as variáveis
selecionadas e grava `Diferença (p.p.)` e `Status` como valores, evitando células
sem valor armazenado em leitores que não recalculam fórmulas do Excel.

## Compatibilidade

Se o usuário não editar os componentes, o comportamento anterior é preservado:
cada variável recebe seu próprio `pesoN` na ordem de seleção.
