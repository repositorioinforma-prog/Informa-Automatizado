# Ponderação — método institucional fixo v10

## Objetivo

Evitar que usuários selecionem por engano uma regra diferente do procedimento
adotado pela equipe nas ponderações operacionais.

## Regra aplicada

A interface utiliza sempre:

- **método:** razão simples e multiplicação dos componentes;
- **tolerância de validação:** 2,00 pontos percentuais.

Esses parâmetros deixaram de ser editáveis na tela. O método e a tolerância
continuam documentados no Registro Histórico e na memória de cálculo, permitindo
auditoria sem criar uma decisão metodológica desnecessária para o usuário final.

## Fluxo da interface

Depois de configurar os componentes de peso, a tela segue diretamente para:

1. preenchimento ou conferência do Universo;
2. cálculo da ponderação;
3. validação dos resultados;
4. geração dos arquivos.

A mensagem de validação não sugere mais o ajuste iterativo. Quando uma categoria
fica acima de 2,00 p.p., o usuário é orientado a revisar metas do Universo,
categorias e componentes de peso.

## Compatibilidade

O motor de razão simples, o agrupamento de variáveis no mesmo `pesoN`, o Registro
Histórico, a memória de cálculo, a syntax SPSS e o assistente de exclusões foram
preservados. Todos os recálculos acionados pela interface seguem a mesma regra fixa.
