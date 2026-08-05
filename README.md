# Aplicativo de Análise de Dados — Mapas de Correspondência

Aplicação web desenvolvida em **Python** e **Streamlit** para análise exploratória de dados, com foco na criação de **Mapas de Correspondência** a partir de duas variáveis categóricas.

O sistema permite carregar bases em CSV, Excel ou SPSS, construir automaticamente uma tabela de contingência, calcular a Análise de Correspondência e visualizar as relações entre as categorias em mapas estáticos e interativos.

## Visão geral

A Análise de Correspondência é uma técnica estatística utilizada para representar graficamente associações existentes em tabelas de contingência. No aplicativo, as categorias das duas variáveis selecionadas são posicionadas em um plano bidimensional.

Categorias próximas no mapa tendem a apresentar perfis semelhantes ou algum grau de associação. Categorias mais afastadas da origem geralmente possuem maior contribuição para a diferenciação dos resultados.

> A proximidade visual deve ser interpretada em conjunto com a inércia explicada pelas dimensões e com o contexto da pesquisa. O mapa indica padrões de associação, não causalidade.

## Principais funcionalidades dos Mapas de Correspondência

- Carregamento de arquivos nos formatos `.csv`, `.xlsx` e `.sav`;
- Leitura de rótulos de valores presentes em arquivos SPSS;
- Seleção de variáveis categóricas diretamente pela interface;
- Construção automática da tabela de contingência;
- Cálculo manual da Análise de Correspondência por decomposição em valores singulares — SVD;
- Exibição da inércia explicada pelas duas dimensões;
- Geração de mapa estático com Matplotlib;
- Geração de mapa interativo com Plotly;
- Edição dos nomes das categorias;
- Ajuste manual da posição dos rótulos;
- Remoção individual de pontos e rótulos do mapa;
- Personalização das legendas de linhas e colunas;
- Opção para exibir ou ocultar a legenda;
- Reposicionamento de anotações no gráfico interativo.

## Como a análise funciona

O aplicativo executa as seguintes etapas:

1. O usuário carrega uma base de dados.
2. Duas variáveis categóricas são selecionadas.
3. Uma tabela de contingência é criada com `pandas.crosstab`.
4. A tabela é convertida em proporções.
5. São calculadas as massas das linhas e das colunas.
6. As frequências observadas são comparadas às frequências esperadas sob independência.
7. A matriz padronizada é decomposta por SVD.
8. São obtidas as coordenadas principais das categorias.
9. Os resultados são apresentados em duas dimensões.

Embora a interface permita selecionar mais de duas colunas, a versão atual utiliza as **duas primeiras variáveis selecionadas** para produzir o mapa.

## Interpretação do mapa

No gráfico padrão:

- Os **círculos azuis** representam as categorias da variável posicionada nas linhas da tabela de contingência;
- Os **triângulos vermelhos** representam as categorias da variável posicionada nas colunas;
- O eixo horizontal corresponde à **Dimensão 1**;
- O eixo vertical corresponde à **Dimensão 2**;
- As linhas tracejadas indicam a origem dos eixos.

### Inércia explicada

Antes dos mapas, o aplicativo apresenta o percentual de inércia explicado por cada dimensão.

Quanto maior a soma da inércia das duas dimensões, melhor o plano bidimensional resume as associações presentes na tabela original. Quando esse percentual é baixo, parte importante da estrutura dos dados pode estar em dimensões não exibidas.

### Distância da origem

Categorias próximas do centro geralmente possuem perfis menos diferenciados. Categorias mais afastadas da origem tendem a ser mais relevantes para a formação das dimensões.

### Proximidade entre categorias

A proximidade entre categorias pode indicar associação ou semelhança de perfil, especialmente quando os pontos pertencem a conjuntos diferentes — linhas e colunas. Essa leitura deve considerar a qualidade de representação das dimensões e o tamanho das frequências observadas.

## Requisitos da base de dados

Para gerar um mapa adequado:

- Selecione duas variáveis categóricas;
- Evite variáveis com quantidade excessiva de categorias;
- Verifique a presença de categorias vazias ou com frequência muito baixa;
- Evite tabelas com apenas uma linha ou uma coluna contendo frequência positiva;
- Padronize previamente categorias duplicadas por diferenças de grafia, acentuação ou uso de maiúsculas e minúsculas;
- Em arquivos SPSS, confirme se os códigos e rótulos de valores estão corretos.

### Exemplo de estrutura

| Faixa etária | Avaliação do serviço |
|---|---|
| 18 a 24 anos | Ótimo |
| 25 a 34 anos | Bom |
| 35 a 44 anos | Regular |
| 18 a 24 anos | Bom |

Nesse exemplo, o mapa pode ser utilizado para explorar a associação entre as categorias de faixa etária e avaliação do serviço.

## Instalação

### 1. Clone ou baixe o projeto

```bash
git clone <URL_DO_REPOSITORIO>
cd app-analises-main
```

### 2. Crie um ambiente virtual

No Windows:

```bash
python -m venv .venv
.venv\Scripts\activate
```

No Linux ou macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

### 4. Execute o aplicativo

```bash
streamlit run app.py
```

Depois, acesse no navegador o endereço informado pelo Streamlit, normalmente:

```text
http://localhost:8501
```

## Como gerar um Mapa de Correspondência

1. Abra o aplicativo.
2. Carregue um arquivo CSV, Excel ou SPSS.
3. No menu lateral, selecione **Mapas de Correspondência**.
4. Escolha pelo menos duas variáveis categóricas.
5. Consulte a inércia explicada pelas dimensões.
6. Ajuste os nomes das legendas, caso necessário.
7. Ative **Editar rótulos e deslocamentos** para renomear, mover ou remover categorias.
8. Analise o mapa estático.
9. Utilize o mapa interativo para explorar e reposicionar visualmente os rótulos.

## Personalização dos rótulos

Ao ativar a opção **Editar rótulos e deslocamentos**, é possível:

- Substituir o nome exibido de cada categoria;
- Ajustar o deslocamento horizontal e vertical do texto;
- Remover uma categoria da visualização;
- Reduzir sobreposições em mapas com muitas categorias.

A remoção é apenas visual e não altera a tabela de contingência nem recalcula a análise.

## Formatos suportados

### CSV

- Leitura inicial em UTF-8;
- Tentativa automática com Latin-1 em caso de erro de codificação.

### Excel

- Arquivos `.xlsx`;
- Leitura da primeira planilha do arquivo.

### SPSS

- Arquivos `.sav`;
- Preservação dos valores numéricos para os cálculos;
- Uso dos rótulos de valores para exibição quando disponíveis.

## Tecnologias utilizadas

- Python;
- Streamlit;
- Pandas;
- NumPy;
- Matplotlib;
- Plotly;
- SciPy;
- Pyreadstat;
- AdjustText;
- OpenPyXL;
- Scikit-learn;
- Statsmodels.

A biblioteca `prince` permanece entre as dependências por compatibilidade, mas o módulo de Mapas de Correspondência utiliza uma implementação própria baseada em SVD.

## Outros módulos disponíveis

Além dos Mapas de Correspondência, o aplicativo possui opções para:

- Tratamento de dados;
- Geração de gráficos;
- Relatório automatizado;
- Relatório em Excel;
- Análise de resíduos;
- Análise de correlação;
- Frequência ponderada;
- Módulo de predição ainda não implementado nesta versão.

## Limitações atuais

- O mapa trabalha com apenas duas variáveis por análise;
- A seleção de mais de duas variáveis não gera uma análise de correspondência múltipla;
- Não há exportação direta do mapa em arquivo pela interface;
- A aplicação não apresenta contribuições, massas ou qualidade de representação de cada categoria;
- Categorias raras podem produzir posições extremas e dificultar a interpretação;
- O sentido dos eixos pode variar sem alterar a interpretação estatística;
- A logo é carregada de uma URL externa e pode não aparecer quando não houver conexão com a internet.

## Boas práticas de análise

- Verifique previamente a tabela de contingência;
- Agrupe categorias com frequências muito pequenas quando isso fizer sentido metodológico;
- Não interprete somente a proximidade visual;
- Observe a inércia explicada;
- Compare o mapa com frequências absolutas, percentuais e resíduos padronizados;
- Documente qualquer recodificação ou agrupamento realizado na base;
- Evite concluir causalidade a partir das associações observadas.

## Estrutura do projeto

```text
app-analises-main/
├── app.py
├── requirements.txt
└── README.md
```

## Possíveis melhorias futuras

- Exportação dos mapas em PNG, SVG, PDF e HTML;
- Download das coordenadas das categorias em Excel ou CSV;
- Exibição da tabela de contingência utilizada;
- Cálculo das contribuições das categorias para cada dimensão;
- Exibição de cos² e qualidade de representação;
- Filtros antes da análise;
- Suporte à Análise de Correspondência Múltipla;
- Personalização de cores, símbolos, títulos e nomes dos eixos;
- Geração automática de um relatório interpretativo;
- Salvamento das configurações de rótulos e deslocamentos.

## Instituição

Aplicativo desenvolvido para apoiar as rotinas de análise de dados do **Instituto Informa**.

## Licença

Defina neste espaço a licença aplicável ao projeto e as regras de uso, distribuição e modificação do código.
