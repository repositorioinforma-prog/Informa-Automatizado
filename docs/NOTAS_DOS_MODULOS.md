# Notas dos módulos

Notas técnicas de módulos específicos (antes ficavam no fim do README). Para a visão geral do app, veja o [README](../README.md).

## Revisor de Syntax SPSS

O menu **Revisor de Syntax SPSS** audita arquivos `.sps` e aceita opcionalmente o banco `.sav` original e um `.sav` processado. A primeira versão verifica estrutura de `RECODE`, regras duplicadas/conflitantes, coerência com `VALUE LABELS`, seções vazias, encoding, referências de `MRSETS`, variáveis inexistentes e categorias reais do banco que seriam perdidas por `ELSE=SYSMIS` ou ficariam sem mapeamento. A análise é conservadora: ela não substitui o motor do IBM SPSS e só interpreta construções reconhecidas com segurança.

## Ferramentas analíticas (portadas do app "Relatória Analítica")

As ferramentas do antigo app desktop (PySide6) estão sendo liberadas uma por vez.

| Menu | O que faz | Lógica | Tela |
| --- | --- | --- | --- |
| Cruzamentos (Tabelas) | Tabelas ponderadas com ordenação automática, ajuste de linhas/colunas e pilha de exportação Excel | `core/cruzamentos_math.py` | `analises/cruzamentos.py` + `componentes/relatoria_cruzamentos/` |
| Mapa de Calor | Pinta KMLs com Mín/Média/Máx do relatório, numeração e fila de PPTX | `core/mapa_calor_math.py` | `analises/mapa_calor.py` + `componentes/relatoria_mapa_calor/` |
| Matriz de Correlação | Pearson/Spearman ponderados com Sig. (2-tailed), missing guiado e exportação CSV/XLSX | `core/correlacao_math.py` | `analises/correlacao.py` + `componentes/relatoria_correlacao/` |

As telas são as páginas HTML originais do app desktop rodando como componentes do Streamlit (`analises/_ponte_relatoria.py` + `componentes/ponte_qt.js`).

Próxima: Análise Residual.

## Leitura de .sav sem o pyreadstat

Se o pyreadstat não carregar (ex.: Windows com política de Controle de Aplicativo bloqueando a DLL dentro da `.venv`), o app lê os `.sav`/`.zsav` com um leitor em Python puro (`core/sav_puro.py`), automaticamente e com o mesmo resultado. Só a gravação de `.sav` (exportação da Ponderação) continua precisando do pyreadstat.

## Ponderação com componentes de peso dinâmicos

O módulo de Ponderação aceita mais de cinco variáveis. Por padrão, cada variável
recebe `peso1`, `peso2`, `peso3` e assim por diante, sem limite fixo. Na interface,
a coluna **Peso** pode ser editada para agrupar várias variáveis no mesmo
componente. O motor combina os fatores preenchidos em cada caso: variáveis
simultâneas são multiplicadas e variáveis alternativas por região podem compartilhar
o mesmo `pesoN` sem exigir que todas estejam preenchidas. O registro histórico,
a memória de cálculo e a syntax SPSS documentam essa associação.

### Fluxo do Universo na ponderação

A interface segue a ordem operacional: **selecionar as variáveis**, **baixar o
modelo do Universo** e **importar o mesmo arquivo preenchido**. Download e upload
ficam reunidos na mesma seção, reduzindo deslocamentos e deixando explícito que o
modelo é gerado somente com as variáveis escolhidas para aquela ponderação.

### Método institucional fixo de ponderação

A interface não oferece mais seleção de método nem edição da tolerância. Todas as
ponderações executadas pelo aplicativo utilizam **razão simples e multiplicação dos
componentes**, com validação fixa de **2,00 pontos percentuais**. O método continua
registrado nos arquivos de auditoria, mas não pode ser alterado acidentalmente pelo
usuário final.

### Assistente de exclusões mais compacto

O painel **Travas e proteções de exclusão** inicia recolhido por padrão. As regras
continuam disponíveis ao abrir o painel, mas deixam de ocupar grande parte da tela
antes de serem necessárias.
