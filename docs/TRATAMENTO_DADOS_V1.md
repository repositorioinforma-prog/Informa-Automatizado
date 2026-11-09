# Tratamento de Dados — V1

Nova seção da tela **Início** (antes de "Programações") e novo item do menu lateral. Permite manipular o
banco `.sav` de forma parecida com o SPSS, com histórico, desfazer e geração da syntax equivalente.

## Arquivos
- `core/tratamento_dados_math.py` — núcleo puro (sem Streamlit): estado do banco, avaliador de expressões,
  operações, histórico e leitura/gravação/exportação.
- `analises/tratamento_dados.py` — tela (`modulo_tratamento_dados`): faixas de comandos por tarefa, janelas de
  diálogo, abas de baixo e barra de status. Chaves de estado com prefixo `td_`.
- `core/tratamento_dados_dialogos.py` — partes puras da tela (sem Streamlit): organização dos comandos em faixas, textos de ajuda,
  conversão das grades (regras do Recode, rótulos, valores ausentes, Variable View) em parâmetros do núcleo.
- `analises/tela_inicial.py` — seção "Tratamento de Dados" (primeira da lista `SECOES`).
- `app.py` — item "Tratamento de Dados" em `MENU` e no despacho.
- `core/sav_puro.py` — passou a ler o nível de medida (`meta.variable_measure`, registro subtipo 11).
- `controle/exporters/sav_sem_pyreadstat.py` — o escritor em Python puro ganhou parâmetros **opcionais**
  `missing`, `medida` e `decimais` (sem eles, o comportamento é o de antes).
- `tests/test_tratamento_dados.py` (núcleo), `tests/test_tratamento_dados_dialogos.py` (partes puras) e
  `tests/test_tratamento_dados_ui.py` (fiação da tela com um `streamlit` falso).

## Organização da tela
Os comandos ficam **sempre visíveis**, em abas (faixas) por **tarefa**, na ordem do fluxo de trabalho. Cada comando tem
ícone, nome e uma frase curta; a estrutura está em `FAIXAS` (core/tratamento_dados_dialogos.py).

| Faixa | Grupos → comandos |
|---|---|
| **Variáveis** | Criar e transformar → Calcular · Recodificar em outras / nas mesmas variáveis · Organizar → Renomear · Excluir |
| **Casos** | Filtrar → Selecionar casos · Remover duplicados · Organizar → Ordenar · Resumir → Agregar |
| **Rótulos e ausentes** | Descrever as variáveis → Rótulo da variável · Rótulos dos valores · Nível de medida · Valores ausentes |
| **Juntar bancos** | Combinar com outro .sav → Adicionar casos · Adicionar variáveis |
| **Arquivo** | Salvar e exportar → Salvar .sav · Exportar Excel/CSV · Salvar syntax |
| **Ajuda** | Funções e operadores · Como usar este módulo |

- **Desfazer** e **Desfazer tudo** ficam sempre à vista, acima das abas.
- Cada comando abre uma **janela** (`st.dialog`) com **OK · Ver syntax · Reiniciar · Cancelar**. “Ver syntax” mostra o
  comando sem aplicar; “Reiniciar” recria os campos em branco.
- **Calcular variável** e **Selecionar casos** têm o editor de expressão: lista de variáveis, grupos de funções e
  teclado de operadores que inserem no fim da expressão.
- **Recodificar**: linhas de campos “valor antigo → novo valor” (Valor, Intervalo, Do menor valor até…, De… até o maior,
  Ausente do sistema, Ausente, Todos os outros), com o botão “＋ Adicionar linha”, ou texto de syntax. “Todos os outros” vai sempre por último.
- **Valores ausentes**: sem ausentes · até 3 discretos · faixa + 1 discreto. **Rótulos dos valores**: caixa de texto
  com uma linha “código = rótulo”, já preenchida com os rótulos atuais.
- **Abas de baixo** (como o Data View/Variable View do SPSS): *Visualização de dados* (tabela HTML: cabeçalho cinza,
  nº do caso, ausente = “.”) · *Visualização de variáveis* (Nome, Tipo, Rótulo, Valores, Ausentes, Medida, Casos
  válidos; escolha a variável e use Nome/Rótulo/Valores/Ausentes/Medida (botões sem reticências, texto sem quebra) para abrir a janela já nela) · *Syntax*.
- **Barra de status**: arquivo, casos × variáveis, passos aplicados e o último passo.
- **Cores explícitas**: o app força texto escuro e o tema do Streamlit pode ser escuro; por isso o CSS da tela define
  fundo *e* texto juntos nas faixas, nas grades, nas abas, nos expanders e nas janelas.
- Requisitos do Streamlit: `st.dialog`, `st.segmented_control`, `st.button(icon=...)` e `st.rerun(scope="fragment")`.
  Em versões ≥ 1.43 os downloads não recarregam a tela.

## Operações (cada uma gera a syntax SPSS)
| Grupo | Operação | Syntax gerada |
|---|---|---|
| Variáveis | Calcular (com condição opcional) | `COMPUTE` / `IF` |
| | Recodificar | `RECODE ... [INTO]` |
| | Renomear · Excluir variáveis | `RENAME VARIABLES` · `DELETE VARIABLES` |
| Casos | Manter/excluir casos por condição | `SELECT IF` / `SELECT IF NOT` |
| | Ordenar · Remover duplicados | `SORT CASES` · `MATCH FILES /FIRST` + `SELECT IF` |
| Rótulos e missings | Rótulo da variável · Rótulos dos valores | `VARIABLE LABELS` · `VALUE LABELS` / `ADD VALUE LABELS` |
| | Valores ausentes · Nível de medida | `MISSING VALUES` · `VARIABLE LEVEL` |
| Arquivo | Adicionar casos · Adicionar variáveis | `ADD FILES` · `MATCH FILES` (`/FILE` ou `/TABLE`) |
| | Agregar | `AGGREGATE` (`MODE=ADDVARIABLES` ou banco agregado) |

## Regras do SPSS reproduzidas
- Missing de usuário vira **ausente** nas contas; `MISSING()` os enxerga, `VALUE()` devolve o código bruto.
- Lógica de 3 valores: comparar com ausente não é verdadeiro nem falso; `IF`/`SELECT IF` só agem onde é verdadeiro.
  `SELECT IF NOT` também remove os casos com ausente na condição (a tela avisa quantos).
- `RECODE`: vale a primeira regra que casar; `LO`/`HI` incluem missing de usuário e não o sysmis; `ELSE` inclui o
  sysmis; em **variável nova** o que não foi citado vira sysmis, em variável existente fica como estava.
- Funções com `TO` (`MEAN(Q1 TO Q5)`) seguem a ordem das colunas do arquivo; `MEAN.2(...)` exige 2 válidos.
- Ordenação: ausente é o menor valor (primeiro no crescente, último no decrescente).
- `MISSING VALUES`: até 3 valores discretos, ou 1 faixa + 1 valor (limite do formato .sav).

## Saídas
`.sav` (rótulos, missings e medida), `.sps` (todos os passos, de `GET FILE` a `SAVE OUTFILE`), Excel (aba Dados +
aba Variáveis) e CSV (`;`). Excel/CSV podem trocar códigos por rótulos.

## Limites conhecidos
- A tela nova **não foi aberta no Streamlit real** (só testada com um `streamlit` falso): o visual (CSS da barra de
  faixas, das grades e da barra de status) pode precisar de ajuste fino.
- As grades de dados e de variáveis são somente leitura (HTML); a edição é feita pelas janelas. Dentro das janelas
  (Recodificar, Rótulos, Agregar) não há mais `st.data_editor` (ele é desenhado em canvas e seguia o tema escuro):
  são linhas de campos comuns, com CSS claro explícito (janela localizada por `[role="dialog"]`).
- As expressões são lidas por um analisador próprio (sem `eval`). Não há funções de data nem `STRING()`; use `CONCAT`.
- No escritor em Python puro (usado quando o pyreadstat não carrega), o `.sav` sai sem compressão e variáveis de data
  são gravadas como número de segundos (aviso na tela). Texto vazio no `.sav` é "em branco", não nulo.
- `MATCH FILES` por chave exige chaves únicas no segundo banco; a syntax gerada lembra de ordenar os dois arquivos.
- Gravação pelo **pyreadstat** (Streamlit Cloud) usa `missing_ranges` e `variable_measure`; esse caminho não foi
  exercitado nos testes (o pyreadstat não estava instalado no ambiente de desenvolvimento). Se ele recusar algum
  parâmetro, o módulo cai sozinho para o escritor alternativo e avisa.
- O registro de medida (subtipo 11) gravado pelo escritor puro segue a descrição pública do formato (PSPP); vale abrir
  um `.sav` gerado no SPSS real ao menos uma vez.
