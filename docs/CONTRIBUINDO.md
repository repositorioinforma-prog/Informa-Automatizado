# Guia para quem cria ou altera módulos

Resumo do jeito de trabalhar neste projeto, para os módulos que os colegas mantêm (Ponderação, Controle, Divisor…) se encaixarem sem quebrar os demais.

## Antes de começar

1. Rode os testes (`python3 -m pytest tests -q`) e confirme que está tudo verde antes de mexer.
2. Ao integrar código de outra pessoa, compare com o repositório atual (diff) em vez de copiar arquivos por cima: outros módulos podem ter sido alterados desde a cópia dela.

## Onde cada coisa vai

| O quê | Onde |
| --- | --- |
| Regra de negócio, cálculos, leitura/escrita de arquivos | `core/<nome>_math.py` (sem `import streamlit`) |
| Tela | `analises/<nome>.py`, função `modulo_<nome>()` |
| Testes | `tests/test_<nome>.py` — lógica testada em `core/`; a tela, só o que for texto/HTML puro |
| Documentação do módulo | `docs/<NOME>.md` (se houver regra que não é óbvia pelo código) |

## Registrar uma ferramenta nova

1. **`app.py`**: adicione o nome em `MENU` e o `if analise == "<nome>": modulo_<nome>()` no despacho.
2. **`analises/tela_inicial.py`**: crie o cartão (título, descrição, ícone, `alvo` = nome no `MENU`). Há teste que falha se uma ferramenta do `MENU` ficar sem cartão.
3. **Tema escuro (recomendado)**: inclua a ferramenta em `TEMA_ESCURO` (`analises/_tema_ferramenta.py`) com subtítulo, ícone SVG e cor da seção. O cabeçalho (título + marca «informa» que volta para a Central) passa a ser desenhado pelo app — não use `st.header` nem botão de voltar na tela.
4. **Estado**: use um prefixo único de 2–3 letras nas chaves de `st.session_state` e em `key=` de widgets (veja os já usados no README).

## Assistentes em etapas

Quem for fazer um fluxo em etapas (como o Total Automático e o Relatório Automatizado) pode usar `analises/_assistente_ui.py`:

- `mostrar_stepper(etapas, atual, nomes_curtos)` — pílulas com barra de progresso;
- `mostrar_boas_vindas(...)` / `mostrar_cartoes(...)` — cartões explicando o que o fluxo faz;
- `mostrar_arquivo(nome)` — chip com o arquivo em andamento.

Convenções de tela: botão principal da etapa com `type="primary"`; download final em destaque no topo da tela de conclusão; pré-visualização em HTML dentro de `html_sobre_fundo_branco(...)` para ficar legível no tema escuro.

## Regras de ouro

- **Nunca** adicione `wkhtmltopdf` a `packages.txt` (derrubou o deploy).
- Não adicione dependência em `requirements.txt` sem importar de fato no código; o deploy é mais rápido com poucas bibliotecas.
- Toda alteração de regra precisa de teste: use dados sintéticos pequenos (os testes existentes mostram como montar `.sav`/Excel de exemplo).
- Excel gerado pelo app deve ser conferido relendo o arquivo (e, para layout, abrindo no Excel/LibreOffice) — não só checando valores de células.
- Trabalho em Excel com `openpyxl`: leia `docs/` e o código de `core/planilha_utils.py` antes de mexer em bordas, mesclagens e larguras (há calibração empírica para a fonte DIN).
