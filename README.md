# Informa-Automatizado

Aplicativo web (Python + Streamlit) do **Instituto Informa** que reúne as rotinas de pós-processamento de pesquisas: do banco `.sav` e das programações SPSS até o relatório em Excel formatado para entrega. A ideia central é trocar o fluxo antigo de macros VBA e conferências manuais por assistentes guiados, com a mesma lógica de formatação e menos retrabalho.

O app abre numa **Central de ferramentas** (tela inicial com cartões por seção). Cada cartão leva a uma ferramenta; clicar na marca «informa» no topo de qualquer ferramenta volta para a Central.

## Ferramentas

| Seção | Ferramenta | O que faz |
| --- | --- | --- |
| Programações | Revisor de Syntax SPSS | Audita `.sps` (RECODE, VALUE LABELS, MRSETS…) contra o banco |
| | Tratamento do Banco | Roda a Syntax no `.sav`, confere a recodificação, gera o banco tratado; calcula e confere o Total e o Completo, com painel de missings dos segmentos e exportação em Excel no layout do SPSS |
| Formatação | Total (Total Automático) | Assistente que formata o relatório pelo total |
| | Descritivo (Relatório Automatizado) | Assistente que aplica os códigos 01–17 do antigo VBA no relatório descritivo |
| | Padronização de Títulos | Catálogo do título padrão de cada tipo de pergunta |
| | Divisor de Tabelas de Regiões | Divide as tabelas que passam do limite de altura da página; gera o Índice |
| Analíticos | Relatório Analítico, Análise das Perguntas, Perfil | Conferência do descritivo, perguntas cruzadas por segmento, tabelas de Perfil |
| | Cruzamentos, Mapas de Calor, Correlação, Análise Residual | Ferramentas portadas do app desktop «Relatória Analítica» |
| | Mapa de Correspondência (V2; V1 antigo) | Ver [docs/MAPAS_DE_CORRESPONDENCIA.md](docs/MAPAS_DE_CORRESPONDENCIA.md) |
| Controle | Controle Geral, Gerador de Amostra, Banco de Amostras | Validação geográfica/eleitoral e distribuição da amostra |
| Codificação e ponderação | Ponderação | Cálculo e aplicação dos pesos da amostra |
| Validações | Base Reduzida, Exclusão de Colunas, Exclusões, Base nas Múltiplas, Legendas, Códigos Individuais | Conferências e etapas avulsas de formatação |

Os cartões sem destino (Gerador de Programações e Codificação) aparecem com a marca «Em breve» e entram quando o app externo correspondente for integrado.

## Como rodar

```bash
git clone <URL_DO_REPOSITORIO>
cd Informa-Automatizado
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

O app abre em `http://localhost:8501`. Para rodar os testes:

```bash
python3 -m pytest tests -q
```

Se o `pyreadstat` não carregar (por exemplo, Windows com bloqueio de DLL), a leitura de `.sav` cai automaticamente num leitor em Python puro (`core/sav_puro.py`). Só a gravação de `.sav` precisa do `pyreadstat`.

## Como o código está organizado

```text
app.py              menu, despacho das ferramentas e cabeçalho
analises/           telas (uma ferramenta por arquivo, função modulo_<nome>())
  tela_inicial.py     cartões da Central de ferramentas
  _tema_ferramenta.py tema escuro + cabeçalho das ferramentas (TEMA_ESCURO)
  _assistente_ui.py   stepper, cartões e chip de arquivo dos assistentes
core/               lógica pura, sem Streamlit (*_math.py) — é aqui que se testa
controle/           pacote do Controle Geral
componentes/        páginas HTML das ferramentas portadas do app desktop
assets/ dados/      logo, modelos e bancos auxiliares
tests/              testes (pytest)
docs/               documentação por módulo
```

Convenções que o projeto segue: a regra de negócio fica em `core/` e a tela em `analises/`; cada ferramenta usa um prefixo próprio nas chaves de `st.session_state` (`ra_` Relatório Automatizado, `ta_` Total Automático, `tb_`/`tc_` Tratamento do Banco, `cab_` Codificação de Abertas…); toda ferramenta do `MENU` precisa de um cartão na tela inicial (há teste para isso).

Quer criar ou alterar um módulo? Leia [docs/CONTRIBUINDO.md](docs/CONTRIBUINDO.md).

## Deploy

O app roda no **Streamlit Community Cloud** com `app.py` como ponto de entrada. Restrição importante: **não** coloque `wkhtmltopdf` em `packages.txt` — o pacote não existe no ambiente apt do Streamlit Cloud e já derrubou o app uma vez. A pré-visualização em PDF só aparece quando o programa está disponível.

## Documentação por módulo

- [Mapas de Correspondência](docs/MAPAS_DE_CORRESPONDENCIA.md)
- [Notas dos módulos](docs/NOTAS_DOS_MODULOS.md) — Revisor de Syntax, ferramentas analíticas portadas, leitura de `.sav` sem pyreadstat, Ponderação
- Outros arquivos em `docs/` (Amostra, Banco de Amostras, Controle, Ponderação V6–V9)
