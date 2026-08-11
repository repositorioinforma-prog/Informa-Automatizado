# App Analítico — Instituto Informa

Aplicação web em **Python 3.12 + Streamlit** para apoiar rotinas de análise, preparação e finalização de pesquisas. O projeto reúne, em uma única interface, módulos de **Análise de Correspondência**, tratamento de relatórios Excel, equilíbrio de cotas por exclusão, ponderação, geração de amostras e automações que substituem etapas antes executadas manualmente em Excel/VBA.

> O projeto trabalha principalmente com arquivos `.xlsx`, `.sav` e `.csv`. Alguns módulos preservam a formatação original do Excel usando `openpyxl`; outros trabalham com dados tabulares usando `pandas` e `pyreadstat`.

---

## Funcionalidades disponíveis

O menu lateral atual do aplicativo contém os seguintes módulos:

| Módulo | Situação | Entrada principal | Saída principal |
|---|---|---|---|
| Mapas de Correspondência | Funcional | CSV, XLSX ou SAV | Visualizações estática e interativa |
| Mapas de Correspondência (Múltiplas Variáveis) | Funcional | CSV, XLSX ou SAV | Vários mapas + PowerPoint |
| Processador de Base Reduzida | Funcional | XLSX | XLSX consolidado + ZIP com tabelas |
| Exclusões | Funcional | XLSX ou SAV | CSV de IDs + syntax SPSS `.sps` |
| Base nas Múltiplas | Funcional | 2 arquivos XLSX | Relatório XLSX com bases inseridas |
| Legendas | Funcional | 2 arquivos XLSX | Relatório XLSX com legendas |
| Relatório Automatizado | Funcional | XLSX + arquivos auxiliares quando necessários | Relatório XLSX processado |
| Códigos Individuais | Funcional | XLSX | XLSX processado por uma etapa específica |
| Gerador de Amostra | Funcional | Bases mestre internas | Plano amostral XLSX |
| Ponderação | Funcional | SAV + universo manual ou importado | XLSX, `.sps`, SAV ponderado e relatório técnico |
| Tratamento de Dados | **Reservado / sem fluxo implementado nesta versão** | — | — |

---

## 1. Mapas de Correspondência

### Correspondência simples

Permite analisar a associação entre **duas variáveis categóricas**.

Fluxo principal:

1. Carrega um arquivo CSV, XLSX ou SPSS/SAV.
2. Seleciona duas variáveis categóricas.
3. Monta a tabela de contingência com `pandas.crosstab`.
4. Calcula a Análise de Correspondência por SVD usando o motor de `core/ca_math.py`.
5. Exibe a inércia explicada pelas dimensões.
6. Gera mapa estático com Matplotlib.
7. Gera mapa interativo com Plotly.
8. Permite editar nomes, deslocamentos e visibilidade dos rótulos.

A proximidade entre categorias deve ser interpretada em conjunto com a inércia explicada e com o contexto da pesquisa. O mapa mostra **associação**, não causalidade.

### Correspondência com múltiplas variáveis

Permite escolher uma **variável principal** e cruzá-la com várias variáveis secundárias em sequência.

Principais recursos:

- geração de um mapa para cada cruzamento;
- edição individual dos mapas;
- configuração de categorias;
- exportação dos mapas para um único arquivo PowerPoint (`.pptx`).

O módulo fica em `analises/correspondencia_multipla.py` e reutiliza o motor matemático de `core/ca_math.py`.

---

## 2. Processador de Base Reduzida

Trabalha diretamente com Workbooks do Excel para preservar a formatação do arquivo original.

O módulo:

- recebe um relatório `.xlsx`;
- identifica e processa tabelas de base reduzida;
- gera um arquivo consolidado;
- pode disponibilizar as tabelas processadas separadamente dentro de um ZIP.

Saídas atuais:

- `Bases reduzidas - Consolidado.xlsx`;
- `Tabelas_processadas.zip`.

---

## 3. Exclusões — Equilíbrio de Cotas

Módulo para equilibrar uma amostra por **remoção de entrevistas**, em vez de ponderação.

Permite:

- carregar `.xlsx` ou `.sav`;
- escolher a variável identificadora do respondente;
- selecionar variáveis de cota;
- informar tamanho esperado da amostra;
- definir tolerância;
- definir base mínima por categoria;
- configurar metas em percentual ou número absoluto;
- importar metas quando aplicável;
- calcular quais entrevistas podem ser excluídas sem violar as restrições definidas.

Regras centrais do motor:

- não reduzir categorias abaixo da base mínima;
- não excluir casos de categorias que já estejam abaixo da própria meta;
- buscar equilíbrio entre as variáveis de cota dentro da tolerância configurada.

Saídas:

- `ids_excluir.csv`;
- `exclusoes.sps` para aplicação no SPSS.

Lógica matemática: `core/exclusoes_math.py`.

---

## 4. Base nas Múltiplas

Algumas tabelas de perguntas de resposta múltipla não recebem automaticamente uma linha de **Base** no relatório exportado.

Este módulo recebe:

1. um relatório com tabelas de perguntas múltiplas;
2. um relatório do mesmo projeto contendo tabelas com bases válidas.

O sistema casa as colunas pelos textos de **grupo + categoria**, e não apenas pela posição, e insere as bases correspondentes no relatório de múltiplas.

Saída:

- `Relatorio_Multiplas_com_Base.xlsx`.

Interface: `analises/base_multiplas.py`  
Motor: `core/base_multiplas_math.py`

---

## 5. Legendas

Automatiza a inserção de legendas em blocos de tabelas segmentadas por território, como regiões, bairros e municípios.

O módulo recebe:

- relatório principal em Excel;
- arquivo de legendas em Excel.

Ele localiza blocos compatíveis, realiza o pareamento e insere a legenda no local esperado do relatório, preservando a estrutura da planilha.

Saída:

- `Relatorio_com_Legendas.xlsx`.

Interface: `analises/legendas.py`  
Motor: `core/legendas_math.py`

---

## 6. Relatório Automatizado

Assistente para finalizar relatórios descritivos em Excel, portando para Python/openpyxl uma sequência de rotinas originalmente executadas em VBA.

A versão atual possui as rotinas numeradas **01 a 16** integradas ao fluxo, incluindo etapas de:

- organização e ordenação das tabelas;
- formatação de rótulos;
- layout e base reduzida;
- linhas, espaçamentos e quebras de página;
- inserção de legendas;
- ajustes de bases e perguntas;
- correção de cabeçalhos repetidos;
- cabeçalho/rodapé de impressão;
- inserção de capas de resultados.

Algumas etapas pedem arquivos auxiliares, por exemplo:

- arquivo de legenda;
- arquivo de referência para correção de cabeçalhos.

O processamento mantém um **log de etapas** e oferece pré-visualização antes do download.

Saída padrão:

- `processado_<nome_do_arquivo>.xlsx`.

Interface: `analises/relatorio_automatizado.py`  
Motor principal: `core/relatorio_automatizado_math.py`

Módulos auxiliares importantes:

- `core/cabecalho_correcao_math.py`;
- `core/cabecalho_imagem.py`;
- `core/capas_resultados_math.py`;
- `core/planilha_utils.py`;
- `core/legendas_math.py`.

### Observação sobre AutoFit

O `openpyxl` não reproduz o motor visual do Excel com total fidelidade. Em etapas que dependem de AutoFit real, o projeto usa estimativas de altura baseadas no conteúdo das células. Por isso, relatórios muito específicos podem exigir conferência visual final no Excel.

---

## 7. Códigos Individuais

Tela de apoio para executar **uma única rotina** do Relatório Automatizado isoladamente.

É útil para:

- testar uma etapa específica;
- depurar uma rotina em um arquivo real;
- validar o resultado sem executar todo o assistente.

A lógica não é duplicada: a interface chama as mesmas funções usadas pelo Relatório Automatizado.

---

## 8. Gerador de Amostra

Gera planos amostrais a partir da hierarquia territorial oficial utilizada nas bases internas do projeto.

Hierarquia disponível no motor:

- Região Intermediária;
- Região Imediata;
- Município;
- Distrito.

O cálculo distribui entrevistas proporcionalmente à população usando o **Método do Maior Resto**, e também calcula cotas demográficas para:

- sexo;
- idade;
- renda.

A interface permite, entre outros parâmetros:

- selecionar UF;
- filtrar regiões;
- incluir ou não distritos;
- definir o nível de quebra territorial;
- informar o total da amostra;
- escolher a base populacional;
- opcionalmente dividir a amostra em campanhas.

Bases utilizadas pelo app:

```text
dados/gerador_amostra/master_municipios.csv
dados/gerador_amostra/master_distritos.csv
```

Essas bases foram preparadas a partir de fontes do IBGE e são carregadas pelo motor `core/amostra_math.py`.

Saída:

- `Amostra_<UF-ou-Brasil>_<N>.xlsx`.

### Reconstrução das bases mestre

O script:

```text
scripts/amostrador_build_master.py
```

serve para reconstruir os arquivos mestre quando as fontes territoriais/demográficas forem atualizadas. Os arquivos brutos do IBGE não fazem parte do fluxo normal do aplicativo e devem ser obtidos novamente quando houver necessidade de reconstrução.

---

## 9. Ponderação

Módulo para criação e validação de pesos em bases SPSS.

Métodos disponíveis no motor:

- **Razão simples**: calcula fatores a partir da relação universo/amostra por categoria e combina os fatores das variáveis selecionadas;
- **Raking**: recalibra iterativamente as margens até tentar atingir os percentuais de universo dentro da tolerância informada.

O fluxo permite:

- carregar uma base `.sav`;
- selecionar variáveis de ponderação;
- informar ou reutilizar distribuições de universo;
- escolher o método;
- configurar tolerância e parâmetros de cálculo;
- validar distribuição antes/depois;
- gerar arquivos de aplicação e auditoria.

Saídas disponíveis no código atual:

- `Ponderacao_<registro>.xlsx`;
- `Calculo_da_Ponderacao_<registro>.xlsx`;
- `Pesos_<registro>.sps`;
- `Base_com_peso_<registro>.sav`;
- `Relatorio_tecnico_<registro>.xlsx`.

O arquivo `assets/modelo_ponderacao.xlsx` é usado pelo módulo como modelo auxiliar de saída.

Interface: `analises/ponderacao.py`  
Motor: `core/ponderacao_math.py`

---

## Formatos de arquivo

### Entrada geral de dados

O carregador compartilhado (`core/dados.py`) aceita:

- `.csv`;
- `.xlsx`;
- `.sav`.

Para CSV, existe tentativa de leitura com codificações alternativas quando necessário. Para SPSS, o projeto usa `pyreadstat` e preserva rótulos de valores para exibição quando disponíveis.

### Excel

Os módulos voltados a relatórios usam principalmente `openpyxl`, pois precisam manter formatação, mesclagens, estilos, impressão e estrutura das planilhas.

### SPSS

Os módulos de ponderação e exclusões usam `.sav` como formato de entrada e podem gerar syntax `.sps` para reprodução das operações no SPSS.

---

## Arquitetura do projeto

```text
Mapa-de-Correspondencia-/
├── app.py                         # Entrada do Streamlit e roteamento do menu
├── requirements.txt               # Dependências Python
├── README.md
│
├── analises/                      # Camada de interface Streamlit
│   ├── base_multiplas.py
│   ├── base_reduzida.py
│   ├── codigos_individuais.py
│   ├── correspondencia_multipla.py
│   ├── correspondencia_simples.py
│   ├── exclusoes.py
│   ├── gerador_amostra.py
│   ├── legendas.py
│   ├── ponderacao.py
│   └── relatorio_automatizado.py
│
├── core/                          # Regras de negócio e motores matemáticos
│   ├── amostra_exportador.py
│   ├── amostra_math.py
│   ├── base_multiplas_math.py
│   ├── ca_math.py
│   ├── cabecalho_correcao_math.py
│   ├── cabecalho_imagem.py
│   ├── capas_resultados_math.py
│   ├── config.py
│   ├── dados.py
│   ├── exclusoes_math.py
│   ├── legendas_math.py
│   ├── planilha_utils.py
│   ├── ponderacao_math.py
│   └── relatorio_automatizado_math.py
│
├── assets/
│   ├── logo.jpg
│   └── modelo_ponderacao.xlsx
│
├── dados/
│   └── gerador_amostra/
│       ├── master_distritos.csv
│       └── master_municipios.csv
│
└── scripts/
    └── amostrador_build_master.py
```

### Separação de responsabilidades

A arquitetura segue uma divisão simples:

- `app.py`: configura página, tema, logo e roteamento;
- `analises/`: interfaces e estados do Streamlit;
- `core/`: cálculos e manipulações reutilizáveis, sem dependência direta do Streamlit na maior parte dos motores;
- `assets/`: recursos visuais e modelos;
- `dados/`: bases locais necessárias para funcionalidades específicas;
- `scripts/`: tarefas de manutenção/reconstrução que não fazem parte do fluxo diário da interface.

---

## Instalação

### Pré-requisitos

Recomendado:

- Python **3.12**;
- `pip` atualizado;
- Windows, Linux ou macOS.

O ambiente virtual encontrado no projeto original foi criado com Python 3.12.10. Não é necessário reutilizar a pasta `.venv` de outra máquina; o ideal é recriar o ambiente localmente.

### 1. Entre na pasta do projeto

```bash
cd Mapa-de-Correspondencia-
```

### 2. Crie um ambiente virtual

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Windows CMD:

```bat
python -m venv .venv
.venv\Scripts\activate.bat
```

Linux/macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Atualize o pip

```bash
python -m pip install --upgrade pip
```

### 4. Instale as dependências

```bash
pip install -r requirements.txt
```

### 5. Execute o aplicativo

```bash
streamlit run app.py
```

O Streamlit normalmente disponibiliza a aplicação em:

```text
http://localhost:8501
```

---

## Dependências principais

O `requirements.txt` atual inclui:

- Streamlit;
- Pandas;
- NumPy;
- Matplotlib;
- Plotly;
- SciPy;
- Scikit-learn;
- Statsmodels;
- OpenPyXL;
- XlsxWriter;
- Pyreadstat;
- python-pptx;
- Prince;
- AdjustText;
- Requests;
- Pillow;
- Sweetviz;
- Seaborn.

Também existem restrições explícitas:

```text
starlette<1.0
setuptools<81
```

---

## Uso rápido

### Para gerar um Mapa de Correspondência

1. Execute `streamlit run app.py`.
2. Escolha **Mapas de Correspondência**.
3. Carregue CSV, XLSX ou SAV.
4. Selecione duas variáveis.
5. Confira a inércia explicada.
6. Ajuste rótulos e posições, se necessário.
7. Explore os mapas estático e interativo.

### Para processar um relatório Excel

1. Escolha **Relatório Automatizado**.
2. Carregue o `.xlsx` principal.
3. Avance pelas etapas do assistente.
4. Forneça arquivos auxiliares quando a etapa solicitar.
5. Confira o log e a pré-visualização.
6. Baixe o arquivo processado.

### Para ponderar uma base SPSS

1. Escolha **Ponderação**.
2. Carregue a base `.sav`.
3. Defina as variáveis e distribuições de universo.
4. Selecione Razão Simples ou Raking.
5. Configure tolerância e demais parâmetros.
6. Confira a validação.
7. Baixe os arquivos finais desejados.

---

## Configurações compartilhadas

O arquivo:

```text
core/config.py
```

centraliza constantes usadas por diferentes módulos, incluindo:

- variáveis sugeridas nos cruzamentos de correspondência;
- variáveis de cota usadas em exclusões;
- nomes amigáveis de variáveis;
- regras de categorias usadas em configurações específicas dos mapas.

Antes de alterar regras recorrentes do questionário, verifique esse arquivo.

---

## Validação e testes

O projeto ainda não possui uma suíte automatizada de testes versionada no ZIP analisado.

Como validação mínima de desenvolvimento, execute:

```bash
python -m compileall -q app.py analises core scripts
```

Para uma validação funcional, recomenda-se também:

1. iniciar o Streamlit;
2. abrir cada módulo afetado pela alteração;
3. executar com um arquivo real de teste;
4. abrir os XLSX gerados no Excel/LibreOffice;
5. validar SAV/SPS no SPSS quando a alteração envolver ponderação ou exclusões;
6. abrir o PowerPoint gerado quando a alteração envolver múltiplos mapas.

---

## Limitações conhecidas

- A opção **Tratamento de Dados** está presente no menu, mas não possui um fluxo funcional implementado no `app.py` desta versão.
- A Correspondência Simples trabalha com um cruzamento por vez.
- A qualidade visual de ajustes equivalentes ao AutoFit do Excel pode variar porque `openpyxl` não possui o mesmo motor de renderização do Excel.
- Algumas rotinas do Relatório Automatizado dependem da estrutura esperada dos relatórios produzidos no fluxo do Instituto Informa.
- Bases e categorias muito pequenas podem gerar mapas de correspondência instáveis ou difíceis de interpretar.
- Alterações na estrutura dos relatórios de origem podem exigir atualização dos parsers em `core/`.
- A reconstrução das bases do Gerador de Amostra depende de arquivos brutos externos do IBGE que não fazem parte do fluxo normal do aplicativo.

---

## Boas práticas para desenvolvimento

- Preserve a separação entre interface (`analises/`) e lógica (`core/`).
- Evite colocar regras matemáticas diretamente nos componentes Streamlit quando elas puderem ser testadas isoladamente.
- Não versione ambientes virtuais (`.venv`) nem caches (`__pycache__`).
- Não coloque credenciais, tokens ou caminhos absolutos no código.
- Valide sempre os arquivos Excel gerados em um aplicativo compatível antes de liberar uma alteração.
- Ao modificar rotinas portadas de VBA, compare o resultado com um arquivo de referência produzido pelo processo antigo.

---

## Dados e privacidade

As bases carregadas pela interface são utilizadas pelo processo da sessão do Streamlit. Ao implantar o aplicativo em servidor, considere:

- controle de acesso;
- política de retenção de arquivos temporários;
- proteção de bases de respondentes;
- logs sem dados pessoais;
- uso de HTTPS;
- limpeza periódica de arquivos temporários.

Evite incluir bases reais de respondentes no repositório.

---

## Instituição

Aplicativo desenvolvido para apoiar as rotinas de análise de dados do **Instituto Informa**.

---

## Licença

A licença de uso, distribuição e modificação do projeto ainda deve ser definida formalmente. Até que isso seja feito, trate o código como de uso interno da organização.
