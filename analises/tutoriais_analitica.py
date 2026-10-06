"""
Tutoriais das ferramentas analíticas (Mapa de Calor, Cruzamentos, Matriz
de Correlação, Análise Residual).

São os guias do app desktop "Relatória Analítica" (assets/tutorial/*.html)
com os mesmos passos e nomes de botões da tela original, que agora roda
dentro do Streamlit; só mudam os pontos próprios do navegador (seletor de
arquivo, download, avisos).

`botao_tutorial(nome)` mostra o guia recolhido no topo de cada tela (o
botão "?" do original); `modulo_tutoriais()` é a página "Tutoriais" do menu.
"""
from __future__ import annotations

import os

import streamlit as st

_ASSETS = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "tutoriais")

TUTORIAIS = {
    # Liberados um por vez (Cruzamentos primeiro); os demais entram junto
    # com a ferramenta correspondente.
    "mapa_calor": {
        "titulo": "Mapa de Calor",
        "resumo": "Importar o relatório e os KMLs, escolher caso e segmento, conferir a base, pintar, numerar e exportar PPTX.",
        "imagem": ("mapa_calor_exemplo.jpg", "Exemplo de resultado esperado"),
        "texto": """
**Pré-requisitos**
- *Relatório Descritivo Quantitativo* (Excel) do projeto.
- **KMLs** das regiões, com nomes **idênticos** aos do relatório. Se não baterem, renomeie.

> 💡 Se uma região não pinta, quase sempre é diferença de nome entre o KML e o relatório.

#### 1) Importar o relatório (Excel) e os KMLs
- Clique em **Importar Excel** e escolha o Relatório Descritivo Quantitativo no seletor do navegador.
- Em **Importar KML (vários)**, clique em **Escolher arquivos** e selecione **todos** os KMLs de uma vez.
- Logo abaixo aparece a **conferência dos KMLs**: depois de escolher o caso, ela fica **verde** quando todas as regiões do caso têm KML, ou **vermelha** listando as regiões *sem KML* e os KMLs que *não estão no caso*.

#### 2) Configurações visuais
- **Fundo = Satélite (Esri)** e **Opacidade = 83** já vêm selecionados (padrão do processo).
- Ajuste **Espessura das bordas** e **Tamanho dos números** conforme o mapa.

#### 3) Caso
- Escolha o **Caso** que precisa mapear e confira no relatório se é a tabela certa — o nome mostra a aba, a linha do *Header* e da *Base*.

> ✅ Regra de ouro: valide o caso no relatório antes de pintar, para evitar retrabalho.

#### 4) Segmento
- Selecione o **Segmento** (ex.: *Certamente voto*, *Aprovação*) e confirme que é o recorte pedido.

#### 5) Base e cálculos
- A **Base (Total)** normalmente vem do Excel, mas confira. Com base reduzida, digite a base manualmente (**Usar base do Excel** volta ao valor detectado).
- Passe o mouse em **Mín / Média / Máx** para ver o cálculo completo.
- Regiões com **base menor que 30** aparecem num aviso **amarelo** (valores instáveis — avalie antes de usar).

> ⚠️ Base errada gera cores erradas, mesmo com o KML certo.

#### 6) Pintura automática
- Clique em **Aplicar ao mapa**. A conferência fica **verde** (“Todas as N regiões do Excel foram pintadas”) ou **vermelha**, listando as regiões do Excel sem KML e os KMLs sem dado no caso.
- Clique numa região e valide o **Valor** no balão; confira se a cor bate com Mín / Média / Máx.
- As cores de *Baixo / Médio / Alto* podem ser trocadas em **Paleta ativa** (**Restaurar padrão** volta às originais).

#### 7) Numeração
- Em **5) Numeração**, preencha AP1, AP2… conforme a legenda do relatório. Os números podem ser **arrastados** no mapa.
- Cada região também aceita cor manual; **Limpar** tira a pintura de uma região e **Limpar pintura** de todas.

#### 8) Fila e exportação para PPTX
- Clique em **Carregar logo** e escolha a logo (Pesquisa e Dados).
- Com o mapa pronto, clique em **Adicionar à fila**: um aviso confirma. Se o mesmo **caso + segmento** já estiver na fila, o app pergunta se deve **substituir** (útil depois de corrigir a numeração).
- A fila fica na **barra lateral escura** (**Fila do PPTX**): um card por mapa; clique no card para ver a **miniatura**, **Mín / Média / Máx** com as cores e o caso/segmento. A **lixeira** tira um mapa.
- **Exportar PPTX** (barra lateral) baixa **mapas_calor.pptx**; **Limpar fila** recomeça. **Exportar PNG** (na ferramenta) baixa só o mapa atual.
- Trocar de ferramenta no menu e voltar **não perde nada**: KMLs, pintura, numeração, caso e fila são restaurados.

> ✅ Checklist: Pintadas ok • cores conferidas • numeração ok • logo carregada • fila completa.

**Problemas comuns**
- *Região não pinta / “Pintadas” menor:* nome do KML diferente do relatório ou KML faltando.
- *Cores estranhas:* caso/segmento errado ou base incorreta.
- *Números ilegíveis:* aumente o tamanho, reduza as bordas ou arraste o número.
- *Aviso amarelo de base:* a região tem menos de 30 entrevistas; confira com o relatório antes de usar.
""",
    },
    "correlacao": {
        "titulo": "Matriz de Correlação",
        "resumo": "Importar a base, definir o peso, selecionar colunas, tratar missing, ajustar nomes, calcular e exportar CSV/XLSX.",
        "texto": """
**Pré-requisitos**
- Banco do projeto em **CSV, XLSX ou SAV**, com uma **variável de peso**.
- Pelo menos **2 colunas** para montar a matriz.
- Códigos como 98, 99, “não sabe” ou campos vazios devem ser tratados em **Missing**.

> ✅ Regra prática: primeiro importe a base, confirme o peso e revise os missing. Depois ajuste nomes, método e exportação.

#### 1) Importar o banco (0) Dados)
- Clique em **Importar CSV/XLSX/SAV** e escolha a base no seletor do navegador.
- Depois da carga, o nome do arquivo aparece e as colunas numéricas ficam disponíveis. Bases do SPSS trazem os rótulos das variáveis.

#### 2) Definir o peso (1) Peso — obrigatório)
- Se a base tiver **uma única coluna com “peso”** no nome, ela já vem selecionada (aviso verde). Senão, escolha em **Variável de peso**. Sem peso o cálculo não conclui.
- O peso nunca entra como variável da matriz: se ele estiver marcado nas colunas, sai sozinho.

#### 3) Selecionar as colunas (2) Selecionar colunas)
- Marque as variáveis da matriz. **Todos** (verde) marca as colunas da busca e mostra quantas estão selecionadas; **Nenhum** (vermelho) desmarca. As marcadas ficam destacadas em verde (mínimo 2).

> 💡 Comece com poucas variáveis para validar o resultado e depois amplie.

#### 4) Missing (3) Missing — por variável selecionada)
- Escolha a variável (o nome completo aparece abaixo), clique em **Carregar valores**, marque os valores que são missing e clique em **Adicionar missing**.
- **Todos / Nenhum** marcam a lista; **Limpar desta variável** desfaz. O **Missing manual (global)** vale para todas as variáveis (ex.: `98, 99, Não sei`).

#### 5) Nomes (4) Nomes — edição dinâmica)
- Renomeie as variáveis selecionadas aqui **ou** com **2 cliques no cabeçalho** da tabela. *Usar rótulos do SPSS* começa com os rótulos originais.

#### 6) Método e configurações (5) Configurações e cálculo)
- **Método:** Pearson (padrão), Spearman ou Kendall (Kendall com peso vira “sem ponderação”).
- Defina as **casas decimais** e mantenha **Tentar converter texto em número** quando necessário.

#### 7) Calcular e conferir
- Clique em **Calcular**. Só aparecem as correlações com **Sig. entre 0,000 e 0,005**, com as cores do VBA.
- Confira sinais, magnitudes e coerência. Se algo estranhar, revise peso, missing e colunas.

#### 8) Exportar (6) Exportar)
- **Exportar CSV** (`correlacao.csv`, separador `;`) ou **Exportar XLSX** (`correlacao.xlsx`: abas *Correlações* e *Sig. (2-tailed)*, col=10, linha=40, fonte DIN 10, cores do VBA). O arquivo é baixado pelo navegador.
- Sem resultado calculado, a exportação avisa e não gera arquivo.

> 💡 A matriz acompanha a rolagem da página, então fica visível enquanto você mexe no painel. Trocar de ferramenta no menu e voltar mantém colunas, missing, nomes, método e a última matriz.

> ✅ Resultado esperado: matriz consistente com o peso, missing tratados, nomes claros, método/casas definidos e arquivo exportado.
""",
    },
    "cruzamentos": {
        "titulo": "Cruzamentos (Tabelas)",
        "resumo": "Importar o banco, configurar variáveis e peso, gerar tabelas, ordenar linhas/colunas e exportar Excel.",
        "texto": """
**Pré-requisitos**
- Banco do projeto **com peso** (SPSS .sav, CSV ou Excel). O **Peso é obrigatório** no fluxo do projeto.
- Tenha o *Relatório Descritivo Quantitativo* aberto para copiar a ordem certa de linhas e colunas.

> ✅ Regra do processo: sempre usar Peso e só exportar depois de ajustar ordem, remoções e título para ficar igual ao relatório.

#### 1) Importar o banco
- Em **0) Dados**, clique em **Carregar dados (CSV/XLSX/SAV)** e escolha o arquivo no seletor do navegador.
- Aguarde o aviso “Enviando…” sumir e confira se o **nome do arquivo** aparece ao lado do botão.

#### 2) Configurar variáveis e Peso (1) Configurações das Variáveis)
- **Variável (Linhas):** digite para buscar e clique na variável principal (ex.: intenção de voto). O nome completo aparece logo abaixo, em azul.
- **Peso:** selecione **Peso**. Só aparecem colunas com “peso” no nome.
- **Percentual:** deixe *Coluna (%)*, o padrão do projeto.
- **Variáveis (Colunas):** marque as variáveis do cruzamento. *Buscar variável* filtra a lista; **Marcar tudo** (verde) marca as que estão na lista e mostra quantas estão selecionadas; **Desmarcar** (vermelho) desmarca. As marcadas ficam destacadas em verde e aparecem em *Selecionadas* (clique no **×** para tirar uma).
- **Incluir NA** cria a categoria *(Sem resposta)*; normalmente fica desmarcado.

> Se “Peso” não aparecer, revise o nome da coluna: ela precisa conter “peso”.

#### 3) Gerar cruzamentos
- Clique em **Gerar cruzamentos**. Cada tabela aparece num cartão com título, base e botões de ajuste.
- Confira a **Base (N)** antes de ordenar ou remover itens. **Limpar resultados** apaga os cartões.
- Ordem automática das linhas: pelo **Total** (maior → menor). A escala *Certamente voto … Não conheço suficiente para votar* fica no topo, na ordem própria. Respostas fixas da lista herdada do VBA (ex.: *Sim/Não*, *Regular*, *Aprovação/Reprovação*, *Outros*) **não entram na ordenação** e vão para o fim, e *Não sei / Branco / Nulo / Não sabe / NS* ficam por último.

#### 4) Ajustar a tabela: título e ordem
- **Título:** edite direto no campo de título do cartão.
- **Ordenar linhas:** no modal, use **↑ / ↓** para mover cada linha e clique em **Salvar** (deixe igual ao relatório).
- **Colunas removidas automaticamente:** colunas com **base menor que 30** e as respostas **Branco/Nulo**, **Não irá votar** e **Não sabe** já vêm removidas. O cartão mostra uma faixa amarela dizendo quais foram.
- **Ordenar colunas:** no modal, use **↑ / ↓** para a ordem e **Remover** nas colunas **sem base**. As removidas automaticamente aparecem riscadas com a etiqueta *auto*; use **Restaurar** para trazer de volta (ex.: a **coluna do cliente** com base baixa). Clique em **Salvar**. O *Total* fica fixo na coluna B.

#### 5) Adicionar à pilha
- Com a tabela ajustada, clique em **Adicionar à pilha** no cartão. Um aviso confirma e o botão vira **✓ Na pilha**.
- O mesmo cruzamento (mesmas linhas × colunas, peso, percentual e NA) **não entra duas vezes**: aparece um alerta dizendo em que posição ele já está. Para trocar a versão, remova o item na pilha e adicione de novo.
- **Adicionar todos à pilha** só depois de revisar tudo; os repetidos são pulados.

#### 6) Exportar Excel (barra lateral)
- A pilha fica na **barra lateral escura**, em **Pilha de exportação**: um card numerado (01, 02…) por tabela, com o título.
- Clique no card (ou na seta) para ver a **Pergunta base** (linhas) e a pergunta do **Cruzamento** (colunas) por extenso; clique de novo para fechar.
- **Exportar Excel** baixa **cruzamentos.xlsx** com todas as tabelas da pilha numa aba, uma embaixo da outra.
- A **lixeira** vermelha de um card tira só aquela tabela; **Limpar pilha** recomeça.
- **Exportar (1)**, no cartão, baixa só aquela tabela (*cruzamento.xlsx*).

> ✅ Checklist: Peso selecionado • ordem igual ao relatório • colunas sem base removidas • pilha completa.

**Dicas desta versão web**
- Os avisos (antes janelas do Windows) aparecem como faixas no canto superior direito da ferramenta.
- Trocar de página no menu e voltar mantém o banco, os resultados e a pilha.

**Problemas comuns**
- *Peso não aparece:* a coluna precisa conter “peso” no nome e ser numérica.
- *Ordem diferente do relatório:* use Ordenar linhas / Ordenar colunas.
- *Colunas sem base:* remova em Ordenar colunas (botão Remover), exceto a coluna do cliente.
- *Sumiu uma coluna que eu queria:* ela foi removida automaticamente (base < 30 ou Branco/Nulo, Não irá votar, Não sabe). Use Ordenar colunas → Restaurar.
- *“Esse cruzamento já está na pilha”:* ele já foi adicionado; veja na barra lateral.
- *Exportação errada:* remova o item da pilha, corrija a tabela e adicione de novo.
""",
    },
}


def _mostrar(nome: str) -> None:
    t = TUTORIAIS[nome]
    st.markdown(t["texto"])
    img = t.get("imagem")
    if img:
        caminho = os.path.join(_ASSETS, img[0])
        if os.path.exists(caminho):
            st.image(caminho, caption=img[1])


def botao_tutorial(nome: str) -> None:
    """Guia recolhido no topo da tela da ferramenta (o botão “?” do original)."""
    if nome not in TUTORIAIS:
        return
    with st.expander(f"📘 Tutorial — {TUTORIAIS[nome]['titulo']} (passo a passo)"):
        _mostrar(nome)


def modulo_tutoriais() -> None:
    st.header("Tutoriais")
    st.caption("Selecione o tutorial que deseja abrir.")
    nomes = list(TUTORIAIS)
    escolha = st.radio(
        "Tutorial", nomes, key="tut_escolha", horizontal=True,
        format_func=lambda n: TUTORIAIS[n]["titulo"], label_visibility="collapsed",
    )
    st.caption(TUTORIAIS[escolha]["resumo"])
    st.divider()
    _mostrar(escolha)
