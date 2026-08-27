# Metas manuais por unidade — v5

## Objetivo

A distribuição manual passou a representar a meta bruta de cada unidade territorial selecionada. O usuário não edita gênero, idade ou renda manualmente. Ele informa somente quantas entrevistas cada município, distrito ou região deve receber; o aplicativo recalcula os segmentos demográficos dentro desse total.

## Regra de escolha do nível editável

O aplicativo usa o nível mais detalhado entre:

- o nível de saída configurado;
- as Regiões Intermediárias selecionadas explicitamente;
- as Regiões Imediatas selecionadas explicitamente;
- os municípios selecionados explicitamente;
- os distritos selecionados explicitamente.

Exemplo:

```text
Nível de quebra configurado: Região Intermediária
Municípios selecionados: 5
Método: Metas manuais por unidade
```

O editor será aberto por **Município**, com cinco linhas. O resultado não será condensado em uma única Região Intermediária.

## Fluxo operacional

1. Selecione o território e as unidades específicas.
2. Informe a amostra total, por exemplo 500.
3. Escolha **Metas manuais por unidade**.
4. Clique em **Definir metas por unidade**.
5. O sistema apresenta a sugestão proporcional para cada unidade.
6. Edite somente a coluna **Meta da unidade (n)**.
7. A soma das linhas deve fechar exatamente em 500.
8. Ao aplicar, as cotas de gênero, idade e renda são recalculadas dentro da meta de cada unidade.

Exemplo:

```text
Caieiras                 100
Cajamar                   90
Francisco Morato         110
Franco da Rocha          120
Mairiporã                  80
TOTAL                     500
```

Para Caieiras, o sistema calcula os segmentos sobre 100 entrevistas. Para Franco da Rocha, calcula sobre 120, e assim por diante.

## Modo por camadas

Quando o nível específico selecionado é mais detalhado que a última camada configurada, ele é acrescentado como último nível do cálculo manual.

Exemplo:

```text
Camadas configuradas: UF → Região Intermediária
Municípios específicos: Caieiras, Cajamar e Mairiporã
```

No modo manual, o cálculo efetivo passa a ser:

```text
UF → Região Intermediária → Município
```

As metas são editadas nos municípios e os totais das camadas superiores são reconstruídos pela soma dos filhos.

## Faixa mínimo/máximo

A faixa de referência continua disponível internamente com o padrão de 20%, mas o controle foi retirado da interface. Ela não altera a amostra principal nem a soma das metas e, por isso, não precisa ser configurada pelo usuário final.

## Excel e auditoria

A aba **Informações** registra:

- método **Metas manuais por unidade**;
- nível em que as metas foram editadas;
- nível inicialmente configurado, quando diferente;
- sugestão proporcional;
- amostra aplicada;
- diferença em relação à sugestão;
- fechamento da amostra total;
- cotas recalculadas dentro de cada unidade.
