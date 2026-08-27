# Banco de Amostras v1

## Objetivo

O banco `dados/banco_amostras/banco_amostras.sqlite` consolida os dados territoriais e demográficos utilizados na montagem de amostras. Ele foi criado em SQLite para não acrescentar uma nova dependência ao projeto e para permitir consultas rápidas, auditáveis e reproduzíveis no Streamlit.

O banco passou a ser a **fonte recomendada dentro do aplicativo Streamlit** para os níveis de município e distrito. O Gerador ainda oferece os CSVs anteriores como fallback explícito, permitindo comparar resultados e voltar temporariamente à fonte legada sem remover arquivos.

## Conteúdo

O arquivo contém:

- Brasil, grandes regiões, UFs, regiões intermediárias e regiões imediatas;
- 5.570 municípios;
- 10.698 distritos;
- 643 subdistritos;
- 17.576 registros de bairros;
- 34.487 registros com o bloco demográfico `V01006` a `V01041`;
- 3.400 definições do dicionário de variáveis;
- uma referência municipal de 16+, reaproveitando a idade 16–19 da base mestre do projeto;
- renda de 2010 para municípios e distritos quando disponível;
- procedência, checksums e metadados das fontes;
- validações de cobertura pai → filhos e alertas de qualidade.

## Principais tabelas

| Tabela | Finalidade |
|---|---|
| `territorios` | Hierarquia territorial e chaves não ambíguas. |
| `demografia_agregada` | Valores brutos `V01006` a `V01041`. |
| `indicadores_amostragem` | Referência municipal de 16+ e renda do projeto. |
| `dicionario_variaveis` | Dicionário completo recebido. |
| `cobertura_hierarquia` | Cobertura populacional entre pais e filhos. |
| `alertas_qualidade` | Inconsistências e limitações preservadas. |
| `fontes` | Procedência, checksum, formato e referência temporal. |

## Visões prontas para uso

- `vw_base_amostragem`: visão unificada com território, população, sexo, idade, 16+ e renda;
- `vw_amostra_municipios`;
- `vw_amostra_distritos`;
- `vw_amostra_subdistritos`;
- `vw_amostra_bairros`;
- `vw_alertas_qualidade`;
- `vw_territorios`.

## Chaves territoriais

Os códigos são guardados como texto. Para consultas e seleções no aplicativo, prefira `chave_territorial`:

```text
municipio:3550308
distrito:355030805
bairro:3205200069:1
bairro:3205200069:2
```

A ocorrência final só aparece quando o mesmo código foi repetido na fonte. Isso evita sobrescrever dados ou selecionar a unidade errada.

## Regras importantes

### Valores representados por ponto

O ponto (`.`) encontrado nas fontes significa dado indisponível. Ele foi convertido para `NULL`, nunca para zero.

### População de 16 anos ou mais

- **Municípios:** referência calculada com o total 16–19 do SIDRA e faixas de 20 anos ou mais do agregado. A divisão de 16–19 por sexo é proporcional à composição observada em 15–19.
- **Distritos, subdistritos e bairros:** a fonte só possui a faixa 15–19. A visão calcula uma estimativa explícita usando 80% dessa faixa. O método fica registrado em `metodo_populacao_16_mais_referencia`.

Nenhuma referência 16+ deve ser apresentada sem o campo de método. Nos níveis não municipais, o valor é explicitamente estimado.

### Renda

A renda vem da base de 2010 já utilizada pelo projeto. Quando há zero total, traço ou ausência de correspondência, `renda_disponivel` é zero e as colunas amigáveis de renda retornam `NULL`.

### Fechamento das faixas

Os valores brutos foram preservados. As tabelas de sexo, idade e sexo × idade não fecham exatamente em todas as unidades, possivelmente por diferenças de universo ou tratamento na fonte. As diferenças estão registradas em `qualidade_demografia`; o banco não força fechamento artificial.

## Resultado das validações

- Integridade SQLite: `ok`;
- chaves estrangeiras inválidas: `0`;
- municípios → distritos: todos fecham exatamente;
- unidades sem demografia: 2 distritos, 4 subdistritos e 102 bairros;
- bairros disponíveis para 895 municípios;
- cobertura de bairros exatamente igual ao município em 41 desses 895 municípios;
- subdistritos disponíveis em 77 distritos;
- cobertura de subdistritos exatamente igual ao distrito em 40 desses 77 distritos;
- renda indisponível ou incompleta: 6 municípios e 535 distritos;
- um código de bairro duplicado foi preservado em dois registros:
  - `São Toquarto`, população 1.632;
  - `São Torquato`, população 2.444;
  - código recebido: `3205200069`.

## Uso pelo código Python

```python
from core.banco_amostras import carregar_base_amostragem

municipios_sp = carregar_base_amostragem(
    "municipio",
    uf="SP",
)

bairros_vila_velha = carregar_base_amostragem(
    "bairro",
    codigo_municipio="3205200",
)
```

Para localizar chaves antes de carregar os dados:

```python
from core.banco_amostras import buscar_territorios

resultado = buscar_territorios(
    "bairro",
    texto="São Tor",
    codigo_municipio="3205200",
)
```

## Reconstrução

O script `scripts/construir_banco_amostras.py` reconstrói o arquivo a partir das fontes. Exemplo:

```bash
python scripts/construir_banco_amostras.py \
  --municipios dados_brutos_banco_amostras/Agregados_por_municipios_demografia_BR.xlsx \
  --distritos dados_brutos_banco_amostras/Agregados_por_distritos_demografia_BR.csv \
  --subdistritos dados_brutos_banco_amostras/Agregados_por_subdistritos_demografia_BR.xlsx \
  --bairros dados_brutos_banco_amostras/Agregados_por_bairros_demografia_BR.xlsx \
  --dicionario dados_brutos_banco_amostras/dicionario_de_dados_agregados_por_setores_censitarios_20260520.xlsx \
  --master-municipios dados/gerador_amostra/master_municipios.csv \
  --master-distritos dados/gerador_amostra/master_distritos.csv \
  --saida dados/banco_amostras/banco_amostras.sqlite
```

## Próxima integração recomendada

Município e distrito já alimentam o motor por meio de `carregar_base_motor_amostragem()`. Bairro e subdistrito continuam bloqueados no cálculo até que o fluxo exija uma decisão explícita sobre `cobertura_hierarquia`, especialmente nos territórios com cobertura parcial.

Consulte também `docs/INTEGRACAO_BANCO_STREAMLIT_V2.md`.
