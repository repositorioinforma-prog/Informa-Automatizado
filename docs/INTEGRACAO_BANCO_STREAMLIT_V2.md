# Integração do Banco ao Streamlit — v2

## Decisão arquitetural

A aplicação usa três formatos com papéis diferentes:

- **SQLite:** fonte oficial dos dados territoriais, demográficos, renda, cobertura, procedência e alertas;
- **JSON:** parâmetros portáteis de cada projeto de amostra;
- **Excel:** saída operacional para revisão, campo e compartilhamento.

O JSON não contém tabelas populacionais. Ele guarda apenas escolhas como território, amostra total, base populacional, modo e camadas. Isso mantém os projetos pequenos e evita múltiplas cópias divergentes do banco.

## Experiência no aplicativo

O menu **Banco de Amostras** oferece:

- resumo por nível territorial;
- verificação de integridade;
- busca por município, distrito, subdistrito e bairro;
- detalhes demográficos e método da referência 16+;
- cobertura de município → distrito/bairro e distrito → subdistrito;
- alertas de qualidade;
- fontes, checksums e referências temporais;
- download do SQLite para backup.

O menu **Gerador de Amostra** oferece:

- `Banco integrado (SQLite — recomendado)` como fonte padrão;
- `Bases anteriores em CSV (fallback)` como contingência explícita;
- salvamento do desenho em JSON;
- abertura de um JSON para restaurar parâmetros;
- recorte direto por Região Intermediária, Região Imediata, município e distrito;
- cálculo clássico e por camadas usando a mesma camada matemática;
- cotas de renda em branco quando a cobertura não é completa, em vez de zeros artificiais.

## Arquivos principais

```text
analises/banco_amostras_ui.py   interface de consulta do banco
analises/gerador_amostra.py     seleção de fonte + projetos JSON
core/banco_amostras.py          consultas e adaptação para o motor
core/projetos_amostra.py        contrato e validação dos projetos JSON
core/amostra_math.py            motor clássico com base injetável
core/amostra_camadas.py         motor hierárquico com base injetável
```

## Contrato do JSON

Exemplo simplificado:

```json
{
  "schema_version": "1.2",
  "tipo": "projeto_amostra",
  "nome": "Nacional 2026",
  "fonte_dados": "banco_sqlite",
  "modo": "camadas",
  "amostra_total": 2000,
  "base_populacional": "16_mais",
  "faixa_flex_pct": 20,
  "territorio": {
    "uf": null,
    "regioes_intermediarias": [],
    "regioes_imediatas": [],
    "municipios": [],
    "distritos": []
  },
  "classico": {
    "nivel_quebra": null,
    "incluir_distritos": false
  },
  "camadas": ["uf", "regiao_intermediaria", "municipio"],
  "distribuicao": {
    "metodo": "proporcional",
    "alocacoes_manuais": []
  },
  "exportacao": {
    "exibicao_cotas": "ambos"
  }
}
```

A versão do esquema é validada ao carregar. Projetos `1.0` continuam compatíveis e são migrados em memória para `1.2`. Resultados calculados e campanhas ainda não fazem parte do JSON; o projeto restaura parâmetros e exige novo cálculo contra a versão atual do banco.

## Limites desta versão

- o banco permanece somente leitura dentro do Streamlit;
- importação e substituição de fontes ainda são feitas pelo script de construção;
- bairros e subdistritos estão disponíveis no explorador, mas não no motor de amostra;
- o modo por camadas oferece distribuição proporcional ou edição manual no último nível;
- persistência local depende do ambiente de hospedagem, mas o download do JSON e do SQLite não depende dela.

## Próxima etapa sugerida

Implementar importação administrativa com validação, relatório de impacto, backup automático e substituição atômica do SQLite. Essa etapa deve ficar protegida de usuários comuns e não deve editar o banco ativo linha a linha.
