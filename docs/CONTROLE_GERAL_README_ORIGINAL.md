# Gerador de Controle Geral — MVP 2

Aplicativo em Python/Streamlit para criação e auditoria de variáveis de controle em bancos de pesquisas quantitativas.

## O que esta versão faz

- importa banco SPSS `.sav`;
- lê todos os blocos `VALUE LABELS` de uma syntax `.sps`;
- permite configurar livremente ID, novas variáveis e variáveis-fonte;
- permite marcar a nova variável como Município, Bairro, Distrito ou não geográfica;
- lê uma Base Brasil em Excel com municípios, bairros e distritos;
- identifica uma localidade primeiro na Base Brasil e só depois verifica se ela é válida nos `VALUE LABELS` do projeto;
- permite usar uma resposta de bairro para descobrir o município correspondente;
- diferencia `NÃO IDENTIFICADO` de `FORA DA AMOSTRA`;
- não recodifica automaticamente uma localidade conhecida que esteja fora da amostra;
- gera auditoria com texto, fonte, método, confiança, localidade identificada, município e UF;
- verifica coerência Cidade × Bairro/Distrito quando essas variáveis de controle forem configuradas;
- exporta banco provisório em XLSX, auditoria em CSV, inconsistências geográficas em CSV e syntax de VALUE LABELS.

## Princípio de funcionamento

A aplicação separa três coisas:

1. **Resposta** — vem do banco da pesquisa.
2. **Significado geográfico** — vem da Base Brasil.
3. **Validade no projeto** — vem dos `VALUE LABELS` da syntax enviada.

A Base Brasil nunca é usada para ampliar automaticamente a amostra do projeto.

## Base Brasil

Para que a base seja interna ao aplicativo, coloque a planilha em `data/` com um dos nomes:

- `Cidades_Bairros_Distritos - Brasil.xlsx`
- `Cidades_Bairros_Distritos - Brasil(1).xlsx`

Se o arquivo não estiver nessa pasta, a interface permite enviá-lo manualmente durante a sessão.

O leitor aceita múltiplas abas e tenta reconhecer nomes usuais de colunas equivalentes a:

- UF;
- Município / Cidade;
- Código IBGE do Município;
- Distrito;
- Código do Distrito;
- Bairro / Localidade.

## Exemplo importante

Se o banco tiver:

- Cidade respondida: `Rio de Janeiro`
- Bairro respondido: `Olavo Bilac`

E a Base Brasil identificar `Olavo Bilac` como bairro de `Duque de Caxias/RJ`, o sistema pode:

- sugerir/gerar `Duque de Caxias` para uma variável de controle do tipo Município, se esse município estiver nos VALUE LABELS do projeto;
- marcar `FORA DA AMOSTRA` caso Duque de Caxias não esteja nos VALUE LABELS;
- apontar a incoerência Cidade × Bairro em relatório separado.

## Instalação

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Testes

```bash
pytest -q
```

A suíte atual cobre parser de VALUE LABELS, matching, carregamento da Base Brasil, derivação de município a partir de bairro, bloqueio de localidade fora da amostra e coerência geográfica.

## Próxima etapa

A tela de Revisão já mostra os casos pendentes, mas a persistência de decisões manuais ainda não é aplicada ao banco final. A próxima camada deverá implementar:

- aceitar sugestão;
- escolher manualmente outro valor válido;
- manter sem recodificação;
- marcar para exclusão;
- registrar decisão e motivo na auditoria;
- gerar o `.sav` final preservando os originais.
