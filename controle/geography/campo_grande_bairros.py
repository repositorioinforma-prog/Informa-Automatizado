"""
Hierarquia oficial de bairros de Campo Grande/MS, conforme a Lei Complementar
que define as Regiões Urbanas e Bairros do município (ANEXO 4 - Mapa das
Regiões Urbanas e Bairros). Fonte: prefeitura de Campo Grande.

Estrutura: Região Urbana -> Bairro oficial -> lista de parcelamentos/loteamentos
(nomes populares/informais) que ficam de verdade DENTRO daquele bairro oficial.

Isso serve para responder com precisão uma pergunta que a Base Brasil sozinha
não responde: quando alguém escreve, por exemplo, "Jardim Ipanema", isso é
uma ÁREA DENTRO do bairro oficial "Sobrinho" (pode ser recodificado com
segurança para o bairro aprovado que representa esse mesmo bairro oficial,
se houver um). Já "Rita Vieira" é, ela mesma, um BAIRRO OFICIAL PRÓPRIO,
não uma área dentro de nenhum outro — então, se não estiver na amostra
aprovada, é honestamente "fora da amostra", e não deve ser silenciosamente
trocado por outro bairro só porque fica na mesma região.

Cada cidade tem sua própria lei; por enquanto só Campo Grande está mapeada.
"""

CAMPO_GRANDE_MS: dict[str, dict] = {
    # ---------------- REGIÃO URBANA DO CENTRO ----------------
    "Centro": {"regiao": "Centro", "parcelamentos": [
        "Cidade", "Vila Alta", "Vila General Wolfgrand", "Vila América", "Vila Ilgenfritz",
        "Vila Clementina", "Jardim Aclimação", "Vila Bartiria",
    ]},
    "São Francisco": {"regiao": "Centro", "parcelamentos": [
        "Vila São Francisco", "Vila Helena", "Vila Anfe", "Vila Cristina", "Vila São Sebastião",
        "Vila Aprazível", "Jardim Cidade", "Jardim Brasil", "Vila Alto das Paineiras",
        "Vila São Thomé", "Vila Capri", "Vila Benjamim", "Nossa Senhora de Fátima",
        "Monte Castelo", "Vila Esplanada", "Vila São Luís", "Cofermat", "Vila Santa Bárbara",
        "Vila Lídia", "Cascudo", "Jardim São Paulo",
    ]},
    "Cruzeiro": {"regiao": "Centro", "parcelamentos": [
        "Cruzeiro", "Clube Campestre Ypê", "Coronel Antonino", "Monte Castelo",
        "Nossa Senhora de Fátima", "Vila Rosa", "Vila Marman", "Vila Gomes", "Vila Célia",
        "Vila Sílvia 2ª Seção", "Coophagrande", "Conjunto Eudes Costa", "Vila Suíça",
        "Conjunto Residencial Monte Castelo", "Coophaban",
    ]},
    "Jardim dos Estados": {"regiao": "Centro", "parcelamentos": [
        "Jardim dos Estados", "Vila Santa Odete", "Vila Guaraciaba", "Vila Tupaceretan",
        "Vila Esportiva", "Jardim Aclimação", "Vila XV de Novembro", "Jardim 7 de Setembro",
        "Vila Mandeta", "Cachoeirinha", "Vila São Jorge", "Cachoeira", "Vila Onze",
        "Vila São Gabriel", "Vila Mariana", "Vila Bernardo Goldman", "Vila Isís", "Vila Abdo",
        "Vila Alto Campo de Marte", "Vila Santério", "Vila da Saúde", "Vila Paulistana",
        "Vila Paraíso", "Vila São Elias", "Vila Lia", "Chácara dos Coqueiros", "Vila Rolim",
        "Vila Santos Gomes", "Vila Suburbano", "Vila Sant'ana", "Vila Maria", "Chácara Cachoeira",
    ]},
    "Bela Vista": {"regiao": "Centro", "parcelamentos": [
        "Jardim Bela Vista", "Vila Costa Lima", "Jardim Santa Catarina", "Chácara Vendas",
        "Jardim Nova Era", "Residencial Village", "Vila Antônio Vendas", "Vila Miguel Couto",
        "Jardim Alegre", "Jardim Ibirapuera", "Chácara Boa Vista", "Villa Di Parma",
    ]},
    "Itanhangá": {"regiao": "Centro", "parcelamentos": [
        "Itanhangá Park", "Jardim Piracicaba", "Vila Gatão", "Vila Rosa Pires",
        "Vila Dr. João Rosa", "Vila Rezende", "Jardim Vista Alegre", "Vila Joselito",
        "Coophamorena",
    ]},
    "São Bento": {"regiao": "Centro", "parcelamentos": [
        "Jardim São Bento", "Vila Nova São Bento", "Jardim Guarujá", "Vila Galvão",
    ]},
    "Monte Líbano": {"regiao": "Centro", "parcelamentos": [
        "Jardim Monte Líbano", "Vila Santo André", "Jardim de Allah",
        "Vila Antônio Inácio de Souza",
    ]},
    "Glória": {"regiao": "Centro", "parcelamentos": [
        "Vila Glória", "Vila Fortuna", "Vila Eva", "Vila Oriente", "Vila Gaspar", "Vila Taveira",
        "Vila Sol Nascente", "Vila Ornelas", "Vila Santa Filomena", "Vila São Miguel",
        "Vila Castelo", "Vila Jardim Alvorada", "Vila Liberdade", "Vila Santa Dorothea",
    ]},
    "Carvalho": {"regiao": "Centro", "parcelamentos": [
        "Vila Carvalho", "Vila Nossa Senhora de Lourdes", "Vila São José", "Vila Santa Maria",
        "Vila Quito", "Vila Carvalho Baís", "Vila São Rafael", "Vila Oliveira",
        "Vila Americana", "Vila Santa Luiza",
    ]},
    "Amambaí": {"regiao": "Centro", "parcelamentos": [
        "Amambaí", "Vila Perseverança", "Vila Maracajú", "Vila Guenka", "Vila São João",
        "Vila Santo Antônio", "Vila Olga", "Vila Barão do Rio Branco", "Vila Aurora",
        "Vila Warde", "Vila Floresta", "Vila Portão de Ferro", "Vila Orpheu Baís",
        "Vila São Vicente", "Cohasmat",
    ]},
    "Cabreúva": {"regiao": "Centro", "parcelamentos": [
        "Cabreúva", "Vila Esplanada 2ª Seção", "Vila Feliciana Carolina", "Vila Santa Rosa",
        "Vila Marisa", "Vila Guarani", "Vila Independência",
    ]},
    "Planalto": {"regiao": "Centro", "parcelamentos": [
        "Vila Planalto", "Vila Soares", "Vila Soares 2ª Seção", "Vila São Manoel",
        "Vila Santa Tereza", "Vila Estephania", "Vila Santa Rosa", "Vila Feliciana Carolina",
        "Vila Alto Sumaré", "Vila Corumbá", "Vila Santos", "Vila Boa Vista", "Monte Verde",
        "Colinas de Campo Grande", "Sky Residence",
    ]},

    # ---------------- REGIÃO URBANA DO SEGREDO ----------------
    "José Abrão": {"regiao": "Segredo", "parcelamentos": [
        "Núcleo Habitacional José Abrão", "Parque dos Laranjais", "Jardim das Paineiras",
        "Vila Oeste", "Manoel Taveira", "Núcleo Parque dos Laranjais", "Jardim das Virtudes",
    ]},
    "Nasser": {"regiao": "Segredo", "parcelamentos": [
        "Vila Nasser", "Vila Nasser 2ª Seção", "Santa Luzia", "Vila Cox",
        "Jardim Alto São Francisco", "Jardim das Acácias", "Vila Lili",
        "Parque Residencial Azaléia", "Jardim Veneza", "Jardim Fluminense",
        "Nossa Senhora das Graças", "Vila Nilza 1ª Seção", "Vila Nilza 2ª Seção",
        "Jardim Paquetá", "Coophasul", "Vila Nossa Senhora Aparecida", "Vila Neuza",
        "Vila Marli", "Vila Novo Horizonte", "Bosque da Saúde", "São Caetano",
        "Jardim Paradiso", "Jardim Monte Alto", "Setvillage I", "Setvillage II",
        "Água Limpa Park", "Residencial Carajás", "Morada dos Deuses",
        "Residencial Alto Tamandaré", "Residencial Recanto do Cerrado", "Bom Retiro",
        "Residencial Tolentino", "Jardim da Mooca",
    ]},
    "Seminário": {"regiao": "Segredo", "parcelamentos": [
        "Jardim Seminário", "Jardim Seminário II", "Vila Santa Lúcia",
        "Vila Jardim Maria Amélia", "Vila Nossa Senhora da Conceição", "Jardim Oracília",
        "Vila Dalila", "Vila Leda", "Vila Antonieta", "Vila São Roque", "Vila Saraiva",
        "Portal do Gramado", "São Benedito", "Lagoa da Cruz", "Vila Lídia", "North Park",
    ]},
    "Monte Castelo": {"regiao": "Segredo", "parcelamentos": [
        "Monte Castelo", "Residencial Vale do Sol I", "Residencial Vale do Sol II",
        "Residencial Vale do Sol III", "Residencial Indaiá", "Jardim São Paulo",
        "Vila São João Bosco", "Residencial Otávio Pécora", "Jardim Bosque de Avilan",
        "Center Park", "Jardim Aruba", "Loteamento Campo Dourado", "Loteamento Costa Verde",
        "Praia da Urca", "Residencial Gabura",
    ]},
    "Mata do Segredo": {"regiao": "Segredo", "parcelamentos": [
        "Jardim das Cerejeiras", "Jardim Campo Novo", "Jardim Presidente", "Jardim Campo Belo",
        "Núcleo das Cerejeiras", "Jardim Nascente do Segredo", "Residencial Gama",
    ]},
    "Coronel Antonino": {"regiao": "Segredo", "parcelamentos": [
        "Coronel Antonino", "Conjunto Residencial Estrela do Sul", "Jardim Imperial",
        "Jardim Mirasol", "Eldorado", "Vila Califórnia", "Vila Triângulo", "Guanabara",
        "Morada Verde", "Conjunto Residencial Nova Olinda", "Jardim Campo Verde",
        "Favela Rio de Janeiro", "Favela Guatambu", "Favela Nacional", "Núcleo Morada Verde",
        "Parque Izabel Garden's", "Jardim Talismã", "Residencial Atlântico Sul",
        "Jardim Barcelona", "Recanto Pantaneiro", "Villa Ravenna", "Villa Ravenna II",
    ]},
    "Nova Lima": {"regiao": "Segredo", "parcelamentos": [
        "Nova Lima", "Jardim Anache", "Jardim Columbia", "Favela Jardim Anache",
        "Jardim Vida Nova", "Loteamento Vida Nova II", "Loteamento Tarsila do Amaral",
        "Loteamento Vida Nova III", "Parque Iguatemi",
    ]},

    # ---------------- REGIÃO URBANA DO PROSA ----------------
    "Autonomista": {"regiao": "Prosa", "parcelamentos": [
        "Jardim Autonomista", "Jardim Autonomista II", "Jardim Autonomista III",
        "Jardim Giocondo Orsi", "Jardim Giocondo Orsi II", "Vila Monte Carlo", "Vila Rica",
        "Vila Taquari", "Vila Cacique", "Vila Pagé", "Jardim Vitrine", "Vila Orsi",
        "Vila Cruzeiro do Sul", "Loteamento Petit Park", "Coophabanco", "Vila Boa Esperança",
        "Coophafé",
    ]},
    "Santa Fé": {"regiao": "Prosa", "parcelamentos": [
        "Santa Fé", "Vila Santos Gomes", "Vila Boa Esperança", "Conjunto Residencial Nova Ipanema",
        "Coophafé", "Vivendas do Bosque", "Royal Park",
    ]},
    "Chácara Cachoeira": {"regiao": "Prosa", "parcelamentos": [
        "Chácara Cachoeira", "Chácara Cachoeira II", "Cachoeirinha", "Jatiúca Park",
        "Vila Miguel Couto", "Vila Miguel Couto 2ª Seção", "Vila Miguel Couto 3ª Seção",
        "Vila Manoel da Costa Lima", "Jardim Umuarama", "Cidade Jardim", "San Marino Park",
        "Nahima Park", "Altos da Afonso Pena",
    ]},
    "Carandá": {"regiao": "Prosa", "parcelamentos": [
        "Carandá Bosque", "Carandá Bosque II", "Carandá Bosque III", "Golden Gate Park",
        "Portal Itayara", "Vila do Polonês", "Tayamã Park", "Vila Nascente",
        "Loteamento Copacabana", "Vivendas do Bosque", "Residencial Itacolomi",
        "Tropical Park", "Residencial Via Park", "Residencial Via Park Itália",
    ]},
    "Margarida": {"regiao": "Prosa", "parcelamentos": [
        "Vila Margarida", "Vila Lucinda", "Vila Catarina", "Vila Catarina II", "Vila Carolina",
        "Jardim Marabá", "Núcleo Marabá",
    ]},
    "Mata do Jacinto": {"regiao": "Prosa", "parcelamentos": [
        "Mata do Jacinto", "Conjunto Mata do Jacinto", "Favela do Limão",
        "Loteamento Abaeté", "Loteamento Sóter",
    ]},
    "Novos Estados": {"regiao": "Prosa", "parcelamentos": [
        "Jardim Montevidéu", "Conjunto Residencial Novo Amazonas",
        "Conjunto Residencial Nova Bahia", "Conjunto Residencial Novo Pernambuco",
        "Conjunto Residencial Novo Rio Grande do Sul", "Conjunto Residencial Novo Maranhão",
        "Conjunto Residencial Novo Sergipe", "Conjunto Residencial Novo São Paulo",
        "Conjunto Residencial Novo Alagoas", "Conjunto Residencial Novo Minas Gerais",
        "Conjunto Residencial Novo Paraná", "Jardim Jacarandá", "Alphaville Campo Grande",
        "Alphaville Campo Grande II", "Alphaville Campo Grande III", "Alphaville Campo Grande IV",
    ]},
    "Estrela Dalva": {"regiao": "Prosa", "parcelamentos": [
        "Jardim Estrela Dalva I", "Jardim Estrela Dalva II", "Jardim Estrela Dalva III",
        "Taquaral Bosque",
    ]},
    "Veraneio": {"regiao": "Prosa", "parcelamentos": [
        "Jardim Veraneio", "Vila Futurista", "Vila Danúbio Azul", "Jardim Tayana",
        "Vila Abdalla", "Desbarrancado", "Jardim Arco-Íris", "Loteamento Bosque da Esperança",
        "Beirute Residence Park", "Bosque da Esperança II",
    ]},
    "Chácara dos Poderes": {"regiao": "Prosa", "parcelamentos": [
        "Chácara dos Poderes", "Jardim Pinheiros", "Jardim Cabral", "Vila Telma", "Vila Raquel",
        "Vila Sônia",
    ]},
    "Noroeste": {"regiao": "Prosa", "parcelamentos": [
        "Jardim Noroeste", "Loteamento Nova Serrana", "Serraville", "Residencial Shalom",
    ]},

    # ---------------- REGIÃO URBANA DO BANDEIRA ----------------
    "Jardim Paulista": {"regiao": "Bandeira", "parcelamentos": [
        "Jardim Paulista", "Vila Progresso", "Vila Progresso 2ª Seção", "Paranaense",
    ]},
    "TV Morena": {"regiao": "Bandeira", "parcelamentos": [
        "Jardim TV Morena", "Jardim Paulista", "Paranaense", "Vila Carlota",
    ]},
    "Vilasboas": {"regiao": "Bandeira", "parcelamentos": [
        "Vila Vilasboas", "Jardim Alegre", "Indiana Park", "Jardim Mansur",
        "Jardim Auxiliadora", "Jardim Ipanema", "Parque Dallas", "Residencial Vila Olímpica",
        "Villas Park Residence", "Amantini Residence",
    ]},
    "São Lourenço": {"regiao": "Bandeira", "parcelamentos": [
        "Jardim São Lourenço", "Jardim Ibirapuera", "Vila Almeida Lima", "Vila Zoe",
        "Vila Antônio Vendas",
    ]},
    "Tiradentes": {"regiao": "Bandeira", "parcelamentos": [
        "Tiradentes", "Tiradentes Suplemento", "Tiradentes 2ª Seção",
        "Residencial Nova Tiradentes", "Jardim Flamboyant", "Jardim Flamboyant II",
        "Desbarrancado", "Parque Residencial Arnaldo Estevão de Figueiredo",
        "Parque Residencial Arnaldo Estevão de Figueiredo II", "Núcleo Tiradentes",
        "Regina", "Residencial Itatiaia", "Jardim Vitória", "Parque Residencial Anhembi",
        "Jardim Cristo Redentor", "Jardim São Judas Tadeu", "Jardim Jerusalém",
    ]},
    "Maria Aparecida Pedrossian": {"regiao": "Bandeira", "parcelamentos": [
        "Parque Residencial Maria Aparecida Pedrossian", "Panorama", "Vivendas do Parque",
        "Jardim Samambaia", "Residencial Oiti", "Núcleo Panorama",
        "Parque Residencial Damha", "Parque Residencial Damha II",
        "Parque Residencial Damha III", "Parque Residencial Damha IV",
    ]},
    "Rita Vieira": {"regiao": "Bandeira", "parcelamentos": [
        "Parque Rita Vieira", "Vila Dom Pedrito", "Vila Morumbi", "Jardim Auxiliadora",
        "Coopharádio", "Chácara José Antônio Pereira", "Jardim Lagoa Dourada",
        "Jardim Nossa Senhora do Perpétuo Socorro", "Jardim Itamaracá", "Jardim Anhanguera",
        "Jardim Águas Vivas",
    ]},
    "Carlota": {"regiao": "Bandeira", "parcelamentos": [
        "Vila Carlota", "Vila Ieda", "Vila Dr. Albuquerque", "Vila Morumbi",
        "Vila Portinho Frederico Pache", "Jardim Itapema",
    ]},
    "Dr. Albuquerque": {"regiao": "Bandeira", "parcelamentos": [
        "Vila Dr. Albuquerque", "Vila Olinda", "Vila Progresso 3º Seção", "Vila Maciel",
    ]},
    "Universitário": {"regiao": "Bandeira", "parcelamentos": [
        "Universitário Seção A", "Universitário Seção B", "Residencial Betaville",
        "Vila Concórdia", "Vila Santo Eugênio", "Jardim Ametista", "Jardim Tropical",
        "Pequena Flor I", "Jardim das Perdizes", "Recanto das Palmeiras", "Jardim Campo Limpo",
        "Núcleo Habitacional Recanto dos Rouxinóis", "Jardim Moema", "Jardim Campina Verde",
        "Jardim Campo Alto", "Jardim Pacaembu", "Jardim Indianápolis", "Vila Julieta",
        "Jardim Antares",
    ]},
    "Moreninha": {"regiao": "Bandeira", "parcelamentos": [
        "Núcleo Habitacional Moreninha I", "Núcleo Habitacional Moreninha II",
        "Núcleo Habitacional Moreninha III", "Loteamento Moreninha IV",
        "Chácara Novo Horizonte", "Jardim Santa Felicidade", "Vila Cidade Morena",
        "Jardim Gramado", "Jardim Nova Capital", "Jardim Nova Jerusalém", "Jardim do Córrego",
        "Paraíso do Lageado",
    ]},

    # ---------------- REGIÃO URBANA DO ANHANDUIZINHO ----------------
    "Taquarussu": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Jardim Taquarussu", "Cohafama", "Vila Santo Afonso", "Vila Afonso Pena",
        "Vila Afonso Pena Júnior", "Vila Itamarati",
    ]},
    "Jockey Club": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Vila Jardim Jockey Club", "Vila Marcos Roberto", "Vila Bom Jesus",
        "Vila Santa Amélia Baís", "Residencial Santa Celina",
    ]},
    "América": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Vila Jardim América", "Vila Valparaíso", "Vila Progresso",
    ]},
    "Piratininga": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Vila Piratininga", "Jardim Nhanhá", "Promorar", "Vila Ipiranga", "Núcleo Piratininga",
        "Vila Getúlia Barbosa", "Vitta Bella",
    ]},
    "Jacy": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Jardim Jacy", "Vila Nova Bandeirantes",
    ]},
    "Guanandi": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Guanandi", "Favela Dona Neta", "Favela Núcleo Guanandi I",
    ]},
    "Aero Rancho": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Loteamento Aero Rancho", "Núcleo Habitacional Aero Rancho",
        "Núcleo Habitacional Aero Rancho II", "Núcleo Habitacional Aero Rancho III",
        "Núcleo Habitacional Aero Rancho IV", "Núcleo Habitacional Aero Rancho V",
        "Loteamento Guanandi II", "Jardim das Hortênsias I", "Jardim das Hortênsias II",
        "Jardim das Hortênsias III",
    ]},
    "Parati": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Jardim Parati", "Jardim Parati II", "Granja Bandeira", "Loteamento Alto da Boa Vista",
        "Jardim das Nações",
    ]},
    "Pioneiros": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Jardim Colonial", "Residencial do Lago", "Vila Adelina", "Vila Maciel",
        "Universitário Seção C", "Vila Jardim Pioneiros", "Vila Santa Branca",
        "Vila Santa Branca 2ª Seção", "Jardim Santa Úrsula", "Recanto das Andorinhas",
        "Jardim das Mansões Universitárias", "Conjunto Habitacional Jardim Anápolis",
        "Jardim Rubiácea", "Jardim Jane", "Jardim Botafogo", "Jardim Morenão",
        "Jardim Vicentino", "Jardim Roselândia", "Residencial Botafogo",
        "Parque Residencial Lisboas", "Granja Bandeira", "Jardim Botânico", "Jardim Botânico II",
        "Jardim Agulhas Negras", "Porto Galo",
    ]},
    "Alves Pereira": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Vila Alves Pereira", "Universitário Seção D", "Jardim Monumento", "Vila Antunes",
        "Parque do Trabalhador", "Jardim Colibri", "Jardim Colibri II",
        "Núcleo Habitacional Universitárias I", "Núcleo Habitacional Universitárias II",
        "Jardim Macapá", "Jardim Nashville", "Vila Clélia", "Residencial Ilhéus", "Cidade Nova",
    ]},
    "Centenário": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Jardim Centenário", "Jardim Centenario", "Jardim Monte Alegre", "Vila Nogueira",
        "Vila Amapá", "Vila Aimoré", "Vila Aimoré II", "Jardim Pênfigo",
        "Residencial Vila Bela", "Residencial Ouro Preto", "Jardim Manaíra",
        "Jardim Monterey", "Jardim Radialista", "Jardim das Princesas I",
        "Jardim das Princesas II", "Vila Áurea",
    ]},
    "Lageado": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Parque do Lageado", "Parque do Sol", "Jardim Colorado", "Parque dos Sabiás",
    ]},
    "Los Angeles": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Jardim Los Angeles", "Jardim Sumatra", "Jardim Morada do Sol", "Jardim Uirapuru",
        "Residencial Terra Morena",
    ]},
    "Centro-Oeste": {"regiao": "Anhanduizinho", "parcelamentos": [
        "Jardim Centro Oeste", "Jardim Marajoara", "Jardim Bálsamo", "Jardim Campo Nobre",
        "Jardim das Macaúbas", "Jardim das Meninas", "Jardim Canguru",
        "Jardim Paulo Coelho Machado", "Parque Novo Século", "Jardim Mário Covas",
        "Varandas do Campo",
    ]},

    # ---------------- REGIÃO URBANA DO LAGOA ----------------
    "Taveirópolis": {"regiao": "Lagoa", "parcelamentos": [
        "Vila Taveiropolis", "Santos Dumont", "Vila Belo Horizonte", "Vila Belo Horizonte 2ª Seção",
    ]},
    "Bandeirantes": {"regiao": "Lagoa", "parcelamentos": [
        "Vila Bandeirantes", "Coophavila", "Vila Jurema",
    ]},
    "Caiçara": {"regiao": "Lagoa", "parcelamentos": [
        "Caiçara", "Vila dos Marimbas", "Vila Jardim Anahy", "Vila Jardim Anahy 2ª Seção",
        "Vila Maringá", "Jardim Leblon",
    ]},
    "União": {"regiao": "Lagoa", "parcelamentos": [
        "Parque Residencial União", "Parque Residencial União II", "Residencial Oliveira I",
        "Residencial Oliveira II", "Residencial Oliveira III", "Residencial das Flores",
        "Parque Residencial dos Girassóis",
    ]},
    "Leblon": {"regiao": "Lagoa", "parcelamentos": [
        "Jardim Leblon", "Jardim Leblon 2ª Seção", "Vila Jussara", "Conjunto Habitacional Bonança",
        "Jardim Europa", "Coophamat", "Jardim da Lapa", "Jardim Antarctica", "Vila Ouro Fino",
        "Jardim Tatiana", "Vila Ospampas", "Loteamento Bonjardim", "Núcleo Habitacional Buriti",
        "Alto Leblon",
    ]},
    "São Conrado": {"regiao": "Lagoa", "parcelamentos": [
        "Jardim São Conrado", "Jardim Santa Emília", "Vila Major Juares",
        "Residencial Aquárius I", "Residencial Aquárius II",
    ]},
    "Tijuca": {"regiao": "Lagoa", "parcelamentos": [
        "Jardim Tijuca", "Jardim Tijuca II", "Jardim dos Boggi", "Vila Vilma",
        "Jardim São Conrado", "São Pedro", "Jardim Verdes Mares", "Residencial Barra da Tijuca",
        "Residencial Barra da Tijuca II",
    ]},
    "Caiobá": {"regiao": "Lagoa", "parcelamentos": [
        "Portal Caiobá", "Portal Caiobá II", "Rancho Alegre II", "Vila Fernanda",
        "Rivieira Park", "Riviera Park", "Jardim Rancho Alegre I", "Bela Laguna",
    ]},
    "Batistão": {"regiao": "Lagoa", "parcelamentos": [
        "Jardim Batistão", "São Jorge da Lagoa", "Jardim Mato Grosso",
        "Conjunto Residencial Serra Azul", "Lagoa Park", "Jardim Villa Lobos",
        "Jardim Villa Lobos II",
    ]},
    "Coophavila II": {"regiao": "Lagoa", "parcelamentos": [
        "Coophavila II", "Jardim Vila Kellem", "Jardim Vila Kellem 2ª Secção",
        "Jardim Ouro Verde 1ª Secção", "Jardim Ouro Verde 2ª Secção", "Favela Tarumã",
    ]},
    "Tarumã": {"regiao": "Lagoa", "parcelamentos": [
        "Jardim Tarumã", "Conjunto Residencial Tarumã", "Portal das Laranjeiras",
        "Jardim Sol Poente", "Jardim Corcovado", "Vila Jandaia", "Arapongas",
    ]},

    # ---------------- REGIÃO URBANA DO IMBIRUSSU ----------------
    "Sobrinho": {"regiao": "Imbirussu", "parcelamentos": [
        "Vila Sobrinho", "Vila Acrópolis", "Vila Santa Rita", "Vila Rosalina",
        "Vila Nossa Senhora Auxiliadora", "Cophaco", "Parque dos Ipês", "Jardim Leonidia",
        "Coopermat", "Vila Alba", "Vila Espanhola", "Jardim Ipanema", "Vila Duque de Caxias",
        "Vila Cinamomo", "Lar do Trabalhador", "Residencial Parque dos Flamingos",
        "Vila Aviação", "Parque São Domingos", "Vila Oeste",
    ]},
    "Santo Amaro": {"regiao": "Imbirussu", "parcelamentos": [
        "Vila Santo Amaro", "Vila Jardim Beija-Flor", "Parque dos Laranjais", "Manoel Taveira",
        "Santa Carmélia", "Jardim Itapuã", "Coophatrabalho", "Vila Dr. Jair Garcia",
        "Jardim Canadá", "Vila São Marcos", "Vila Almeida 1ª Seção", "Vila Almeida 2ª Seção",
        "Vila Palmira", "Jardim Mandala", "Residencial Sírio Libanês I",
        "Residencial Sírio Libanês II", "Jardim das Virtudes", "Residencial Hugo Rodrigues",
    ]},
    "Santo Antônio": {"regiao": "Imbirussu", "parcelamentos": [
        "Santo Antônio", "Jardim Imá", "Jardim Imá 2ª Seção", "Vila Nova",
        "Vila Doriza", "Jardim Petrópolis", "Vila Bosque da Saudade", "Vila Coutinho",
        "Vila Sílvia Regina",
    ]},
    "Panamá": {"regiao": "Imbirussu", "parcelamentos": [
        "Jardim Panamá", "Jardim Panamá II", "Jardim Panamá III", "Jardim Panamá IV",
        "Jardim Panamá V", "Jardim Panamá VI", "Parque Residencial dos Bancários",
        "Residencial Sagarana", "Jardim Aroeira", "Recanto dos Pássaros", "Jardim do Zé Pereira",
        "Residencial Ana Maria do Couto", "Parque Residencial Bellinate", "Residencial Búzios",
        "Portal do Panamá", "Jardim Mathilde", "Bosque das Araras",
    ]},
    "Popular": {"regiao": "Imbirussu", "parcelamentos": [
        "Nova Campo Grande Bloco 11", "Nova Campo Grande Bloco 12", "Jardim das Reginas",
        "Jardim Petrópolis", "Jardim Sayonara", "Jardim Pantanal", "Jardim Aeroporto",
        "Jardim Itália", "Bosque Santa Mônica", "Bosque Santa Mônica II", "Vila Romana",
    ]},
    "Nova Campo Grande": {"regiao": "Imbirussu", "parcelamentos": [
        "Nova Campo Grande Bloco 01", "Nova Campo Grande Bloco 02", "Nova Campo Grande Bloco 03",
        "Nova Campo Grande Bloco 04", "Nova Campo Grande Bloco 05", "Nova Campo Grande Bloco 06",
        "Nova Campo Grande Bloco 07", "Nova Campo Grande Bloco 08", "Vila Eliane 1ª Seção",
        "Vila Eliane 2ª Seção", "Vila Serradinho", "Jardim Carioca", "Residencial Nelson Trad",
    ]},
    "Núcleo Industrial": {"regiao": "Imbirussu", "parcelamentos": [
        "Núcleo Industrial", "Jardim Inápolis", "Vila Manoel Secco Thomé", "Vila Entroncamento",
        "Pólo Empresarial Oeste", "Morada Imperial",
    ]},
}
