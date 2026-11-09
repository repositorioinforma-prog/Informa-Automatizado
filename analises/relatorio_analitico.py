"""
Relatório Analítico — tela.

Etapa 1 (pronta): confere o Relatório Descritivo e aponta divergências entre as
tabelas dos TOTAIS e as dos SEGMENTOS (títulos, opções de resposta, texto da
pergunta). Lógica em `core/verificador_relatorio_math.py`; aqui só upload,
botões e mensagens (convenção do projeto: core/ sem streamlit, analises/ com a tela).

Etapa 2 (pronta): configuração do estudo (local, datas, tempo médio, base, cargos,
análises previstas) e geração do esqueleto da apresentação (.pptx) + manifesto
(.json) para a IA preencher os dados nos lugares certos. Lógica em
`core/esqueleto_apresentacao_math.py`.

Etapa 3 (em andamento): preenche o esqueleto com os arquivos do estudo (um envio só, aceita zip ou
arquivos soltos; cada arquivo é identificado pelo conteúdo). Já preenche o Perfil. Lógica em
`core/preenchimento_math.py` e `core/perfil_math.py`.
"""
from __future__ import annotations

import io
import json

import openpyxl
import pandas as pd
import streamlit as st

from core.capa_descritivo_math import ler_capa
from core.esqueleto_apresentacao_math import (
    ANALISES,
    CARGOS,
    ESTADOS,
    GOVERNOS,
    ConfigEstudo,
    compor_objetivo,
    detectar_cargos,
    montar_esqueleto,
    sugerir_base,
    sugerir_nomes,
    sugerir_margem_erro,
    sugerir_objetivos,
)
from core.preenchimento_math import ARQUIVOS_ESPERADOS, classificar, preencher
from core.verificador_relatorio_math import (
    ROTULO_TIPO,
    comparar,
    escolher_aba,
    montar_inventario,
    observacao,
    relatorio_para_xlsx,
    resumo_inventario,
)

_P = "relatorio_analitico_"


def _verificar(uploaded, aba):
    wb = openpyxl.load_workbook(io.BytesIO(uploaded.getvalue()), rich_text=True)
    ws = wb[aba]
    inv = montar_inventario(ws)
    achados = comparar(inv)
    st.session_state[_P + "inv"] = inv
    st.session_state[_P + "achados"] = achados
    st.session_state[_P + "nome"] = uploaded.name
    st.session_state[_P + "xlsx"] = relatorio_para_xlsx(uploaded.name, inv, achados)
    try:
        st.session_state[_P + "capa"] = ler_capa(uploaded.getvalue())
    except Exception:  # a capa é só conveniência: se falhar, o formulário fica em branco
        st.session_state[_P + "capa"] = None


def modulo_relatorio_analitico():
    aba_conferencia, aba_esqueleto, aba_preencher = st.tabs(
        ["1. Conferir o relatório", "2. Esqueleto da apresentação", "3. Preencher com os dados"])
    with aba_conferencia:
        _aba_conferencia()
    with aba_esqueleto:
        _aba_esqueleto()
    with aba_preencher:
        _aba_preencher()


def _aba_conferencia():
    st.write(
        "Envie o **Relatório Descritivo** (Excel). O app compara cada pergunta dos "
        "**resultados pelo total** com as tabelas dos **resultados pelos segmentos** e "
        "aponta o que está diferente: título, opções de resposta e texto da pergunta."
    )

    uploaded = st.file_uploader(
        "Carregar Relatório Descritivo",
        type=["xlsx", "xlsm"],
        key=_P + "upload",
    )
    if uploaded is None:
        return

    try:
        abas = escolher_aba(openpyxl.load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True))
    except Exception as exc:
        st.exception(exc)
        return
    aba = abas[0] if len(abas) == 1 else st.selectbox("Aba com os resultados", abas, key=_P + "aba")

    if st.button("Verificar relatório", key=_P + "verificar"):
        try:
            with st.status("Lendo perguntas e títulos...", expanded=False) as status:
                _verificar(uploaded, aba)
                status.update(label="Verificação concluída", state="complete")
        except Exception as exc:
            st.exception(exc)
            return

    inv = st.session_state.get(_P + "inv")
    achados = st.session_state.get(_P + "achados")
    if inv is None or achados is None:
        return
    if st.session_state.get(_P + "nome") != uploaded.name:
        st.info("O resultado abaixo é de outro arquivo — clique em **Verificar relatório** para atualizar.")

    r = resumo_inventario(inv, achados)
    c1, c2, c3 = st.columns(3)
    c1.metric("Tabelas nos totais", r["tabelas_totais"])
    c2.metric("Tabelas nos segmentos", r["tabelas_segmentos"])
    c3.metric("Achados", r["achados"])

    if not r["tem_secoes"]:
        st.warning(
            "Este arquivo não tem as duas seções ('RESULTADOS PELO TOTAL' e "
            "'RESULTADOS PELOS SEGMENTOS'), então não há o que comparar."
        )
        return
    if not achados:
        st.success("Nenhuma divergência encontrada entre totais e segmentos.")
    else:
        tipos = sorted({a["tipo"] for a in achados})
        escolhidos = st.multiselect(
            "Filtrar por tipo", tipos, default=tipos,
            format_func=lambda t: ROTULO_TIPO.get(t, t), key=_P + "filtro",
        )
        linhas = [
            {
                "Tipo": ROTULO_TIPO.get(a["tipo"], a["tipo"]),
                "Título nos totais": a["titulo_total"],
                "Título nos segmentos": a["titulo_segmento"],
                "Detalhe": a["detalhe"],
                "Onde": a["linhas"],
                "Observação": observacao(a),
            }
            for a in achados if a["tipo"] in escolhidos
        ]
        st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

    st.download_button(
        "Baixar relatório de conferência (Excel)",
        data=st.session_state[_P + "xlsx"],
        file_name=f"Verificacao_{st.session_state.get(_P + 'nome', 'relatorio')}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=_P + "download",
    )


_MODULOS_OPCIONAIS = {
    "reflexao_60x30": "Reflexão sobre a relação 60×30",
    "candidato_gestao": "Candidato foco — percepção da gestão e comparativo dos governos",
    "candidato_perfil_ipv": "Candidato foco — perfil e IPV (deputado)",
    "cav": "Candidato foco — CAV completo (níveis de conhecimento, IPV, correspondência)",
    "cav_enxuto": "Candidato foco — CAV enxuto (só Descritivo)",
}


def _periodo(valor):
    """date_input em modo intervalo devolve tupla com 0, 1 ou 2 datas."""
    if isinstance(valor, (tuple, list)):
        if len(valor) == 2:
            return valor[0], valor[1]
        if len(valor) == 1:
            return valor[0], valor[0]
        return None, None
    return (valor, valor) if valor else (None, None)


def _aba_esqueleto():
    st.write(
        "Configure o estudo e gere o **esqueleto da apresentação**: capa, metodologia, rodapés e "
        "módulos já montados; os dados de resultado ficam marcados como **pendentes** (aviso vermelho "
        "no slide) junto com um **manifesto** que diz de qual arquivo vem cada dado. Nada é estimado "
        "nem preenchido com zero."
    )
    inv = st.session_state.get(_P + "inv")
    achados = st.session_state.get(_P + "achados") or []
    nome = st.session_state.get(_P + "nome") or ""
    if inv is None:
        st.caption("Dica: confira o Relatório Descritivo na aba 1 antes — o app sugere cargos e base "
                   "e leva as divergências para o manifesto. Também dá para gerar sem ele.")
    else:
        st.caption(f"Usando o Relatório Descritivo já conferido: {nome}")

    suf = str(abs(hash(nome)) % 10_000_000)

    def k(n):
        return f"{_P}cfg_{n}_{suf}"

    capa = st.session_state.get(_P + "capa") or {}
    if not capa.get("achou"):
        capa = {}
    mapa_bytes = None
    if capa:
        with st.expander("Dados lidos da capa do Relatório Descritivo", expanded=False):
            lidos = [("Local", capa.get("local") or capa.get("estado")), ("Mês/ano", capa.get("mes_ano")),
                     ("Campo", f"{capa['campo'][0]:%d/%m/%Y} a {capa['campo'][1]:%d/%m/%Y}" if capa.get("campo") else None),
                     ("Tempo médio", f"{capa['tempo_medio']} min" if capa.get("tempo_medio") else None),
                     ("Regiões", capa.get("n_regioes"))]
            st.markdown("  \n".join(f"- **{r}:** {v}" for r, v in lidos if v))
            for av in capa.get("avisos", []):
                st.warning(av)
            mapas = capa.get("mapas") or []
            if mapas:
                escolha = 0
                if len(mapas) > 1:
                    escolha = st.radio("Imagem do mapa a usar no slide de abrangência", range(len(mapas)), horizontal=True,
                                       format_func=lambda i: f"Imagem {i + 1}", key=k("mapa_sel"))
                mapa_bytes = mapas[escolha]["bytes"]
                st.image(mapa_bytes, caption="Mapa das regiões (da capa do Descritivo) — vai para o slide de abrangência",
                         use_container_width=True)
            if capa.get("regioes_secundarias_texto"):
                st.caption(f"Também há uma lista «{capa.get('regioes_secundarias_titulo') or 'regiões adicionais'}» na capa; "
                           "ela não é usada no slide de abrangência.")
        st.caption("Campos abaixo pré-preenchidos com a capa do Descritivo — confira e ajuste se preciso.")

    st.subheader("Estudo")
    c1, c2, c3 = st.columns([3, 1, 2])
    local = c1.text_input("Município / local", value=capa.get("local") or "", key=k("local"))
    uf_ini = capa.get("uf") if capa.get("uf") in ESTADOS else "RJ"
    uf = c2.selectbox("UF", list(ESTADOS), index=list(ESTADOS).index(uf_ini), key=k("uf"))
    modo = c3.radio("Coleta", ["rua", "online"], index=1 if capa.get("modo") == "online" else 0, horizontal=True, key=k("modo"),
                    format_func=lambda m: "Rua (presencial)" if m == "rua" else "Online")

    c1, c2, c3 = st.columns(3)
    periodo = c1.date_input("Período de campo (início e fim)", value=tuple(capa["campo"]) if capa.get("campo") else (),
                            key=k("campo"), format="DD/MM/YYYY")
    tempo = c2.number_input("Tempo médio (minutos)", min_value=0, max_value=120, value=int(capa.get("tempo_medio") or 0),
                            step=1, key=k("tempo"),
                            help="0 = não informado (fica pendente)")
    sug_base = sugerir_base(inv) or 0
    base = c3.number_input("Base (nº de entrevistas)", min_value=0, value=int(sug_base), step=1, key=k("base"),
                           help="Sugerida pela base mais frequente dos totais do Descritivo. 0 = pendente.")

    c1, c2, c3 = st.columns(3)
    sug_margem = sugerir_margem_erro(base) if base else ""
    margem = c1.text_input("Margem de erro", value=sug_margem, key=k(f"margem_{base}"),
                           help="Sugestão calculada a partir da base (95%, amostra aleatória simples). Confirme com o Descritivo.")
    n_reg = c2.number_input("Nº de regiões", min_value=0, max_value=40, value=int(capa.get("n_regioes") or 0), step=1,
                            key=k("nreg"),
                            help="0 = pendente")
    titulo_capa = c3.text_input("Título da capa",
                                value=(capa.get("subtitulo") or "").upper() or "DIAGNÓSTICO POLÍTICO ELEITORAL - PESQUISA QUANTITATIVA",
                                key=k("titulo"))
    regioes = st.text_area("Regiões e bairros (uma região por linha, ex.: «Região 1 (28,49%): Bairro A, Bairro B»)",
                           value=capa.get("regioes_texto") or "", key=k("regioes"), height=120,
                           help="Lido da capa do Descritivo quando disponível; edite se preciso.")

    st.subheader("Cargos e governos")
    pre = detectar_cargos(inv) or ["Presidente", "Governador", "Senador", "Deputado Federal", "Deputado Estadual"]
    cargos = st.multiselect("Cargos pesquisados", CARGOS, default=[c for c in CARGOS if c in pre], key=k("cargos"))
    c1, c2 = st.columns(2)
    governos = c1.multiselect("Governos avaliados", list(GOVERNOS), default=list(GOVERNOS),
                              format_func=lambda g: GOVERNOS[g], key=k("governos"))
    c2.caption("Preencha os nomes para aparecerem nos títulos (ou deixe vazio = pendente).")
    c1, c2 = st.columns(2)
    nomes = sugerir_nomes(inv)
    governador = c1.text_input("Governador", value=nomes["governador"], key=k("gov"),
                               help="Sugerido pelo título de avaliação no Descritivo — confirme o nome.")
    prefeito = c2.text_input("Prefeito", value=nomes["prefeito"], key=k("pref"),
                             help="Sugerido pelo título de avaliação no Descritivo — confirme o nome.")
    st.subheader("Análises previstas")
    cols = st.columns(len(ANALISES))
    analises = {}
    for col, (chave_a, rotulo) in zip(cols, ANALISES.items()):
        analises[chave_a] = col.checkbox(rotulo, value=True, key=k(f"an_{chave_a}"))
    st.caption("Desmarcar uma análise remove os slides que dependem dela. Mapas de calor desligados: a avaliação "
               "dos governos usa o layout lado a lado (Federal + Estadual).")

    st.subheader("Módulos opcionais")
    extras = st.multiselect("Incluir", list(_MODULOS_OPCIONAIS), default=[],
                            format_func=lambda m: _MODULOS_OPCIONAIS[m], key=k("extras"))
    candidato = cargo_foco = gestao = apres = ""
    if any(m in extras for m in ("candidato_gestao", "candidato_perfil_ipv", "cav", "cav_enxuto")):
        c1, c2, c3 = st.columns(3)
        candidato = c1.text_input("Candidato foco", key=k("cand"))
        cargo_foco = c2.selectbox("Cargo que disputa", ["Deputado Estadual", "Deputado Federal", "Senador", "Prefeito",
                                                          "Governador"], key=k("cargo_foco"))
        gestao = c3.text_input("Gestão anterior (nome)", key=k("gestao"))
        apres = st.text_area("Texto de apresentação do candidato (nível 2)", key=k("apres"), height=80)

    st.subheader("Objetivo específico do estudo (objetivo 2 da metodologia)")
    cfg_previa = ConfigEstudo(
        local=local, uf=uf, cargos=cargos, governos=governos, analises=analises,
        modulos_extras=["analises_estrategicas"] + extras, candidato_foco=candidato,
        cargo_foco=cargo_foco or "Deputado Estadual", prefeito=prefeito,
    )
    sugestoes = sugerir_objetivos(inv, cfg_previa)
    por_id = {sg["id"]: sg for sg in sugestoes}
    escolhidos = st.multiselect(
        "Objetivos sugeridos a partir dos arquivos e da configuração (escolha um ou mais)", list(por_id),
        format_func=lambda i: por_id[i]["texto"], key=k("obj_sug"),
        help="Só aparecem objetivos com base no que o Descritivo e a configuração têm. O texto abaixo pode ser editado.")
    if sugestoes:
        with st.expander("De onde vem cada sugestão"):
            for sg in sugestoes:
                st.markdown(f"- **{sg['origem']}** → {sg['texto']}")
    else:
        st.caption("Sem sugestões ainda: confira o Descritivo na aba 1 ou marque cargos, governos e análises acima.")
    texto_base = compor_objetivo([por_id[i]["texto"] for i in escolhidos])
    objetivo = st.text_area("Texto do objetivo específico (editável; vazio = pendente)", value=texto_base, height=100,
                            key=k(f"obj_{abs(hash(texto_base)) % 10_000_000}"),
                            help="Se vier do briefing, cole aqui. Mais de ~350 caracteres reduz a fonte no slide.")
    if len(objetivo) > 350:
        st.caption("Texto longo: o slide reduz a fonte automaticamente; considere escolher menos objetivos.")

    if st.button("Gerar esqueleto da apresentação", type="primary", key=_P + "gerar"):
        ini, fim = _periodo(periodo)
        cfg = ConfigEstudo(
            local=local, uf=uf, modo=modo, campo_inicio=ini, campo_fim=fim, tempo_medio=int(tempo) or None,
            base=int(base) or None, margem=margem, n_regioes=int(n_reg) or None, regioes=regioes,
            titulo_capa=titulo_capa, mes_ano=(capa.get("mes_ano") or "") if not ini else "", idade_minima=capa.get("idade_minima"),
            governador=governador, prefeito=prefeito, objetivo_especifico=objetivo,
            cargos=cargos, governos=governos, analises=analises,
            modulos_extras=["analises_estrategicas"] + extras, candidato_foco=candidato,
            cargo_foco=cargo_foco or "Deputado Estadual", gestao_anterior=gestao, apresentacao_candidato=apres,
        )
        try:
            with st.status("Montando esqueleto...", expanded=False) as status:
                dados, manifesto = montar_esqueleto(None, cfg, inv, achados, mapa_regioes=mapa_bytes)
                status.update(label="Esqueleto pronto", state="complete")
        except ValueError as exc:
            st.error(str(exc))
            return
        except Exception as exc:
            st.exception(exc)
            return
        st.session_state[_P + "esq"] = {
            "pptx": dados,
            "json": json.dumps(manifesto, ensure_ascii=False, indent=2).encode("utf-8"),
            "manifesto": manifesto,
            "base": f"Esqueleto_{cfg.local.strip().replace(' ', '_')}",
        }

    esq = st.session_state.get(_P + "esq")
    if not esq:
        return
    man = esq["manifesto"]
    pend = sum(1 for s in man["slides"] if s["status"] == "pendente")
    c1, c2, c3 = st.columns(3)
    c1.metric("Slides", len(man["slides"]))
    c2.metric("Com dados pendentes", pend)
    c3.metric("Campos a informar", len(man["campos_pendentes"]))
    if man["residuos_de_outros_estudos"]:
        st.error("Restos de outro estudo encontrados no esqueleto — veja o manifesto: "
                 + ", ".join(f"slide {r['posicao']}: {r['termo']}" for r in man["residuos_de_outros_estudos"][:8]))
    if man["campos_pendentes"]:
        st.warning("Campos ainda não informados (ficam como «(pendente)»): " + "; ".join(man["campos_pendentes"]))
    for aviso in man["avisos"]:
        st.info(aviso)

    c1, c2 = st.columns(2)
    c1.download_button("Baixar esqueleto (PowerPoint)", data=esq["pptx"], file_name=esq["base"] + ".pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                       key=_P + "dl_pptx")
    c2.download_button("Baixar manifesto (JSON)", data=esq["json"], file_name=esq["base"] + "_manifesto.json",
                       mime="application/json", key=_P + "dl_json")
    st.caption("Entregue o .pptx e o .json à IA (skill relatorio-analitico-informa) junto com os arquivos de análise "
               "para preencher os dados. Slides com aviso vermelho dependem de arquivo; sem o arquivo, o aviso "
               "passa a «Arquivo de X não enviado – conteúdo pendente».")
    with st.expander("Ver plano de slides"):
        st.dataframe(pd.DataFrame([{"Nº": s["posicao"], "Módulo": s["modulo"], "Título": s["titulo"],
                                    "Status": s["status"], "Fonte esperada": s["fonte_esperada"] or ""}
                                   for s in man["slides"]]), use_container_width=True, hide_index=True)


# ----------------------------------------------------------------------------- etapa 3
_TIPOS_ENVIO = ["zip", "pptx", "xlsx", "xlsm", "xls"]


def _aba_preencher():
    st.write(
        "Envie o **zip do estudo** (ou os arquivos soltos). O app identifica cada arquivo pelo conteúdo, "
        "preenche o que reconhece e deixa o resto com o aviso **«Arquivo de X não enviado – conteúdo pendente»**. "
        "Hoje preenche: **Perfil** (gênero, idade, renda e escolaridade)."
    )
    esq_sessao = st.session_state.get(_P + "esq")
    enviados = st.file_uploader("Arquivos do estudo (zip, esqueleto .pptx, Perfil .xls/.xlsx)", type=_TIPOS_ENVIO,
                                accept_multiple_files=True, key=_P + "preencher_upload")
    try:
        achados = classificar([(f.name, f.getvalue()) for f in (enviados or [])])
    except Exception as exc:
        st.exception(exc)
        return
    por_tipo = achados["por_tipo"]

    esqueleto = por_tipo.get("esqueleto")
    if esq_sessao and (esqueleto is None or st.checkbox(
            f"Usar o esqueleto gerado na aba 2 ({esq_sessao['base']}.pptx) em vez do enviado", value=True,
            key=_P + "usar_esq_sessao")):
        esqueleto = (esq_sessao["base"] + ".pptx", esq_sessao["pptx"])

    linhas = [{"Arquivo": "Esqueleto da apresentação", "Situação": "✅ " + esqueleto[0] if esqueleto
               else "❌ faltou — gere na aba 2 ou envie o .pptx"}]
    for tipo, rotulo in ARQUIVOS_ESPERADOS:
        linhas.append({"Arquivo": rotulo, "Situação": "✅ " + por_tipo[tipo][0] if tipo in por_tipo
                       else "❌ não enviado — o slide fica como «conteúdo pendente»"})
    st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)
    for tipo, nomes in achados["duplicados"].items():
        st.warning(f"Há mais de um arquivo de {tipo}; usei o primeiro e ignorei: {', '.join(nomes)}.")
    if achados["ignorados"]:
        with st.expander(f"{len(achados['ignorados'])} arquivo(s) não usados nesta etapa"):
            st.write(", ".join(achados["ignorados"]))

    if esqueleto is None:
        return
    base_esperada = None
    if esq_sessao:
        base_esperada = (esq_sessao["manifesto"].get("configuracao") or {}).get("base")
    if st.button("Preencher apresentação", type="primary", key=_P + "preencher"):
        try:
            with st.status("Preenchendo...", expanded=False) as status:
                dados, conf = preencher(esqueleto[1], por_tipo, base_esperada=base_esperada)
                status.update(label="Apresentação preenchida", state="complete")
        except Exception as exc:
            st.exception(exc)
            return
        nome = esqueleto[0].rsplit(".", 1)[0].replace("Esqueleto_", "")
        st.session_state[_P + "preenchido"] = {
            "pptx": dados, "conf": conf, "xlsx": conf.para_xlsx({"local": nome}), "nome": nome,
        }

    res = st.session_state.get(_P + "preenchido")
    if not res:
        return
    conf = res["conf"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Valores preenchidos", len(conf.linhas))
    c2.metric("Pendências", len(conf.pendencias))
    c3.metric("Divergências", len(conf.divergencias))
    for item, a, b, obs in conf.divergencias:
        st.error(f"**{item}:** {a} × {b}. {obs}")
    for slide, item, motivo in conf.pendencias:
        st.warning(f"Slide {slide} · {item}: {motivo}")
    for aviso in conf.avisos:
        st.info(aviso)
    c1, c2 = st.columns(2)
    c1.download_button("Baixar apresentação preenchida (PowerPoint)", data=res["pptx"],
                       file_name=f"{res['nome']}_preenchido.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                       key=_P + "dl_preenchido")
    c2.download_button("Baixar conferência (Excel)", data=res["xlsx"], file_name=f"Conferencia_{res['nome']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_P + "dl_conf")
