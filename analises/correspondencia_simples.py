"""
Análise de Correspondência — versão original (um único cruzamento por vez).

Esta função é preservada EXATAMENTE como estava, sem nenhuma alteração de
lógica, apenas movida de app.py para este módulo.
"""
import plotly.graph_objects as go
import pandas as pd
import streamlit as st
from adjustText import adjust_text
from matplotlib import pyplot as plt

from core.ca_math import _ca_manual


def analise_correspondencia(dados):
    st.header("Análise de Correspondência")

    colunas_categoricas = st.multiselect(
        "Selecione as colunas categóricas para a análise:",
        dados.columns
    )

    if len(colunas_categoricas) < 2:
        st.warning("Selecione pelo menos duas colunas categóricas para realizar a análise.")
        return

    tabela_contingencia = pd.crosstab(
        dados[colunas_categoricas[0]],
        dados[colunas_categoricas[1]],
        normalize=False
    )

    try:
        coordenadas_linhas, coordenadas_colunas, eigenvalues = _ca_manual(
            tabela_contingencia,
            n_components=2
        )
    except Exception as e:
        st.error(f"Não foi possível calcular a Análise de Correspondência: {e}")
        return

    coordenadas_linhas *= -1
    coordenadas_colunas *= -1

    if coordenadas_linhas.shape[1] < 2 or coordenadas_colunas.shape[1] < 2:
        st.error(
            "A análise de correspondência não conseguiu gerar duas dimensões. "
            "Verifique os dados selecionados."
        )
        return

    st.subheader("Inércia explicada (variância)")
    explained_inertia = eigenvalues / eigenvalues.sum()
    for i, valor in enumerate(explained_inertia):
        st.write(f"Dim {i+1}: {valor * 100:.2f}%")

    if "rotulos_linhas" not in st.session_state:
        st.session_state["rotulos_linhas"] = {
            i: str(i) for i in coordenadas_linhas.index
        }
    else:
        for i in coordenadas_linhas.index:
            st.session_state["rotulos_linhas"].setdefault(i, str(i))

    if "rotulos_colunas" not in st.session_state:
        st.session_state["rotulos_colunas"] = {
            i: str(i) for i in coordenadas_colunas.index
        }
    else:
        for i in coordenadas_colunas.index:
            st.session_state["rotulos_colunas"].setdefault(i, str(i))

    if "deslocamentos_linhas" not in st.session_state:
        st.session_state["deslocamentos_linhas"] = {
            i: (0.0, 0.0) for i in coordenadas_linhas.index
        }
    else:
        for i in coordenadas_linhas.index:
            st.session_state["deslocamentos_linhas"].setdefault(i, (0.0, 0.0))

    if "deslocamentos_colunas" not in st.session_state:
        st.session_state["deslocamentos_colunas"] = {
            i: (0.0, 0.0) for i in coordenadas_colunas.index
        }
    else:
        for i in coordenadas_colunas.index:
            st.session_state["deslocamentos_colunas"].setdefault(i, (0.0, 0.0))

    legenda_linhas = st.text_input("Legenda para Linhas:", "Linhas")
    legenda_colunas = st.text_input("Legenda para Colunas:", "Colunas")
    mostrar_legenda = st.checkbox("Mostrar legenda (Linhas/Colunas)", value=False)
    editar_rotulos = st.checkbox("Editar rótulos e deslocamentos")

    if editar_rotulos:
        st.subheader("Editar Rótulos e Posições")

        st.markdown("**Linhas**")
        for i in coordenadas_linhas.index:
            st.session_state["rotulos_linhas"][i] = st.text_input(
                f"Novo rótulo para linha '{i}':",
                value=st.session_state["rotulos_linhas"][i],
                key=f"rotulo_linha_{i}"
            )

            dx, dy = st.session_state["deslocamentos_linhas"][i]
            dx = st.number_input(
                f"Deslocamento X para linha '{i}':",
                min_value=-1.0,
                max_value=1.0,
                value=float(dx),
                step=0.01,
                format="%.2f",
                key=f"num_desloc_x_linha_{i}"
            )
            dy = st.number_input(
                f"Deslocamento Y para linha '{i}':",
                min_value=-1.0,
                max_value=1.0,
                value=float(dy),
                step=0.01,
                format="%.2f",
                key=f"num_desloc_y_linha_{i}"
            )
            st.session_state["deslocamentos_linhas"][i] = (dx, dy)

            if st.checkbox(f"Remover ponto e rótulo da linha '{i}'", key=f"remover_linha_{i}"):
                st.session_state["rotulos_linhas"][i] = None

        st.markdown("**Colunas**")
        for i in coordenadas_colunas.index:
            st.session_state["rotulos_colunas"][i] = st.text_input(
                f"Novo rótulo para coluna '{i}':",
                value=st.session_state["rotulos_colunas"][i],
                key=f"rotulo_coluna_{i}"
            )

            dx, dy = st.session_state["deslocamentos_colunas"][i]
            dx = st.number_input(
                f"Deslocamento X para coluna '{i}':",
                min_value=-1.0,
                max_value=1.0,
                value=float(dx),
                step=0.01,
                format="%.2f",
                key=f"num_desloc_x_coluna_{i}"
            )
            dy = st.number_input(
                f"Deslocamento Y para coluna '{i}':",
                min_value=-1.0,
                max_value=1.0,
                value=float(dy),
                step=0.01,
                format="%.2f",
                key=f"num_desloc_y_coluna_{i}"
            )
            st.session_state["deslocamentos_colunas"][i] = (dx, dy)

            if st.checkbox(f"Remover ponto e rótulo da coluna '{i}'", key=f"remover_coluna_{i}"):
                st.session_state["rotulos_colunas"][i] = None

    max_deslocamento = 0.05
    pontos_linhas = []
    pontos_colunas = []

    for i, row in coordenadas_linhas.iterrows():
        if st.session_state["rotulos_linhas"][i] is None:
            continue
        x = row.iloc[0]
        y = row.iloc[1]
        dx, dy = st.session_state["deslocamentos_linhas"][i]
        dx = max(-max_deslocamento, min(max_deslocamento, dx * (x / abs(x) if x != 0 else 1)))
        dy = max(-max_deslocamento, min(max_deslocamento, dy * (y / abs(y) if y != 0 else 1)))
        pontos_linhas.append({
            "x": x,
            "y": y,
            "dx": dx,
            "dy": dy,
            "label": st.session_state["rotulos_linhas"][i]
        })

    for i, row in coordenadas_colunas.iterrows():
        if st.session_state["rotulos_colunas"][i] is None:
            continue
        x = row.iloc[0]
        y = row.iloc[1]
        dx, dy = st.session_state["deslocamentos_colunas"][i]
        dx = max(-max_deslocamento, min(max_deslocamento, dx * (x / abs(x) if x != 0 else 1)))
        dy = max(-max_deslocamento, min(max_deslocamento, dy * (y / abs(y) if y != 0 else 1)))
        pontos_colunas.append({
            "x": x,
            "y": y,
            "dx": dx,
            "dy": dy,
            "label": st.session_state["rotulos_colunas"][i]
        })

    st.subheader("Mapa estático")

    fig, ax = plt.subplots(figsize=(8, 6))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    texts = []
    linha_legenda_plotada = False
    coluna_legenda_plotada = False

    for p in pontos_linhas:
        label_legenda = legenda_linhas if (mostrar_legenda and not linha_legenda_plotada) else ""
        ax.scatter(p["x"], p["y"], color="blue", marker="o", s=20, label=label_legenda)
        if mostrar_legenda and not linha_legenda_plotada:
            linha_legenda_plotada = True

        txt = ax.text(
            p["x"] + p["dx"],
            p["y"] + p["dy"],
            p["label"],
            color="blue",
            fontsize=8,
            ha="center",
            va="bottom",
            fontweight="bold",
        )
        texts.append(txt)

    for p in pontos_colunas:
        label_legenda = legenda_colunas if (mostrar_legenda and not coluna_legenda_plotada) else ""
        ax.scatter(p["x"], p["y"], color="red", marker="^", s=30, label=label_legenda)
        if mostrar_legenda and not coluna_legenda_plotada:
            coluna_legenda_plotada = True

        txt = ax.text(
            p["x"] + p["dx"],
            p["y"] + p["dy"],
            p["label"],
            color="red",
            fontsize=8,
            ha="center",
            va="bottom",
            fontweight="bold",
        )
        texts.append(txt)

    adjust_text(
        texts,
        arrowprops=None,
        force_text=(0.5, 1),
        force_points=(0.5, 1),
        expand_text=(1.2, 1.5),
        expand_points=(1.2, 1.5),
        lim=100,
    )

    ax.axhline(0, color="black", linestyle="--", linewidth=0.8)
    ax.axvline(0, color="black", linestyle="--", linewidth=0.8)
    ax.grid(color="lightgray", linestyle="--", linewidth=0.5, alpha=0.7)
    ax.set_aspect("auto")

    if mostrar_legenda and (linha_legenda_plotada or coluna_legenda_plotada):
        ax.legend(loc="upper right", frameon=False, fontsize=12)

    ax.set_title("Análise de Correspondência", fontsize=14)
    st.pyplot(fig)

    st.subheader("Mapa interativo")

    fig_int = go.Figure()

    fig_int.add_trace(go.Scatter(
        x=[p["x"] for p in pontos_linhas],
        y=[p["y"] for p in pontos_linhas],
        mode="markers",
        name=legenda_linhas,
        marker=dict(color="blue", symbol="circle", size=8),
        showlegend=mostrar_legenda and len(pontos_linhas) > 0
    ))

    fig_int.add_trace(go.Scatter(
        x=[p["x"] for p in pontos_colunas],
        y=[p["y"] for p in pontos_colunas],
        mode="markers",
        name=legenda_colunas,
        marker=dict(color="red", symbol="triangle-up", size=10),
        showlegend=mostrar_legenda and len(pontos_colunas) > 0
    ))

    annotations = []
    for p in pontos_linhas:
        annotations.append(dict(
            x=p["x"] + p["dx"],
            y=p["y"] + p["dy"],
            xref="x",
            yref="y",
            text=p["label"],
            showarrow=False,
            font=dict(color="blue", size=10),
            xanchor="center",
            yanchor="bottom",
        ))
    for p in pontos_colunas:
        annotations.append(dict(
            x=p["x"] + p["dx"],
            y=p["y"] + p["dy"],
            xref="x",
            yref="y",
            text=p["label"],
            showarrow=False,
            font=dict(color="red", size=10),
            xanchor="center",
            yanchor="bottom",
        ))

    xs = [p["x"] for p in pontos_linhas] + [p["x"] for p in pontos_colunas]
    ys = [p["y"] for p in pontos_linhas] + [p["y"] for p in pontos_colunas]

    if xs and ys:
        x_min, x_max = min(xs), max(xs)
        y_min, y_max = min(ys), max(ys)
        dx = (x_max - x_min) * 0.1 or 0.1
        dy = (y_max - y_min) * 0.1 or 0.1
        x_range = [x_min - dx, x_max + dx]
        y_range = [y_min - dy, y_max + dy]
    else:
        x_range = [-1, 1]
        y_range = [-1, 1]

    fig_int.update_layout(
        template=None,
        title=dict(
            text="Análise de Correspondência",
            font=dict(color="black", size=14),
            x=0.5
        ),
        font=dict(color="black"),
        paper_bgcolor="white",
        plot_bgcolor="white",
        annotations=annotations,
        showlegend=mostrar_legenda,
        legend=dict(
            x=0.99, y=0.99,
            xanchor="right", yanchor="top",
            bgcolor="rgba(0,0,0,0)",
            font=dict(color="black", size=12),
        ),
        margin=dict(l=40, r=20, t=60, b=40),
        xaxis=dict(range=x_range),
        yaxis=dict(range=y_range),
    )

    fig_int.update_xaxes(
        title=dict(text="", font=dict(color="black")),
        showgrid=True,
        gridcolor="lightgray",
        gridwidth=0.5,
        griddash="dash",
        zeroline=False,
        tickfont=dict(color="black"),
        linecolor="black",
        mirror=True,
    )
    fig_int.update_yaxes(
        title=dict(text="", font=dict(color="black")),
        showgrid=True,
        gridcolor="lightgray",
        gridwidth=0.5,
        griddash="dash",
        zeroline=False,
        tickfont=dict(color="black"),
        linecolor="black",
        mirror=True,
    )

    fig_int.add_shape(
        type="line",
        x0=x_range[0], x1=x_range[1],
        y0=0, y1=0,
        line=dict(color="black", width=1, dash="dash"),
        xref="x", yref="y",
        layer="above"
    )
    fig_int.add_shape(
        type="line",
        x0=0, x1=0,
        y0=y_range[0], y1=y_range[1],
        line=dict(color="black", width=1, dash="dash"),
        xref="x", yref="y",
        layer="above"
    )

    config = {
        "editable": True,
        "edits": {"annotationPosition": True},
    }

    st.plotly_chart(fig_int, use_container_width=True, config=config)

