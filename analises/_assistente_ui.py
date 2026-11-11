"""
Peças visuais compartilhadas pelos assistentes em etapas (Relatório Automatizado, Total Automático):
stepper com as etapas, barra de progresso, tela de boas-vindas em cartões e a "pílula" do arquivo em andamento.

Só desenha HTML/CSS; quem controla o fluxo continua sendo o módulo de cada assistente. As cores vêm das
variáveis do tema escuro das ferramentas (`analises/_tema_ferramenta.py`), com valores de reserva.

As funções `html_*` são puras (devolvem texto) para poderem ser testadas sem Streamlit.
"""

from __future__ import annotations

import html
from typing import Iterable, Mapping, Sequence

import streamlit as st

CSS = """
<style>
.ta-steps{display:flex;flex-wrap:wrap;gap:6px;margin:2px 0 10px}
.ta-passo{display:flex;align-items:center;gap:8px;padding:6px 11px 6px 7px;border-radius:999px;border:1px solid var(--tf-linha,#273244);
  background:rgba(255,255,255,.03);color:var(--tf-fraco,#6f7f96);font-size:13px;font-weight:700;line-height:1;white-space:nowrap}
.ta-bola{flex:none;width:22px;height:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;
  border:1.5px solid var(--tf-linha2,#33415a);color:var(--tf-muted,#9aa8bd)}
.ta-passo.ta-feito .ta-bola{background:rgba(180,226,107,.16);border-color:rgba(180,226,107,.55);color:var(--tf-lima,#b4e26b)}
.ta-passo.ta-atual{color:#0e1520;background:var(--tf-lima,#b4e26b);border-color:var(--tf-lima,#b4e26b);box-shadow:0 6px 18px rgba(164,214,86,.25)}
.ta-passo.ta-atual .ta-bola{background:#0e1520;border-color:#0e1520;color:var(--tf-lima,#b4e26b)}
.stApp .ta-passo .ta-nome{color:var(--tf-fraco,#6f7f96)}
.stApp .ta-passo.ta-feito .ta-nome{color:var(--tf-muted,#9aa8bd)}
.stApp .ta-passo.ta-atual .ta-nome{color:#0e1520}
.stApp .ta-passo .ta-bola{color:var(--tf-muted,#9aa8bd)}
.stApp .ta-passo.ta-feito .ta-bola{color:var(--tf-lima,#b4e26b)}
.stApp .ta-passo.ta-atual .ta-bola{color:var(--tf-lima,#b4e26b)}
.ta-barra{height:4px;border-radius:4px;background:var(--tf-linha,#273244);overflow:hidden;margin-bottom:6px}
.ta-barra > span{display:block;height:100%;background:var(--tf-lima,#b4e26b);border-radius:4px;transition:width .3s}
.ta-etiqueta{font-size:12px;font-weight:800;letter-spacing:1.6px;color:var(--tf-lima,#b4e26b);margin:0 0 2px}
.ta-arquivo{display:inline-flex;align-items:center;gap:8px;max-width:100%;padding:8px 14px;border:1px solid var(--tf-linha,#273244);
  border-radius:12px;background:rgba(255,255,255,.03);font-size:14px;color:var(--tf-muted,#9aa8bd)}
.stApp .ta-arquivo b, .stApp .ta-arquivo{color:var(--tf-muted,#9aa8bd)}
.stApp .ta-arquivo b{color:#fff}
.ta-arquivo b{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ta-boasvindas{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:12px;margin:6px 0 16px}
.ta-item{display:flex;gap:12px;align-items:flex-start;padding:14px 16px;border:1px solid var(--tf-linha,#273244);border-radius:16px;
  background:linear-gradient(180deg,rgba(255,255,255,.035),rgba(255,255,255,.015))}
.stApp .ta-item .ta-num{color:#0e1520}
.ta-num{flex:none;width:30px;height:30px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-weight:800;
  color:#0e1520;background:var(--tf-lima,#b4e26b)}
.stApp .ta-item h4{margin:0 0 2px;font-size:15.5px;color:#fff;padding:0;border:0}
.stApp .ta-item p{margin:0;font-size:13.5px;color:var(--tf-muted,#9aa8bd);line-height:1.4}
.ta-fora{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:0 0 18px;color:var(--tf-muted,#9aa8bd);font-size:13.5px}
.ta-fora span.ta-tag{padding:4px 10px;border-radius:8px;border:1px dashed var(--tf-linha2,#33415a);color:var(--tf-fraco,#6f7f96)}
</style>
"""


def html_stepper(etapas: Sequence[tuple[str, str]], atual: str, nomes_curtos: Mapping[str, str]) -> str:
    """Stepper: uma pílula por etapa (feita ✓ / atual / pendente) + barra de progresso.
    `etapas` = [(chave, nome completo)]. Devolve "" se `atual` não é uma das etapas."""
    chaves = [c for c, _ in etapas]
    if atual not in chaves:
        return ""
    idx = chaves.index(atual)
    passos = []
    for i, chave in enumerate(chaves):
        estado = "ta-feito" if i < idx else ("ta-atual" if i == idx else "")
        bola = "✓" if i < idx else str(i + 1)
        passos.append(
            f'<div class="ta-passo {estado}"><span class="ta-bola">{bola}</span>'
            f'<span class="ta-nome">{html.escape(nomes_curtos.get(chave, chave))}</span></div>'
        )
    pct = round(100 * idx / max(len(chaves) - 1, 1))
    nome = dict(etapas)[atual]
    return (
        f'<div class="ta-etiqueta">ETAPA {idx + 1} DE {len(chaves)}</div>'
        f'<div class="ta-barra"><span style="width:{pct}%"></span></div>'
        f'<div class="ta-steps" title="{html.escape(nome)}">' + "".join(passos) + "</div>"
    )


def html_boas_vindas(
    etapas: Sequence[tuple[str, str]],
    descricoes: Mapping[str, str],
    fora_do_fluxo: Iterable[str] = (),
    ignorar: Iterable[str] = ("final",),
) -> str:
    """Cartões numerados com o que cada etapa faz + a linha «Fora deste fluxo»."""
    pular = set(ignorar)
    itens = []
    n = 0
    for chave, nome in etapas:
        if chave in pular:
            continue
        n += 1
        itens.append(
            f'<div class="ta-item"><span class="ta-num">{n}</span><div><h4>{html.escape(nome)}</h4>'
            f'<p>{html.escape(descricoes.get(chave, ""))}</p></div></div>'
        )
    corpo = '<div class="ta-boasvindas">' + "".join(itens) + "</div>"
    fora = "".join(f'<span class="ta-tag">{html.escape(x)}</span>' for x in fora_do_fluxo)
    if fora:
        corpo += f'<div class="ta-fora"><b>Fora deste fluxo:</b> {fora}</div>'
    return corpo


def html_cartoes(itens: Sequence[tuple[str, str, str]]) -> str:
    """Cartões informativos sem numeração de etapa: itens = [(marcador, título, texto)].
    O marcador é um símbolo curto (ex.: «≤», «÷», «🔒» ou «1»)."""
    corpo = "".join(
        f'<div class="ta-item"><span class="ta-num">{html.escape(m)}</span><div><h4>{html.escape(t)}</h4>'
        f'<p>{html.escape(d)}</p></div></div>'
        for m, t, d in itens
    )
    return f'<div class="ta-boasvindas">{corpo}</div>'


def mostrar_cartoes(itens: Sequence[tuple[str, str, str]], etiqueta: str = "COMO FUNCIONA") -> None:
    st.markdown(CSS + f'<div class="ta-etiqueta">{html.escape(etiqueta)}</div>' + html_cartoes(itens), unsafe_allow_html=True)


def html_arquivo(nome: str | None) -> str:
    return f'<div class="ta-arquivo">Arquivo em andamento <b>{html.escape(nome)}</b></div>' if nome else ""


def mostrar_stepper(etapas: Sequence[tuple[str, str]], atual: str, nomes_curtos: Mapping[str, str]) -> None:
    corpo = html_stepper(etapas, atual, nomes_curtos)
    if corpo:
        st.markdown(CSS + corpo, unsafe_allow_html=True)


def mostrar_boas_vindas(texto_html: str, etapas, descricoes, fora_do_fluxo=()) -> None:
    st.markdown(
        CSS + '<div class="ta-etiqueta">COMO FUNCIONA</div>' + texto_html + html_boas_vindas(etapas, descricoes, fora_do_fluxo),
        unsafe_allow_html=True,
    )


def mostrar_arquivo(nome: str | None) -> None:
    if nome:
        st.markdown(CSS + html_arquivo(nome), unsafe_allow_html=True)
