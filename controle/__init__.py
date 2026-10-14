"""
Controle Geral (antigo app "Controle-Automatico", integrado ao Informa).

Pacote com o MESMO código do app original (core/, geography/, importers/,
exporters/ e data/ com a Base Brasil). Só mudou:
- os imports absolutos ganharam o prefixo `controle.` (o Informa já tem um
  pacote `core/` próprio);
- importers/sav.py e exporters/sav.py usam core/pyreadstat_seguro.py (lê .sav
  mesmo com a DLL do pyreadstat bloqueada no Windows).
A tela fica em analises/controle.py.
"""
from pathlib import Path

PASTA_DADOS = Path(__file__).resolve().parent / "data"
