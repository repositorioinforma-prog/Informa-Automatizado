"""
Acesso ao pyreadstat com plano B para leitura de .sav.

Problema real: no Windows do Instituto, uma política de Controle de
Aplicativo (Smart App Control / WDAC) bloqueia a DLL `_readstat_parser`
do pyreadstat dentro da .venv ("DLL load failed ... Uma política de
Controle de Aplicativo bloqueou este arquivo"). Reinstalar não resolve.

Comportamento:
- O pyreadstat só é importado quando alguém realmente lê/grava um .sav
  (assim o app sempre abre).
- Se ele carregar (ex.: Streamlit Cloud, Linux), é usado normalmente.
- Se NÃO carregar, `read_sav` passa a usar o leitor em Python puro
  (core/sav_puro.py), que devolve DataFrame + metadados no mesmo formato.
- Gravação (`write_sav`, usada na exportação da Ponderação) não tem plano
  B: nesse caso a ação mostra a mensagem abaixo.

Uso:  from core.pyreadstat_seguro import pyreadstat
      df, meta = pyreadstat.read_sav(caminho)
"""

MENSAGEM_BLOQUEIO = (
    "Não foi possível carregar o pyreadstat neste computador (provavelmente a "
    "política de Controle de Aplicativo do Windows bloqueou a DLL dentro da "
    ".venv). A LEITURA de .sav funciona mesmo assim (leitor em Python puro), "
    "mas a GRAVAÇÃO de .sav precisa do pyreadstat: peça à TI uma exceção para "
    "a pasta .venv do projeto ou use o app publicado no Streamlit Cloud."
)


class _PyreadstatSobDemanda:
    _modulo = None
    _falhou = None

    def _carregar(self):
        if _PyreadstatSobDemanda._modulo is None and _PyreadstatSobDemanda._falhou is None:
            try:
                import pyreadstat as _mod
                _PyreadstatSobDemanda._modulo = _mod
            except ImportError as e:  # inclui "DLL load failed"
                _PyreadstatSobDemanda._falhou = e
        return _PyreadstatSobDemanda._modulo

    @property
    def usando_leitor_puro(self) -> bool:
        return self._carregar() is None

    def __getattr__(self, nome):
        mod = self._carregar()
        if mod is not None:
            return getattr(mod, nome)
        if nome == "read_sav":
            from core.sav_puro import read_sav
            return read_sav
        raise ImportError(f"{MENSAGEM_BLOQUEIO}\n\nErro original: {_PyreadstatSobDemanda._falhou}")


pyreadstat = _PyreadstatSobDemanda()
