import sys
import types
import unittest

import pandas as pd

# O ambiente de testes não precisa instalar Streamlit para validar as funções
# puras de preparação das opções territoriais.
sys.modules.setdefault("streamlit", types.ModuleType("streamlit"))

from analises.gerador_amostra import (  # noqa: E402
    _determinar_nivel_alocacao_manual,
    _filtrar_dataframe,
    _garantir_nivel_manual_nas_camadas,
    _montar_opcoes_territoriais,
)


class RecorteTerritorialUiTests(unittest.TestCase):
    def setUp(self):
        self.base = pd.DataFrame(
            [
                {
                    "uf": "AA",
                    "cod_regiao_intermediaria": "A1",
                    "regiao_intermediaria": "Região A",
                    "cod_regiao_imediata": "A11",
                    "regiao_imediata": "Imediata A",
                    "cod_municipio": "001",
                    "municipio": "Município Um",
                    "cod_distrito": "00101",
                    "distrito": "Distrito Um",
                },
                {
                    "uf": "AA",
                    "cod_regiao_intermediaria": "A1",
                    "regiao_intermediaria": "Região A",
                    "cod_regiao_imediata": "A11",
                    "regiao_imediata": "Imediata A",
                    "cod_municipio": "002",
                    "municipio": "Município Dois",
                    "cod_distrito": "00201",
                    "distrito": "Distrito Dois",
                },
            ]
        )

    def test_opcoes_de_regiao_nao_criam_colunas_duplicadas(self):
        opcoes, rotulos = _montar_opcoes_territoriais(
            self.base, "regiao_intermediaria"
        )
        self.assertEqual(opcoes, ["A1"])
        self.assertIn("Região A", rotulos["A1"])

    def test_municipio_pode_ser_filtrado_sem_regiao_pai(self):
        recorte = _filtrar_dataframe(
            self.base,
            codigos_municipios=["002"],
        )
        self.assertEqual(recorte["cod_municipio"].tolist(), ["002"])

    def test_distrito_pode_ser_filtrado_diretamente(self):
        recorte = _filtrar_dataframe(
            self.base,
            codigos_distritos=["00101"],
        )
        self.assertEqual(recorte["cod_distrito"].tolist(), ["00101"])

    def test_municipios_especificos_forcam_editor_por_municipio(self):
        nivel = _determinar_nivel_alocacao_manual(
            "regiao_intermediaria",
            codigos_municipios=["001", "002"],
        )
        self.assertEqual(nivel, "municipio")

    def test_distritos_especificos_forcam_editor_por_distrito(self):
        nivel = _determinar_nivel_alocacao_manual(
            "municipio",
            codigos_municipios=["001"],
            codigos_distritos=["00101"],
        )
        self.assertEqual(nivel, "distrito")

    def test_camadas_recebem_nivel_manual_mais_detalhado(self):
        camadas = _garantir_nivel_manual_nas_camadas(
            ["uf", "regiao_intermediaria"],
            "municipio",
        )
        self.assertEqual(camadas, ["uf", "regiao_intermediaria", "municipio"])


if __name__ == "__main__":
    unittest.main()
