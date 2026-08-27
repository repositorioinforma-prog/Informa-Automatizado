import sys
import types
import unittest

# As funções de decisão territorial são puras; não é necessário instalar
# Streamlit para testá-las.
sys.modules.setdefault("streamlit", types.ModuleType("streamlit"))

from analises.gerador_amostra import (  # noqa: E402
    _determinar_nivel_alocacao_manual,
    _garantir_nivel_manual_nas_camadas,
)
from core.amostra_alocacao import aplicar_alocacao_manual  # noqa: E402
from core.amostra_math import FAIXAS_ETARIAS, calcular_amostra  # noqa: E402
from tests.test_amostra_camadas import base_sintetica  # noqa: E402


class TestMetasManuaisV5(unittest.TestCase):
    def test_municipios_especificos_nao_sao_agregados_na_regiao(self):
        codigos = ["001", "002", "003"]
        nivel = _determinar_nivel_alocacao_manual(
            "regiao_intermediaria",
            codigos_municipios=codigos,
        )
        self.assertEqual(nivel, "municipio")

        sugestao = calcular_amostra(
            codigos_municipios=codigos,
            amostra_total=50,
            nivel_quebra=nivel,
            base_populacional="total",
            df_base=base_sintetica(),
        )
        self.assertEqual(len(sugestao), 3)
        self.assertEqual(set(sugestao["cod_municipio"]), set(codigos))

    def test_meta_bruta_recalcula_segmentos_dentro_de_cada_municipio(self):
        sugestao = calcular_amostra(
            codigos_municipios=["001", "002", "003"],
            amostra_total=50,
            nivel_quebra="municipio",
            base_populacional="total",
            df_base=base_sintetica(),
        )
        manual = aplicar_alocacao_manual(
            sugestao,
            nivel="municipio",
            alocacoes_manuais={"001": 20, "002": 10, "003": 20},
            amostra_total=50,
        )

        self.assertEqual(int(manual["amostra"].sum()), 50)
        self.assertTrue(
            ((manual["cota_masc"] + manual["cota_fem"]) == manual["amostra"]).all()
        )
        colunas_idade = [f"cota_{faixa}" for faixa in FAIXAS_ETARIAS]
        self.assertTrue((manual[colunas_idade].sum(axis=1) == manual["amostra"]).all())

    def test_modo_camadas_adiciona_municipio_como_nivel_final_manual(self):
        camadas = _garantir_nivel_manual_nas_camadas(
            ["uf", "regiao_intermediaria"],
            "municipio",
        )
        self.assertEqual(camadas, ["uf", "regiao_intermediaria", "municipio"])


if __name__ == "__main__":
    unittest.main()
