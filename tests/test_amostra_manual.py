import unittest

from core.amostra_alocacao import (
    AlocacaoManualErro,
    aplicar_alocacao_manual,
    preparar_tabela_alocacao_manual,
)
from core.amostra_camadas import (
    aplicar_alocacao_manual_camadas,
    distribuir_amostra_em_camadas,
)
from core.amostra_math import FAIXAS_ETARIAS, calcular_amostra
from tests.test_amostra_camadas import base_sintetica


class TestAmostraManual(unittest.TestCase):
    def test_aplica_valores_manuais_e_recalcula_cotas(self):
        proporcional = calcular_amostra(
            codigos_municipios=["001", "002", "003"],
            amostra_total=20,
            nivel_quebra="municipio",
            base_populacional="total",
            df_base=base_sintetica(),
        )
        manual = aplicar_alocacao_manual(
            proporcional,
            nivel="municipio",
            alocacoes_manuais={"001": 10, "002": 0, "003": 10},
            amostra_total=20,
        )

        por_codigo = manual.set_index("cod_municipio")
        self.assertEqual(por_codigo["amostra"].to_dict(), {"001": 10, "002": 0, "003": 10})
        self.assertEqual(int(manual["amostra"].sum()), 20)
        self.assertTrue((manual["metodo_distribuicao"] == "manual").all())
        self.assertEqual(
            por_codigo["amostra_proporcional"].astype(int).to_dict(),
            {"001": 10, "002": 5, "003": 5},
        )
        self.assertEqual(
            por_codigo["diferenca_vs_proporcional"].astype(int).to_dict(),
            {"001": 0, "002": -5, "003": 5},
        )

        self.assertTrue(
            ((manual["cota_masc"] + manual["cota_fem"]) == manual["amostra"]).all()
        )
        colunas_idade = [f"cota_{faixa}" for faixa in FAIXAS_ETARIAS]
        self.assertTrue((manual[colunas_idade].sum(axis=1) == manual["amostra"]).all())
        self.assertAlmostEqual(float(manual["percentual_amostra"].sum()), 100.0)

    def test_editor_ignora_alocacao_salva_de_outro_recorte(self):
        proporcional = calcular_amostra(
            codigos_municipios=["001", "002"],
            amostra_total=20,
            nivel_quebra="municipio",
            base_populacional="total",
            df_base=base_sintetica(),
        )
        tabela = preparar_tabela_alocacao_manual(
            proporcional,
            nivel="municipio",
            alocacoes_iniciais={"001": 20, "999": 0},
        )
        self.assertEqual(
            tabela.set_index("cod_municipio")["amostra_manual"].astype(int).to_dict(),
            proporcional.set_index("cod_municipio")["amostra_proporcional"]
            .astype(int)
            .to_dict(),
        )

    def test_rejeita_soma_manual_diferente_da_meta(self):
        proporcional = calcular_amostra(
            codigos_municipios=["001", "002"],
            amostra_total=20,
            nivel_quebra="municipio",
            base_populacional="total",
            df_base=base_sintetica(),
        )
        with self.assertRaisesRegex(AlocacaoManualErro, "soma 19"):
            aplicar_alocacao_manual(
                proporcional,
                nivel="municipio",
                alocacoes_manuais={"001": 10, "002": 9},
                amostra_total=20,
            )

    def test_rejeita_unidade_ausente_ou_codigo_fora_do_recorte(self):
        proporcional = calcular_amostra(
            codigos_municipios=["001", "002"],
            amostra_total=20,
            nivel_quebra="municipio",
            base_populacional="total",
            df_base=base_sintetica(),
        )
        with self.assertRaisesRegex(AlocacaoManualErro, "não corresponde ao recorte"):
            aplicar_alocacao_manual(
                proporcional,
                nivel="municipio",
                alocacoes_manuais={"001": 10, "999": 10},
                amostra_total=20,
            )

    def test_camadas_reconstroem_pais_a_partir_do_nivel_final(self):
        proporcional = distribuir_amostra_em_camadas(
            base_sintetica(),
            amostra_total=40,
            camadas=["uf", "regiao_intermediaria", "municipio"],
            base_populacional="total",
        )
        manual = aplicar_alocacao_manual_camadas(
            proporcional,
            alocacoes_manuais={"001": 10, "002": 10, "003": 10, "004": 10},
            amostra_total=40,
        )

        self.assertEqual(int(manual.resultado["amostra"].sum()), 40)
        self.assertTrue(all(item["ok"] for item in manual.validacoes))
        for resumo in manual.resumos.values():
            self.assertEqual(int(resumo["amostra"].sum()), 40)
            self.assertTrue((resumo["metodo_distribuicao"] == "manual").all())

        ufs = manual.resumos["uf"].set_index("uf")["amostra"].astype(int).to_dict()
        self.assertEqual(ufs, {"AA": 30, "BB": 10})

        por_ri = (
            manual.resumos["municipio"]
            .groupby(["uf", "cod_regiao_intermediaria"], dropna=False)["amostra"]
            .sum()
            .sort_index()
        )
        ris = (
            manual.resumos["regiao_intermediaria"]
            .set_index(["uf", "cod_regiao_intermediaria"])["amostra"]
            .sort_index()
        )
        self.assertTrue(por_ri.equals(ris))


if __name__ == "__main__":
    unittest.main()
