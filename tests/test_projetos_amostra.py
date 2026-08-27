import json
import unittest

from core.projetos_amostra import (
    ProjetoAmostraErro,
    criar_configuracao_projeto,
    desserializar_configuracao_projeto,
    serializar_configuracao_projeto,
)


class ProjetosAmostraTests(unittest.TestCase):
    def test_roundtrip_json_preserva_parametros(self):
        projeto = criar_configuracao_projeto(
            nome="Nacional 2026",
            fonte_dados="banco_sqlite",
            modo="camadas",
            amostra_total=2000,
            base_populacional="16_mais",
            faixa_flex_pct=20,
            uf=None,
            regioes_intermediarias=[],
            regioes_imediatas=[{"codigo": "3501", "nome": "São Paulo"}],
            municipios=[{"codigo": "3550308", "nome": "São Paulo"}],
            distritos=[{"codigo": "355030805", "nome": "São Paulo"}],
            incluir_distritos=True,
            camadas=["uf", "regiao_intermediaria", "municipio"],
        )
        restaurado = desserializar_configuracao_projeto(
            serializar_configuracao_projeto(projeto)
        )
        self.assertEqual(restaurado["nome"], "Nacional 2026")
        self.assertEqual(restaurado["amostra_total"], 2000)
        self.assertEqual(
            restaurado["camadas"],
            ["uf", "regiao_intermediaria", "municipio"],
        )
        self.assertEqual(restaurado["fonte_dados"], "banco_sqlite")
        self.assertEqual(
            restaurado["territorio"]["municipios"],
            [{"codigo": "3550308", "nome": "São Paulo"}],
        )
        self.assertEqual(
            restaurado["territorio"]["distritos"],
            [{"codigo": "355030805", "nome": "São Paulo"}],
        )

    def test_json_nao_contem_base_demografica(self):
        projeto = criar_configuracao_projeto(
            nome="SP",
            fonte_dados="banco_sqlite",
            modo="classico",
            amostra_total=500,
            base_populacional="total",
            faixa_flex_pct=10,
            uf="SP",
            regioes_intermediarias=[{"codigo": "3501", "nome": "São Paulo"}],
            nivel_quebra="municipio",
        )
        dados = json.loads(serializar_configuracao_projeto(projeto).decode("utf-8"))
        self.assertNotIn("populacao", dados)
        self.assertNotIn("territorios", dados)
        self.assertEqual(dados["territorio"]["uf"], "SP")

    def test_migra_projeto_1_0_sem_recortes_novos(self):
        antigo = {
            "schema_version": "1.0",
            "tipo": "projeto_amostra",
            "nome": "Projeto antigo",
            "fonte_dados": "banco_sqlite",
            "modo": "classico",
            "amostra_total": 100,
            "base_populacional": "16_mais",
            "faixa_flex_pct": 20,
            "territorio": {
                "uf": "SP",
                "regioes_intermediarias": [],
            },
            "classico": {
                "nivel_quebra": "municipio",
                "incluir_distritos": False,
            },
            "camadas": [],
        }
        restaurado = desserializar_configuracao_projeto(json.dumps(antigo))
        self.assertEqual(restaurado["schema_version"], "1.2")
        self.assertEqual(restaurado["territorio"]["municipios"], [])
        self.assertEqual(restaurado["territorio"]["distritos"], [])
        self.assertEqual(restaurado["distribuicao"]["metodo"], "proporcional")
        self.assertEqual(restaurado["distribuicao"]["alocacoes_manuais"], [])
        self.assertEqual(restaurado["exportacao"]["exibicao_cotas"], "ambos")

    def test_roundtrip_preserva_distribuicao_manual_e_preferencia_excel(self):
        projeto = criar_configuracao_projeto(
            nome="Manual municipal",
            fonte_dados="banco_sqlite",
            modo="classico",
            amostra_total=20,
            base_populacional="total",
            faixa_flex_pct=20,
            uf="AA",
            municipios=[
                {"codigo": "001", "nome": "Mun A"},
                {"codigo": "002", "nome": "Mun B"},
            ],
            nivel_quebra="municipio",
            metodo_distribuicao="manual",
            alocacoes_manuais=[
                {"codigo": "001", "amostra": 12},
                {"codigo": "002", "amostra": 8},
            ],
            exibicao_cotas="percentuais",
        )
        restaurado = desserializar_configuracao_projeto(
            serializar_configuracao_projeto(projeto)
        )
        self.assertEqual(restaurado["schema_version"], "1.2")
        self.assertEqual(restaurado["distribuicao"]["metodo"], "manual")
        self.assertEqual(
            restaurado["distribuicao"]["alocacoes_manuais"],
            [
                {"codigo": "001", "amostra": 12},
                {"codigo": "002", "amostra": 8},
            ],
        )
        self.assertEqual(restaurado["exportacao"]["exibicao_cotas"], "percentuais")

    def test_rejeita_alocacao_manual_que_nao_fecha_total(self):
        with self.assertRaisesRegex(ProjetoAmostraErro, "somar exatamente"):
            criar_configuracao_projeto(
                nome="Manual inválido",
                fonte_dados="banco_sqlite",
                modo="classico",
                amostra_total=20,
                base_populacional="total",
                faixa_flex_pct=20,
                uf="AA",
                nivel_quebra="municipio",
                metodo_distribuicao="manual",
                alocacoes_manuais=[
                    {"codigo": "001", "amostra": 12},
                    {"codigo": "002", "amostra": 7},
                ],
            )

    def test_rejeita_versao_incompativel(self):
        with self.assertRaises(ProjetoAmostraErro):
            desserializar_configuracao_projeto(
                b'{"schema_version":"9.9","tipo":"projeto_amostra"}'
            )

    def test_modo_camadas_exige_uma_camada(self):
        with self.assertRaises(ProjetoAmostraErro):
            criar_configuracao_projeto(
                nome="Incompleto",
                fonte_dados="banco_sqlite",
                modo="camadas",
                amostra_total=100,
                base_populacional="16_mais",
                faixa_flex_pct=20,
                uf=None,
                camadas=[],
            )


if __name__ == "__main__":
    unittest.main()
