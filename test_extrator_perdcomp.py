import csv
import json
import unittest

import extrator_perdcomp as ex

NUMERO = "27184.28191.230615.1.3.04-0063"


class ClienteFalso:
    """Devolve uma resposta pronta, para testar como o programa trata a IA."""
    def __init__(self, resposta=None, erro=None):
        self.resposta, self.erro = resposta, erro

    def responder(self, texto):
        if self.erro:
            raise self.erro
        return self.resposta


class TestCnpj(unittest.TestCase):
    def test_cnpj_valido_e_invalido(self):
        base = "112223330001"
        valido = base + ex.dv_cnpj(base)
        self.assertTrue(ex.cnpj_valido(valido))
        errado = valido[:13] + str((int(valido[13]) + 1) % 10)
        self.assertFalse(ex.cnpj_valido(errado))

    def test_mascara_remove_o_cnpj(self):
        self.assertEqual(ex.mascarar_cnpj("CNPJ: 11.222.333/0001-81"), "CNPJ: [CNPJ]")


class TestRegra(unittest.TestCase):
    def test_numero_limpo(self):
        self.assertEqual(ex.achar_numeros_por_regra(f"PER/DCOMP: {NUMERO}\nData: 01/01/2026"), [NUMERO])

    def test_numero_quebrado_em_duas_linhas(self):
        quebrado = NUMERO.replace("28191.", "28191.\n")
        self.assertEqual(ex.achar_numeros_por_regra(f"PER/DCOMP: {quebrado}"), [NUMERO])

    def test_ignora_outros_numeros(self):
        texto = f"Protocolo: 1234567890123\nProcesso: 10880.123456/2024-71\nPER/DCOMP: {NUMERO}"
        self.assertEqual(ex.achar_numeros_por_regra(texto), [NUMERO])

    def test_leitura_ruidosa_nao_e_resolvida_pela_regra(self):
        ruidoso = "PER/DCOMP: 271 84.28191.23O615.1.3.O4-0063"
        self.assertEqual(ex.achar_numeros_por_regra(ruidoso), [])


class TestConferenciaDaIa(unittest.TestCase):
    ruidoso = "PER/DCOMP: 271 84.28191.23O615.1.3.O4-0063"

    def test_resposta_correta_e_aceita(self):
        cliente = ClienteFalso(json.dumps({"perdcomp": NUMERO}))
        self.assertEqual(ex.extrair_com_ia(self.ruidoso, cliente), (NUMERO, ""))

    def test_aceita_json_dentro_de_bloco_de_codigo(self):
        cliente = ClienteFalso('```json\n{"perdcomp": "%s"}\n```' % NUMERO)
        self.assertEqual(ex.extrair_com_ia(self.ruidoso, cliente)[0], NUMERO)

    def test_numero_inventado_e_rejeitado(self):
        inventado = "99999.99999.010125.1.1.01-0001"
        numero, motivo = ex.extrair_com_ia(self.ruidoso, ClienteFalso(json.dumps({"perdcomp": inventado})))
        self.assertIsNone(numero)
        self.assertIn("não consta no documento", motivo)

    def test_formato_errado_e_rejeitado(self):
        numero, motivo = ex.extrair_com_ia(self.ruidoso, ClienteFalso(json.dumps({"perdcomp": "12345"})))
        self.assertIsNone(numero)
        self.assertIn("formato", motivo)

    def test_resposta_nula_ou_invalida(self):
        self.assertIsNone(ex.extrair_com_ia(self.ruidoso, ClienteFalso('{"perdcomp": null}'))[0])
        self.assertIsNone(ex.extrair_com_ia(self.ruidoso, ClienteFalso("não sei"))[0])

    def test_falha_da_ia_nao_derruba_o_programa(self):
        numero, motivo = ex.extrair_com_ia(self.ruidoso, ClienteFalso(erro=TimeoutError()))
        self.assertIsNone(numero)
        self.assertIn("falha ao chamar a IA", motivo)

    def test_cnpj_nao_vai_para_a_ia(self):
        enviados = []

        class Espiao(ClienteFalso):
            def responder(self, texto):
                enviados.append(texto)
                return '{"perdcomp": null}'

        ex.extrair_com_ia("CNPJ: 11.222.333/0001-81 " + self.ruidoso, Espiao())
        self.assertNotIn("11.222.333/0001-81", enviados[0])


def comparar_com_gabarito(cliente):
    _, resultados = ex.processar_planilha(ex.PLANILHA_PADRAO, cliente)
    por_processo = {r["processo"]: r for r in resultados}
    iguais, total = 0, 0
    with open("gabarito.csv", encoding="utf-8") as f:
        for linha in csv.DictReader(f):
            total += 1
            r = por_processo[linha["processo"]]
            if r["status"] == linha["status_esperado"] and r["perdcomp"] == linha["perdcomp_esperado"]:
                iguais += 1
    return iguais, total


class TestGabarito(unittest.TestCase):
    def test_com_cliente_simulado_todos_batem(self):
        iguais, total = comparar_com_gabarito(ex.ClienteSimulado())
        self.assertEqual((iguais, total), (30, 30))

    def test_sem_ia_so_os_tres_casos_ruidosos_falham(self):
        iguais, total = comparar_com_gabarito(None)
        self.assertEqual((iguais, total), (27, 30))


if __name__ == "__main__":
    unittest.main()
