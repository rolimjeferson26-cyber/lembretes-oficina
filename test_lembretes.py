"""
Testes. Correr com:
    python -m unittest -v
Nenhum teste envia emails a sério: tudo em modo teste, numa pasta temporária.
"""

import os
import tempfile
import unittest
from datetime import date
from email import message_from_bytes, policy
from pathlib import Path

import pandas as pd

import lembretes

HOJE = date(2026, 10, 1)


def config_teste(ficheiro: Path, avisar=True) -> dict:
    return {"nome_oficina": "Oficina Teste", "email_oficina": "oficina@exemplo.pt",
            "ficheiro_clientes": str(ficheiro), "dias_aviso_oficina": 30,
            "dias_aviso_cliente": 15, "avisar_clientes": avisar, "modo_teste": True,
            "smtp_host": "", "smtp_porta": 587, "smtp_utilizador": "", "smtp_senha": ""}


class TestLembretes(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.antes = os.getcwd()
        os.chdir(self.tmp.name)  # registo e emails_teste ficam na pasta temporária

    def tearDown(self):
        os.chdir(self.antes)
        self.tmp.cleanup()

    def criar_excel(self, linhas: list[dict]) -> Path:
        caminho = Path("clientes.xlsx")
        pd.DataFrame(linhas).to_excel(caminho, index=False)
        return caminho

    def cliente(self, **extra) -> dict:
        base = {"nome": "Ana Costa", "email": "ana@exemplo.pt", "matricula": "AB-12-CD",
                "tipo": "IPO", "data_vencimento": "10/10/2026", "consentimento": "Sim"}
        return {**base, **extra}

    def emails(self) -> list:
        return [message_from_bytes(p.read_bytes(), policy=policy.default) for p in sorted(Path("emails_teste").glob("*.eml"))]

    def test_resumo_e_aviso_com_consentimento(self):
        f = self.criar_excel([self.cliente()])
        r = lembretes.executar(config_teste(f), HOJE)
        self.assertEqual((r["a_vencer"], r["avisos_clientes"]), (1, 1))
        msgs = self.emails()
        self.assertEqual(msgs[0]["To"], "oficina@exemplo.pt")
        self.assertEqual(msgs[1]["To"], "ana@exemplo.pt")

    def test_sem_consentimento_nao_avisa_cliente(self):
        f = self.criar_excel([self.cliente(consentimento="Não")])
        r = lembretes.executar(config_teste(f), HOJE)
        self.assertEqual(r["avisos_clientes"], 0)
        self.assertIn("ligar ao cliente", self.emails()[0].get_content())

    def test_fora_do_prazo_nao_entra(self):
        f = self.criar_excel([self.cliente(data_vencimento="31/12/2026"),   # longe demais
                              self.cliente(data_vencimento="01/06/2026")])  # atrasado há meses
        r = lembretes.executar(config_teste(f), HOJE)
        self.assertEqual((r["a_vencer"], r["avisos_clientes"]), (0, 0))

    def test_nao_repete_aviso_ja_enviado(self):
        f = self.criar_excel([self.cliente()])
        c = lembretes.ler_clientes(f)[0][0]
        lembretes.guardar_no_registo([lembretes.chave(c, "cliente")])
        r = lembretes.executar(config_teste(f), HOJE)
        self.assertEqual(r["avisos_clientes"], 0)

    def test_modo_teste_nao_grava_registo(self):
        f = self.criar_excel([self.cliente()])
        lembretes.executar(config_teste(f), HOJE)
        self.assertFalse(lembretes.REGISTO.exists())

    def test_linhas_com_erro_vao_para_o_resumo(self):
        f = self.criar_excel([self.cliente(email="mau@exemplo"),
                              self.cliente(matricula="1234"),
                              self.cliente(tipo="Lavagem"),
                              self.cliente(data_vencimento="31/02/2026")])
        r = lembretes.executar(config_teste(f), HOJE)
        self.assertEqual(r["erros"], 4)
        self.assertIn("ATENÇÃO", self.emails()[0].get_content())

    def test_injecao_de_cabecalho_bloqueada(self):
        f = self.criar_excel([self.cliente(nome="Ana\nBcc: spam@exemplo.pt")])
        lembretes.executar(config_teste(f), HOJE)
        for msg in self.emails():
            self.assertIsNone(msg["Bcc"])

    def test_config_real_sem_senha_da_erro(self):
        os.environ["MODO_TESTE"] = "false"
        os.environ["SMTP_SENHA"] = ""
        try:
            with self.assertRaises(ValueError):
                lembretes.carregar_config()
        finally:
            del os.environ["MODO_TESTE"]
            del os.environ["SMTP_SENHA"]

    def test_so_aceita_xlsx(self):
        Path("c.xlsm").write_bytes(b"x")
        with self.assertRaises(ValueError):
            lembretes.ler_clientes(Path("c.xlsm"))


if __name__ == "__main__":
    unittest.main()
