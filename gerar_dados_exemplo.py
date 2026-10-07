"""
Cria dados_exemplo/clientes.xlsx com 25 clientes FICTÍCIOS de uma oficina.
As datas são calculadas a partir de hoje, para haver sempre algo a vencer.
Inclui de propósito algumas linhas com erros.

Uso:
    python gerar_dados_exemplo.py
"""

import random
import unicodedata
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

NOMES = ["Ana Costa", "Rui Matos", "Marta Pinho", "João Valente", "Inês Barros",
         "Pedro Tavares", "Sofia Neves", "Tiago Ramos", "Carla Fonseca", "Nuno Leite",
         "Helena Sá", "Bruno Faria", "Luísa Moura", "André Pires", "Diana Coelho",
         "Vasco Cunha", "Filipa Rocha", "Hugo Maia", "Joana Teles", "Paulo Brito"]
PASTA = Path("dados_exemplo")


def matricula_ficticia() -> str:
    letras = "ABCDEFGHJLMNOPRSTVXZ"
    l1 = "".join(random.choices(letras, k=2))
    l2 = "".join(random.choices(letras, k=2))
    return f"{l1}-{random.randint(10, 99)}-{l2}"


def gerar():
    random.seed(3)
    PASTA.mkdir(exist_ok=True)
    hoje = date.today()
    linhas = []
    for nome in NOMES:
        # tira acentos: "Inês Barros" → "ines.barros"
        sem_acentos = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode()
        email = sem_acentos.lower().replace(" ", ".")
        linhas.append({
            "nome": nome,
            "email": f"{email}@exemplo.pt",
            "matricula": matricula_ficticia(),
            "tipo": random.choice(["Revisão", "IPO"]),
            "data_vencimento": (hoje + timedelta(days=random.randint(-10, 90))).strftime("%d/%m/%Y"),
            "consentimento": random.choice(["Sim", "Sim", "Não"]),
        })

    # Linhas com erros de propósito
    linhas.append({"nome": "Rita Sousa", "email": "rita@exemplo", "matricula": "AB-12-CD",
                   "tipo": "IPO", "data_vencimento": (hoje + timedelta(days=5)).strftime("%d/%m/%Y"),
                   "consentimento": "Sim"})                                      # email inválido
    linhas.append({"nome": "Carlos Melo", "email": "carlos@exemplo.pt", "matricula": "1234",
                   "tipo": "Revisão", "data_vencimento": "10/10/2026", "consentimento": "Sim"})  # matrícula má
    linhas.append({"nome": "Vera Lopes", "email": "vera@exemplo.pt", "matricula": "ZZ-99-AA",
                   "tipo": "Lavagem", "data_vencimento": "10/10/2026", "consentimento": "Não"})  # tipo inválido
    linhas.append({"nome": "Rui Dias", "email": "rui.dias@exemplo.pt", "matricula": "AA-10-BB",
                   "tipo": "IPO", "data_vencimento": "31/02/2026", "consentimento": "Sim"})     # data impossível
    linhas.append({"nome": "Teste\nBcc: spam@exemplo.pt", "email": "x@exemplo.pt",
                   "matricula": "CC-20-DD", "tipo": "IPO",
                   "data_vencimento": (hoje + timedelta(days=3)).strftime("%d/%m/%Y"),
                   "consentimento": "Sim"})                                      # tentativa de injeção

    pd.DataFrame(linhas).to_excel(PASTA / "clientes.xlsx", index=False)
    print(f"Criado '{PASTA}/clientes.xlsx' com {len(linhas)} clientes fictícios")


if __name__ == "__main__":
    gerar()
