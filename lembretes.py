"""
lembretes.py
------------
Lembretes automáticos de revisão e inspeção (IPO) para uma oficina.

Lê um Excel de clientes e:
  FASE 1: envia à OFICINA um resumo com tudo o que vence nos próximos N dias.
  FASE 2: envia a cada CLIENTE (só se deu consentimento) um aviso X dias antes.

Uso:
    python lembretes.py                    # usa as definições do ficheiro .env
    python lembretes.py --hoje 2026-10-12  # simula outro dia (útil para testar)

Segurança:
  - MODO_TESTE=true (o padrão): NÃO envia nada, grava os emails como ficheiros
    na pasta emails_teste/ para veres exatamente o que iria sair.
  - A senha do email fica no ficheiro .env, que nunca vai para o GitHub.
  - Ligação ao servidor de email sempre cifrada (TLS) e com certificado verificado.
  - Emails só em texto simples (sem HTML), o que evita conteúdo malicioso.
  - Registo de envios: o mesmo aviso nunca é enviado duas vezes.
  - Só avisa clientes com consentimento = Sim (RGPD).
"""

import argparse
import csv
import os
import re
import smtplib
import ssl
import sys
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

COLUNAS = ["nome", "email", "matricula", "tipo", "data_vencimento", "consentimento"]
TIPOS_VALIDOS = {"revisão": "Revisão", "revisao": "Revisão", "ipo": "Inspeção (IPO)",
                 "inspeção": "Inspeção (IPO)", "inspecao": "Inspeção (IPO)"}
PADRAO_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$")
# Matrículas portuguesas: AA-00-AA, 00-AA-00, 00-00-AA, AA-00-00
PADRAO_MATRICULA = re.compile(r"^([A-Z]{2}-\d{2}-[A-Z]{2}|\d{2}-[A-Z]{2}-\d{2}|\d{2}-\d{2}-[A-Z]{2}|[A-Z]{2}-\d{2}-\d{2})$")
REGISTO = Path("registo_envios.csv")
PASTA_TESTE = Path("emails_teste")


# ----------------------------------------------------------------------
# 1. Configuração (vem do .env, nunca do código)
# ----------------------------------------------------------------------
def carregar_config() -> dict:
    load_dotenv()  # lê o ficheiro .env da pasta atual, se existir
    config = {
        "nome_oficina": os.getenv("NOME_OFICINA", "Oficina"),
        "email_oficina": os.getenv("EMAIL_OFICINA", ""),
        "ficheiro_clientes": os.getenv("FICHEIRO_CLIENTES", "dados_exemplo/clientes.xlsx"),
        "dias_aviso_oficina": int(os.getenv("DIAS_AVISO_OFICINA", "30")),
        "dias_aviso_cliente": int(os.getenv("DIAS_AVISO_CLIENTE", "15")),
        "avisar_clientes": os.getenv("AVISAR_CLIENTES", "false").lower() == "true",
        # Qualquer valor diferente de "false" conta como modo teste: mais seguro
        "modo_teste": os.getenv("MODO_TESTE", "true").lower() != "false",
        "smtp_host": os.getenv("SMTP_HOST", ""),
        "smtp_porta": int(os.getenv("SMTP_PORTA", "587")),
        "smtp_utilizador": os.getenv("SMTP_UTILIZADOR", ""),
        "smtp_senha": os.getenv("SMTP_SENHA", ""),
    }
    if not config["modo_teste"]:
        em_falta = [k for k in ("email_oficina", "smtp_host", "smtp_utilizador", "smtp_senha")
                    if not config[k]]
        if em_falta:
            raise ValueError(f"Faltam definições no .env: {', '.join(em_falta)}")
    return config


# ----------------------------------------------------------------------
# 2. Ler e validar clientes
# ----------------------------------------------------------------------
def sem_quebras(texto: str) -> str:
    """Tira quebras de linha (evita que alguém meta cabeçalhos falsos no email)."""
    return " ".join(str(texto).split())


def ler_clientes(caminho: Path) -> tuple[list[dict], list[str]]:
    """Devolve (clientes_validos, erros). Erros não param o programa."""
    if caminho.suffix.lower() != ".xlsx":
        raise ValueError("O ficheiro de clientes tem de ser .xlsx (sem macros).")
    if not caminho.is_file():
        raise FileNotFoundError(f"Não encontrei '{caminho}'.")

    df = pd.read_excel(caminho, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip().lower() for c in df.columns]
    em_falta = [c for c in COLUNAS if c not in df.columns]
    if em_falta:
        raise ValueError(f"Faltam colunas no Excel: {', '.join(em_falta)}")

    validos, erros = [], []
    for i, linha in df.iterrows():
        n = i + 2  # número da linha no Excel
        nome = sem_quebras(linha["nome"])
        email = linha["email"].strip().lower()
        matricula = linha["matricula"].strip().upper().replace(" ", "-")
        tipo = TIPOS_VALIDOS.get(linha["tipo"].strip().lower())
        consentiu = linha["consentimento"].strip().lower() in ("sim", "s", "yes")

        try:
            # aceita "2026-10-20", "2026-10-20 00:00:00" (Excel) ou "20/10/2026"
            texto_data = linha["data_vencimento"].strip()[:10]
            if "/" in texto_data:
                vencimento = datetime.strptime(texto_data, "%d/%m/%Y").date()
            else:
                vencimento = datetime.strptime(texto_data, "%Y-%m-%d").date()
        except ValueError:
            erros.append(f"Linha {n}: data inválida")
            continue

        if not nome:
            erros.append(f"Linha {n}: nome em branco")
        elif not PADRAO_MATRICULA.match(matricula):
            erros.append(f"Linha {n}: matrícula com formato estranho")
        elif tipo is None:
            erros.append(f"Linha {n}: tipo deve ser Revisão ou IPO")
        elif consentiu and not PADRAO_EMAIL.match(email):
            erros.append(f"Linha {n}: email inválido")
        else:
            validos.append({"nome": nome, "email": email, "matricula": matricula,
                            "tipo": tipo, "vencimento": vencimento, "consentiu": consentiu})
    return validos, erros


# ----------------------------------------------------------------------
# 3. Registo de envios (para nunca repetir)
# ----------------------------------------------------------------------
def chave(c: dict, destino: str) -> str:
    return f"{destino}|{c['matricula']}|{c['tipo']}|{c['vencimento'].isoformat()}"


def ler_registo() -> set[str]:
    if not REGISTO.exists():
        return set()
    with open(REGISTO, encoding="utf-8") as f:
        return {linha["chave"] for linha in csv.DictReader(f)}


def guardar_no_registo(chaves: list[str]) -> None:
    novo = not REGISTO.exists()
    with open(REGISTO, "a", newline="", encoding="utf-8") as f:
        escritor = csv.writer(f)
        if novo:
            escritor.writerow(["chave", "enviado_em"])
        agora = datetime.now().isoformat(timespec="seconds")
        for k in chaves:
            escritor.writerow([k, agora])


# ----------------------------------------------------------------------
# 4. Escrever os emails (texto simples)
# ----------------------------------------------------------------------
def email_resumo_oficina(lista: list[dict], hoje: date, dias: int, nome_oficina: str) -> tuple[str, str]:
    assunto = f"[{nome_oficina}] {len(lista)} revisões/inspeções nos próximos {dias} dias"
    if not lista:
        return assunto, f"Nada a vencer nos próximos {dias} dias. Boa semana!"
    linhas = [f"Resumo de {hoje.strftime('%d/%m/%Y')} — por ordem de data:\n"]
    for c in sorted(lista, key=lambda c: c["vencimento"]):
        faltam = (c["vencimento"] - hoje).days
        quando = "JÁ PASSOU" if faltam < 0 else ("HOJE" if faltam == 0 else f"daqui a {faltam} dias")
        aviso = "aviso por email: sim" if c["consentiu"] else "ligar ao cliente"
        linhas.append(f"- {c['vencimento'].strftime('%d/%m/%Y')} ({quando}) | {c['tipo']} | "
                      f"{c['matricula']} | {c['nome']} | {aviso}")
    return assunto, "\n".join(linhas)


def email_aviso_cliente(c: dict, nome_oficina: str) -> tuple[str, str]:
    primeiro_nome = c["nome"].split()[0]
    assunto = f"{nome_oficina}: {c['tipo']} do {c['matricula']} a {c['vencimento'].strftime('%d/%m/%Y')}"
    corpo = (
        f"Olá {primeiro_nome},\n\n"
        f"Lembramos que a {c['tipo']} da viatura {c['matricula']} "
        f"está prevista para {c['vencimento'].strftime('%d/%m/%Y')}.\n\n"
        f"Se quiser marcar, basta responder a este email ou ligar-nos.\n\n"
        f"Cumprimentos,\n{nome_oficina}\n\n"
        f"---\nRecebe este aviso porque autorizou o contacto. "
        f"Para deixar de receber, responda com a palavra PARAR."
    )
    return assunto, corpo


# ----------------------------------------------------------------------
# 5. Enviar (ou simular)
# ----------------------------------------------------------------------
def criar_mensagem(de: str, para: str, assunto: str, corpo: str) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = de
    msg["To"] = para
    msg["Subject"] = sem_quebras(assunto)
    msg.set_content(corpo)  # texto simples
    return msg


def enviar(mensagens: list[EmailMessage], config: dict) -> None:
    if config["modo_teste"]:
        PASTA_TESTE.mkdir(exist_ok=True)
        for i, msg in enumerate(mensagens, start=1):
            (PASTA_TESTE / f"{i:03d}.eml").write_bytes(bytes(msg))
        return

    contexto = ssl.create_default_context()  # verifica o certificado do servidor
    with smtplib.SMTP(config["smtp_host"], config["smtp_porta"], timeout=30) as servidor:
        servidor.starttls(context=contexto)  # cifra a ligação ANTES de enviar a senha
        servidor.login(config["smtp_utilizador"], config["smtp_senha"])
        for msg in mensagens:
            servidor.send_message(msg)


# ----------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------
def executar(config: dict, hoje: date) -> dict:
    clientes, erros = ler_clientes(Path(config["ficheiro_clientes"]))
    ja_enviados = ler_registo()
    remetente = config["smtp_utilizador"] or "teste@exemplo.pt"
    destino_oficina = config["email_oficina"] or "oficina@exemplo.pt"
    mensagens, novas_chaves = [], []

    # FASE 1: resumo para a oficina. Inclui os atrasados até 30 dias,
    # para não esquecer ninguém (mais antigos do que isso já não interessam).
    limite = hoje + timedelta(days=config["dias_aviso_oficina"])
    inicio = hoje - timedelta(days=30)
    a_vencer = [c for c in clientes if inicio <= c["vencimento"] <= limite]
    assunto, corpo = email_resumo_oficina(a_vencer, hoje, config["dias_aviso_oficina"], config["nome_oficina"])
    if erros:
        corpo += "\n\nATENÇÃO — linhas do Excel com problemas (corrigir):\n" + "\n".join(f"- {e}" for e in erros)
    mensagens.append(criar_mensagem(remetente, destino_oficina, assunto, corpo))

    # FASE 2: aviso a cada cliente (opcional, só com consentimento, uma vez só)
    avisos = 0
    if config["avisar_clientes"]:
        limite_cliente = hoje + timedelta(days=config["dias_aviso_cliente"])
        for c in clientes:
            k = chave(c, "cliente")
            if c["consentiu"] and hoje <= c["vencimento"] <= limite_cliente and k not in ja_enviados:
                assunto, corpo = email_aviso_cliente(c, config["nome_oficina"])
                mensagens.append(criar_mensagem(remetente, c["email"], assunto, corpo))
                novas_chaves.append(k)
                avisos += 1

    enviar(mensagens, config)
    if not config["modo_teste"]:
        guardar_no_registo(novas_chaves)  # em modo teste não marca como enviado
    return {"clientes": len(clientes), "a_vencer": len(a_vencer),
            "avisos_clientes": avisos, "erros": len(erros), "teste": config["modo_teste"]}


def main() -> int:
    parser = argparse.ArgumentParser(description="Lembretes de revisão/IPO.")
    parser.add_argument("--hoje", help="Simular outra data (AAAA-MM-DD)")
    args = parser.parse_args()

    try:
        hoje = date.fromisoformat(args.hoje) if args.hoje else date.today()
        config = carregar_config()
        r = executar(config, hoje)
    except (FileNotFoundError, ValueError) as erro:
        print(f"ERRO: {erro}")
        return 1
    except smtplib.SMTPAuthenticationError:
        print("ERRO: o servidor recusou o login. Confirma a senha de aplicação no .env.")
        return 1
    except (smtplib.SMTPException, OSError) as erro:
        print(f"ERRO ao enviar email: {type(erro).__name__}. Verifica a internet e o .env.")
        return 1

    modo = "MODO TESTE (nada enviado, ver pasta emails_teste/)" if r["teste"] else "ENVIADO"
    print(f"{modo}: {r['clientes']} clientes lidos, {r['a_vencer']} a vencer, "
          f"{r['avisos_clientes']} avisos a clientes, {r['erros']} linhas com erro.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
