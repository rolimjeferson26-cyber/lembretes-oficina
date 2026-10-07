# Lembretes de Revisão e IPO para Oficinas

Uma vez por semana, o script lê o Excel de clientes da oficina e:

1. **Envia à oficina** um resumo com tudo o que vence nos próximos 30 dias, mais o que ficou atrasado. Para cada cliente diz se recebe aviso por email ou se é preciso ligar.
2. **Envia a cada cliente que autorizou** um lembrete 15 dias antes. Esta parte é opcional.

O script nunca envia o mesmo aviso duas vezes. Se houver linhas com erros no Excel, aparecem no resumo da oficina para serem corrigidas.

> Todos os dados de exemplo são **fictícios** (`@exemplo.pt`).

## Exemplo do resumo que a oficina recebe

```
Resumo de 07/10/2026 — por ordem de data:

- 14/10/2026 (daqui a 7 dias) | Inspeção (IPO) | XH-83-SJ | Inês Barros | aviso por email: sim
- 29/10/2026 (daqui a 22 dias) | Inspeção (IPO) | FL-12-NZ | Hugo Maia | ligar ao cliente

ATENÇÃO — linhas do Excel com problemas (corrigir):
- Linha 22: email inválido
```

## O Excel de clientes

| nome | email | matricula | tipo | data_vencimento | consentimento |
|---|---|---|---|---|---|
| Ana Costa | ana.costa@exemplo.pt | AB-12-CD | IPO | 20/10/2026 | Sim |

- `tipo`: `Revisão` ou `IPO`.
- `consentimento`: `Sim` só se o cliente autorizou ser contactado por email (RGPD).

## Instalação

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate   |   Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
python gerar_dados_exemplo.py
```

Depois copia `.env.example` para `.env` e preenche-o.

## Experimentar em segurança (modo teste)

Com `MODO_TESTE=true`, que é o padrão, **nada é enviado**. Os emails ficam gravados em `emails_teste/` e podes abri-los com duplo clique.

```bash
python lembretes.py
python lembretes.py --hoje 2026-11-01   # simula outro dia
```

## Pôr a funcionar a sério

1. Cria uma **senha de aplicação** no email da oficina. No Gmail: Conta Google → Segurança → Verificação em 2 passos (tem de estar ligada) → Palavras-passe de aplicações. **Nunca** uses a senha normal.
2. No `.env`, preenche `SMTP_UTILIZADOR` e `SMTP_SENHA` e muda para `MODO_TESTE=false`.
3. Faz um teste com o teu próprio email em `EMAIL_OFICINA`.
4. Agenda a execução semanal no **Agendador de Tarefas do Windows**:
   - Ação: Iniciar um programa
   - Programa: `C:\caminho\lembretes-oficina\.venv\Scripts\python.exe`
   - Argumentos: `lembretes.py`
   - Iniciar em: `C:\caminho\lembretes-oficina`
   - Disparador: semanal, segunda-feira às 9:00

## Segurança e privacidade

- **Senha:** só existe no `.env`, que nunca vai para o GitHub (está no `.gitignore`).
- **Ligação:** é sempre cifrada (STARTTLS) e o certificado do servidor é verificado.
- **Emails:** só em texto simples, sem HTML, links ou anexos.
- **Injeção de cabeçalhos:** as quebras de linha nos nomes são removidas. Testado com `"Ana\nBcc: spam@..."`.
- **Consentimento:** só recebem aviso os clientes com `Sim`. Cada email explica como deixar de receber.
- **Sem repetições:** o `registo_envios.csv` garante que o mesmo aviso não é enviado duas vezes.
- **Excel:** só aceita `.xlsx` (sem macros).
- **Mensagens de erro:** nunca mostram a senha nem dados de clientes.
- **Dependências:** estão fixadas e foram auditadas com `pip-audit` (sem vulnerabilidades a 07/10/2026).

## Testes

```bash
python -m unittest -v
```

São 9 testes, todos em modo teste numa pasta temporária. Cobrem:
- o resumo e o consentimento;
- os prazos;
- as repetições;
- a injeção de cabeçalhos;
- a configuração sem senha;
- as extensões recusadas.

## O que NÃO faz

SMS, WhatsApp, ligação ao programa de faturação ou marcações. Isso fica para versões futuras, se o cliente quiser.
