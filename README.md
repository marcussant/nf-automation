# Automação de Controle de NFs, Recibos e Impostos

Lê e-mails do Gmail com anexos de NF/recibo/imposto, extrai valor, data e número
via OCR/regex, e organiza automaticamente em uma planilha Excel (`controle.xlsx`),
tudo rodando de graça no GitHub Actions (sem servidor).

---

## Passo 1 — Criar o repositório

```bash
mkdir nf-automation && cd nf-automation
git init
# copie main.py, requirements.txt, .gitignore e a pasta .github/ pra dentro
git add .
git commit -m "estrutura inicial"
```

Cria o repo no GitHub (pode ser público, é portfólio) e sobe:

```bash
git remote add origin https://github.com/marcussant/nf-automation.git
git branch -M main
git push -u origin main
```

---

## Passo 2 — Ativar a Gmail API no Google Cloud

1. Vá em https://console.cloud.google.com/
2. Crie um projeto novo (ex: `nf-automation`)
3. No menu, vá em **APIs e Serviços > Biblioteca**
4. Procure por **Gmail API** e clique em **Ativar**

---

## Passo 3 — Criar credenciais OAuth

1. Ainda em APIs e Serviços, vá em **Tela de consentimento OAuth**
   - Tipo: **Externo**
   - Preencha nome do app, e-mail de suporte, e-mail do desenvolvedor
   - Em "Escopos", não precisa adicionar nada agora
   - Em "Usuários de teste", adicione seu próprio e-mail Gmail (enquanto o app
     não é publicado, só esses e-mails conseguem autorizar)
2. Vá em **Credenciais > Criar Credenciais > ID do cliente OAuth**
   - Tipo de aplicativo: **App para computador**
   - Dê um nome, clique em criar
3. Baixe o JSON gerado e salve como `credentials.json` na raiz do projeto
   (esse arquivo **não** vai pro git — já está no `.gitignore`)

---

## Passo 4 — Gerar o token de autorização (só uma vez, localmente)

Crie um script auxiliar `gerar_token.py` (não faz parte do projeto final,
é só pra rodar uma vez na sua máquina):

```python
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0)

with open("token.json", "w") as f:
    f.write(creds.to_json())

print("token.json gerado com sucesso!")
```

Rode:

```bash
pip install google-auth-oauthlib
python gerar_token.py
```

Isso abre o navegador, você loga com o Gmail que vai ser monitorado e autoriza.
Vai gerar um `token.json` na pasta.

---

## Passo 5 — Colocar o token como Secret no GitHub

O `token.json` **não sobe pro repositório**. Ele vira uma variável secreta:

```bash
base64 -i token.json | pbcopy   # Mac
# ou no Linux:
base64 -w0 token.json
```

Copie o resultado (uma string longa em base64). No GitHub:

1. Vá no repositório > **Settings > Secrets and variables > Actions**
2. **New repository secret**
3. Nome: `GMAIL_TOKEN_JSON`
4. Valor: cole a string em base64
5. Salvar

---

## Passo 6 — Testar

No GitHub, vá em **Actions > Automação NF/Recibos/Impostos > Run workflow**
(botão manual, `workflow_dispatch`). Acompanhe o log — se tudo estiver certo,
ele vai buscar e-mails recentes, processar os anexos e commitar o
`controle.xlsx` atualizado direto no repositório.

Depois disso, ele roda sozinho todo dia às 08h (horário de Brasília), sem
você precisar fazer nada.

---

## Observações importantes pro portfólio

- **Token expira?** O `refresh_token` do Google normalmente não expira
  enquanto o app estiver em modo de teste E você usar ele pelo menos 1x a
  cada 6 meses. Se expirar, é só rodar o Passo 4 de novo e atualizar o Secret.
- **Regex de extração:** hoje é simples (valor `R$ x,xx`, data `dd/mm/aaaa`,
  número de NF). Vai errar em formatos diferentes — é esperado, dá pra
  melhorar depois. Vale citar isso no README do GitHub como "próximos passos"
  (mostra maturidade técnica pra quem for avaliar).
- **Evolução natural pro currículo de DevOps:** depois que isso estiver
  rodando, você pode:
  - Dockerizar o `main.py` (Dockerfile simples)
  - Trocar o commit no git por escrita direto no Google Sheets (API)
  - Adicionar testes automatizados do parser de regex (`pytest`) rodando
    como um segundo job no mesmo workflow — isso já é "CI" de verdade
