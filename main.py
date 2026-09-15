"""
Automação de controle de Notas Fiscais / Recibos / Impostos via Gmail.

Fluxo:
1. Autentica na Gmail API (token OAuth já gerado previamente).
2. Busca e-mails novos com anexos (PDF ou imagem) que batam com uma query.
3. Baixa os anexos.
4. Extrai texto (PDF direto, imagem via OCR).
5. Usa regex para achar valor, data e número da NF.
6. Classifica em "NF/Recibo" ou "Imposto" (por palavras-chave no assunto/anexo).
7. Escreve/atualiza controle.xlsx (duas abas: NFS_RECIBOS e IMPOSTOS).
8. Marca o e-mail como "processado" (label) pra não reprocessar.
"""

import base64
import os
import re
import io
from datetime import datetime

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

import pdfplumber
from PIL import Image
import pytesseract

from openpyxl import Workbook, load_workbook

# ---------- CONFIG ----------
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
QUERY = 'has:attachment (nota fiscal OR NF OR recibo OR imposto OR DARF OR boleto) newer_than:2d'
LABEL_PROCESSADO = "NF-Processado"
PLANILHA = "controle.xlsx"
ABA_NF = "NFS_RECIBOS"
ABA_IMPOSTO = "IMPOSTOS"

PALAVRAS_IMPOSTO = ["darf", "imposto", "das", "irpj", "iss", "icms", "inss"]

# regex simples pra achar valor (R$ 1.234,56) e data (dd/mm/aaaa)
REGEX_VALOR = re.compile(r"R\$\s?([\d\.]+,\d{2})")
REGEX_DATA = re.compile(r"(\d{2}/\d{2}/\d{4})")
REGEX_NF = re.compile(r"(?:NF-?e?\s?n?[ºo°]?\s?)(\d{3,9})", re.IGNORECASE)


# ---------- AUTENTICAÇÃO ----------
def get_gmail_service():
    """
    Espera as variáveis de ambiente GMAIL_CREDENTIALS_JSON e GMAIL_TOKEN_JSON
    (conteúdo dos arquivos credentials.json e token.json, em base64).
    No GitHub Actions isso vem dos Secrets do repositório.
    """
    creds_b64 = os.environ["GMAIL_TOKEN_JSON"]
    creds_json = base64.b64decode(creds_b64).decode("utf-8")

    with open("token.json", "w") as f:
        f.write(creds_json)

    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    return build("gmail", "v1", credentials=creds)


# ---------- BUSCA DE E-MAILS ----------
def buscar_emails(service):
    resultados = service.users().messages().list(userId="me", q=QUERY).execute()
    return resultados.get("messages", [])


def ja_processado(service, msg_id, label_id):
    msg = service.users().messages().get(userId="me", id=msg_id, format="minimal").execute()
    return label_id in msg.get("labelIds", [])


def get_or_create_label(service, nome):
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    for l in labels:
        if l["name"] == nome:
            return l["id"]
    novo = service.users().labels().create(
        userId="me", body={"name": nome, "labelListVisibility": "labelShow", "messageListVisibility": "show"}
    ).execute()
    return novo["id"]


def marcar_processado(service, msg_id, label_id):
    service.users().messages().modify(
        userId="me", id=msg_id, body={"addLabelIds": [label_id]}
    ).execute()


# ---------- EXTRAÇÃO DE ANEXOS ----------
def baixar_anexos(service, msg_id):
    msg = service.users().messages().get(userId="me", id=msg_id).execute()
    partes = msg.get("payload", {}).get("parts", []) or []
    anexos = []

    for parte in partes:
        filename = parte.get("filename")
        body = parte.get("body", {})
        if filename and body.get("attachmentId"):
            att = service.users().messages().attachments().get(
                userId="me", messageId=msg_id, id=body["attachmentId"]
            ).execute()
            dados = base64.urlsafe_b64decode(att["data"])
            anexos.append((filename, dados))

    assunto = next((h["value"] for h in msg["payload"]["headers"] if h["name"] == "Subject"), "")
    return assunto, anexos


def extrair_texto(filename, dados):
    if filename.lower().endswith(".pdf"):
        texto = ""
        with pdfplumber.open(io.BytesIO(dados)) as pdf:
            for pagina in pdf.pages:
                texto += pagina.extract_text() or ""
        return texto
    elif filename.lower().endswith((".png", ".jpg", ".jpeg")):
        imagem = Image.open(io.BytesIO(dados))
        return pytesseract.image_to_string(imagem, lang="por")
    return ""


def parsear_dados(texto, assunto):
    valor_match = REGEX_VALOR.search(texto)
    data_match = REGEX_DATA.search(texto)
    nf_match = REGEX_NF.search(texto)

    valor = valor_match.group(1) if valor_match else "NÃO IDENTIFICADO"
    data = data_match.group(1) if data_match else datetime.today().strftime("%d/%m/%Y")
    numero_nf = nf_match.group(1) if nf_match else "NÃO IDENTIFICADO"

    texto_lower = (texto + assunto).lower()
    tipo = "IMPOSTO" if any(p in texto_lower for p in PALAVRAS_IMPOSTO) else "NF_RECIBO"

    return {"numero": numero_nf, "valor": valor, "data": data, "tipo": tipo}


# ---------- PLANILHA ----------
def abrir_planilha():
    if os.path.exists(PLANILHA):
        wb = load_workbook(PLANILHA)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    for aba, cabecalho in [
        (ABA_NF, ["NF/Recibo", "Valor", "Data", "Assunto do Email", "Data Processamento"]),
        (ABA_IMPOSTO, ["Referência", "Valor", "Data", "Assunto do Email", "Data Processamento"]),
    ]:
        if aba not in wb.sheetnames:
            ws = wb.create_sheet(aba)
            ws.append(cabecalho)

    return wb


def adicionar_linha(wb, dados, assunto):
    aba = ABA_IMPOSTO if dados["tipo"] == "IMPOSTO" else ABA_NF
    ws = wb[aba]
    ws.append([
        dados["numero"],
        dados["valor"],
        dados["data"],
        assunto,
        datetime.today().strftime("%d/%m/%Y %H:%M"),
    ])


# ---------- MAIN ----------
def main():
    service = get_gmail_service()
    label_id = get_or_create_label(service, LABEL_PROCESSADO)
    mensagens = buscar_emails(service)

    if not mensagens:
        print("Nenhum e-mail novo encontrado.")
        return

    wb = abrir_planilha()
    novos = 0

    for m in mensagens:
        msg_id = m["id"]
        if ja_processado(service, msg_id, label_id):
            continue

        assunto, anexos = baixar_anexos(service, msg_id)

        for filename, dados_bin in anexos:
            texto = extrair_texto(filename, dados_bin)
            if not texto:
                continue
            dados = parsear_dados(texto, assunto)
            adicionar_linha(wb, dados, assunto)
            novos += 1
            print(f"Processado: {filename} -> {dados}")

        marcar_processado(service, msg_id, label_id)

    if novos:
        wb.save(PLANILHA)
        print(f"{novos} registro(s) adicionados em {PLANILHA}")
    else:
        print("Nenhum anexo novo processado.")


if __name__ == "__main__":
    main()
