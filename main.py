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
import fitz  # PyMuPDF
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
ABA_PARCELA = "PARCELAS"

PALAVRAS_IMPOSTO = ["darf", "imposto", "das", "irpj", "iss", "icms", "inss"]

# regex simples pra achar valor (R$ 1.234,56) e data (dd/mm/aaaa)
REGEX_VALOR = re.compile(r"R\$\s?([\d\.]+,\d{2})")
REGEX_DATA = re.compile(r"(\d{2}/\d{2}/\d{4})")
# aceita "NF 123", "NF-123", "NF - 123", "NF nº123", "NFe: 123", "NFSE 123", "NFS-e 123" etc.
# usa [ \t] (não \s) pra nunca atravessar quebra de linha e evitar casar com números de outras seções
_GAP = r"[ \t]*-?[ \t]*"
REGEX_NF = re.compile(rf"NF{_GAP}S?{_GAP}E?{_GAP}n?[ºo°]?\.?{_GAP}:?{_GAP}(\d{{1,9}})", re.IGNORECASE)
# fallback pra quando o documento escreve "Número da Nota Fiscal" por extenso, com o número
# na linha seguinte (comum em NFS-e de prefeitura) — aqui SIM permite pular 1 quebra de linha,
# mas só até achar o primeiro número isolado logo em seguida
REGEX_NUMERO_NF_EXTENSO = re.compile(r"n[uú]mero\s+da\s+nota\s+fiscal[ \t]*\n?[ \t]*(\d{1,9})", re.IGNORECASE)

# cabeçalho típico de aviso de pagamento (ex: TK Elevator) com tabela NF + Valor
REGEX_CABECALHO_TABELA = re.compile(r"n[ºo°]?\s*\.?\s*da\s*nf", re.IGNORECASE)
REGEX_LINHA_TABELA = re.compile(r"^\s*(\d{1,9})\s+([\d\.]+,\d{2})\s*$")


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


def ocr_pdf(dados):
    """OCR página a página, usado quando o PDF não tem texto extraível
    (ex: documento escaneado / foto salva como PDF)."""
    texto = ""
    doc = fitz.open(stream=dados, filetype="pdf")
    for pagina in doc:
        pix = pagina.get_pixmap(dpi=200)
        imagem = Image.open(io.BytesIO(pix.tobytes("png")))
        texto += pytesseract.image_to_string(imagem, lang="por")
    return texto


def extrair_texto(filename, dados):
    if filename.lower().endswith(".pdf"):
        texto = ""
        with pdfplumber.open(io.BytesIO(dados)) as pdf:
            for pagina in pdf.pages:
                texto += pagina.extract_text() or ""
        if texto.strip():
            return texto
        print(f"  {filename}: sem texto extraível, tentando OCR...")
        return ocr_pdf(dados)
    elif filename.lower().endswith((".png", ".jpg", ".jpeg")):
        imagem = Image.open(io.BytesIO(dados))
        return pytesseract.image_to_string(imagem, lang="por")
    return ""


def parsear_aviso_pagamento(texto):
    """
    Trata o formato de "aviso de pagamento" que lista várias NFs numa tabela,
    ex:
        N° da NF          Valor
        63                4.975,00
    Retorna uma lista de dicts (uma entrada por NF encontrada na tabela) ou
    lista vazia se o texto não tiver esse formato.
    """
    data_match = REGEX_DATA.search(texto)
    data = data_match.group(1) if data_match else datetime.today().strftime("%d/%m/%Y")

    linhas = texto.splitlines()
    resultados = []
    capturando = False

    for linha in linhas:
        if REGEX_CABECALHO_TABELA.search(linha):
            capturando = True
            continue
        if not capturando:
            continue

        m = REGEX_LINHA_TABELA.match(linha)
        if m:
            resultados.append({
                "numero": m.group(1),
                "valor": m.group(2),
                "data": data,
                "tipo": "PARCELA_NF",
            })
        elif resultados and linha.strip() and not linha.strip().startswith("-"):
            # já capturou pelo menos uma linha da tabela e veio algo que não
            # é mais tabela nem separador -> a tabela acabou
            break

    return resultados


def parsear_dados(texto, assunto):
    valores = REGEX_VALOR.findall(texto)
    # em tabelas com várias colunas (descontos, retenções, total), o valor final
    # listado é, na prática, quase sempre o total/líquido da nota
    valor = valores[-1] if valores else "NÃO IDENTIFICADO"

    data_match = REGEX_DATA.search(texto)
    nf_match = REGEX_NF.search(texto) or REGEX_NUMERO_NF_EXTENSO.search(texto)

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
        (ABA_PARCELA, ["N° da NF", "Valor da Parcela", "Data Pagamento", "Assunto do Email", "Data Processamento"]),
    ]:
        if aba not in wb.sheetnames:
            ws = wb.create_sheet(aba)
            ws.append(cabecalho)

    return wb


def adicionar_linha(wb, dados, assunto):
    if dados["tipo"] == "PARCELA_NF":
        aba = ABA_PARCELA
    elif dados["tipo"] == "IMPOSTO":
        aba = ABA_IMPOSTO
    else:
        aba = ABA_NF

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

            parcelas = parsear_aviso_pagamento(texto)
            if parcelas:
                for dados in parcelas:
                    adicionar_linha(wb, dados, assunto)
                    novos += 1
                    print(f"Processado (parcela): {filename} -> {dados}")
            else:
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
