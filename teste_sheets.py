"""
Teste de conexão com Google Sheets via OAuth2 do usuário.
Na primeira execução abre o navegador para autorizar. O token
fica salvo em token.json para as próximas execuções.
"""

import os
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PASTA_ID = "1MsRODhlMhWxRkKPni5jAJlJnqi5TOLJr"

# Carrega token salvo ou abre o navegador para autorizar
creds = None
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("oauth_credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
    with open("token.json", "w") as f:
        f.write(creds.to_json())

gc = gspread.authorize(creds)
drive = build("drive", "v3", credentials=creds)

# Cria a planilha diretamente na pasta correta
arquivo = drive.files().create(
    body={
        "name": "TESTE_RCO",
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [PASTA_ID],
    },
    fields="id",
).execute()

sheet_id = arquivo["id"]

planilha = gc.open_by_key(sheet_id)
planilha.sheet1.update([["Conexão funcionando!"]], "A1")

print(f"Planilha criada: https://docs.google.com/spreadsheets/d/{sheet_id}")
