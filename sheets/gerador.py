"""
Gerador de diário de classe no Google Sheets.
"""

import os
from datetime import date, timedelta

import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PENALIDADES = [
    ("Fez a atividade",       0),
    ("Falta",               -100),
    ("Não fez a atividade",  -80),
    ("Não terminou",         -50),
    ("Entregou com atraso",  -30),
    ("Muita Conversa",       -30),
    ("Celular",              -60),
    ("Dormindo",             -80),
    ("Evasão(Gaseio)",      -200),
]

# Colunas fixas: A=Aluno, B=(vazio), C=Média1, D=Média2, E=Média3, F=Média, G=Nº
COL_FIXAS = 7
COL_AULAS_INICIO = COL_FIXAS + 1  # H = 8

MEDIAS_COLS = ["C", "D", "E"]   # uma por avaliação (máx 3)


# ---------------------------------------------------------------------------
# Autenticação OAuth2
# ---------------------------------------------------------------------------

def _get_creds():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "oauth_credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as f:
            f.write(creds.to_json())
    return creds


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_letter(n):
    """Converte número de coluna 1-based para letra (1→A, 27→AA)."""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _parsear_turma(nome_turma):
    """Extrai (serie, turno, letra) de '3ª Série - Noite - A'."""
    partes = [p.strip() for p in nome_turma.split(" - ")]
    if len(partes) >= 3:
        return partes[-3], partes[-2], partes[-1]
    if len(partes) == 2:
        return partes[0], "", partes[1]
    return nome_turma, "", ""


def _limpar(texto):
    return "".join(c for c in texto if c.isalnum() or c in ("_", "-")).strip()


def _gerar_nome(turma_data):
    """ALGATE_Noite_3Serie_A_FISICA"""
    escola = _limpar(turma_data["escola"].split()[0])
    serie, turno, letra = _parsear_turma(turma_data["turma"])
    serie_limpa = _limpar(serie.replace("ª", "").replace(" ", ""))
    disciplina = _limpar(turma_data["disciplina"].replace(" ", "_"))
    return f"{escola}_{turno}_{serie_limpa}_{letra}_{disciplina}"


def _gerar_datas(data_inicio, num_aulas, freq):
    if isinstance(data_inicio, str):
        data_inicio = date.fromisoformat(data_inicio)
    dias_aula = list(range(min(freq, 5)))
    datas, atual = [], data_inicio
    while len(datas) < num_aulas:
        if atual.weekday() in dias_aula:
            datas.append(atual)
        atual += timedelta(days=1)
    return datas


def _colunas_aulas(num_aulas, avaliacoes):
    """
    Retorna lista descrevendo cada coluna de aula após as colunas fixas:
      [{"tipo": "nota"|"ocorrencias"|"atividade", "aula": int|None, "avaliacao": int}]
    Cada avaliação tem N pares Nota+Ocorrencias seguidos de uma coluna Atividade.
    """
    if not avaliacoes:
        avaliacoes = [{}]
    aulas_por_aval = num_aulas // len(avaliacoes)
    resto = num_aulas % len(avaliacoes)
    colunas, aula_global = [], 1
    for idx, _ in enumerate(avaliacoes):
        n = aulas_por_aval + (1 if idx < resto else 0)
        for _ in range(n):
            colunas.append({"tipo": "nota",       "aula": aula_global, "avaliacao": idx + 1})
            colunas.append({"tipo": "ocorrencias","aula": aula_global, "avaliacao": idx + 1})
            aula_global += 1
        colunas.append({"tipo": "atividade", "aula": None, "avaliacao": idx + 1})
    return colunas


# ---------------------------------------------------------------------------
# Montagem dos ranges de cada trimestre
# ---------------------------------------------------------------------------

def _trimestre_ranges(nome_aba, turma_data, config, tri_num):
    """
    Retorna lista de {'range': ..., 'values': [...]} para a API batchUpdate.
    Usa a notação 'Nome Aba'!A1 esperada pelo Sheets API.
    """
    num_aulas  = config.get("num_aulas", 0)
    data_inicio = config.get("data_inicio")
    freq       = config.get("frequencia_semanal", 1)
    avaliacoes = config.get("avaliacoes", [])
    alunos     = sorted(turma_data.get("alunos", []), key=lambda a: a["numero"])

    datas   = _gerar_datas(data_inicio, num_aulas, freq) if data_inicio else []
    colunas = _colunas_aulas(num_aulas, avaliacoes)

    titulo = f"{turma_data['disciplina']} - {tri_num} TRI - {turma_data['turma']}"
    p = f"'{nome_aba}'!"   # prefixo de range

    ranges = []

    # ------------------------------------------------------------------
    # Linha 1 — título, label "Data" e datas das aulas
    # ------------------------------------------------------------------
    ranges.append({"range": f"{p}A1", "values": [[titulo]]})
    ranges.append({"range": f"{p}E1", "values": [["Data"]]})

    for i, col in enumerate(colunas):
        if col["tipo"] == "nota" and col["aula"] <= len(datas):
            cl = _col_letter(COL_AULAS_INICIO + i)
            ranges.append({"range": f"{p}{cl}1",
                           "values": [[datas[col["aula"] - 1].strftime("%d/%m")]]})

    # ------------------------------------------------------------------
    # Linha 2 — label "Tema da Aula" (professor preenche abaixo)
    # ------------------------------------------------------------------
    ranges.append({"range": f"{p}E2", "values": [["Tema da Aula"]]})

    # ------------------------------------------------------------------
    # Linha 3 — cabeçalhos
    # ------------------------------------------------------------------
    ranges.append({"range": f"{p}A3",
                   "values": [["Aluno", "", "Média 1", "Média 2", "Média 3", "Média", "Nº"]]})

    nota_cols_por_grupo = {}   # avaliacao → [col_letter, ...]
    ativ_col_por_grupo  = {}   # avaliacao → col_letter

    for i, col in enumerate(colunas):
        cl = _col_letter(COL_AULAS_INICIO + i)
        if col["tipo"] == "nota":
            ranges.append({"range": f"{p}{cl}3", "values": [["Nota"]]})
            nota_cols_por_grupo.setdefault(col["avaliacao"], []).append(cl)
        elif col["tipo"] == "ocorrencias":
            ranges.append({"range": f"{p}{cl}3", "values": [["Ocorrencias"]]})
        else:
            ranges.append({"range": f"{p}{cl}3", "values": [[f"Atividade {col['avaliacao']}"]]})
            ativ_col_por_grupo[col["avaliacao"]] = cl

    # ------------------------------------------------------------------
    # Linhas 4+ — número, nome e fórmulas por aluno
    # ------------------------------------------------------------------
    for idx, aluno in enumerate(alunos):
        row = 4 + idx

        ranges.append({"range": f"{p}G{row}", "values": [[aluno["numero"]]]})
        ranges.append({"range": f"{p}A{row}", "values": [[aluno["nome"]]]})

        # Nota = 100 + penalidade do VLOOKUP na célula Ocorrencias ao lado
        for i, col in enumerate(colunas):
            if col["tipo"] == "nota":
                nota_cl  = _col_letter(COL_AULAS_INICIO + i)
                ocorr_cl = _col_letter(COL_AULAS_INICIO + i + 1)
                formula = (
                    f'=SE(ÉCÉL.VAZIA({ocorr_cl}{row});""; '
                    f'100+PROCV({ocorr_cl}{row};Penalidades!A:B;2;0))'
                )
                ranges.append({"range": f"{p}{nota_cl}{row}", "values": [[formula]]})

        # Média por avaliação (C, D, E)
        for av_idx, av in enumerate(sorted(nota_cols_por_grupo)):
            if av_idx >= len(MEDIAS_COLS):
                break
            todas_cols = nota_cols_por_grupo[av] + (
                [ativ_col_por_grupo[av]] if av in ativ_col_por_grupo else []
            )
            args = ";".join(f"{c}{row}" for c in todas_cols)
            ranges.append({
                "range": f"{p}{MEDIAS_COLS[av_idx]}{row}",
                "values": [[f'=SEERRO(MÉDIA({args});"")']],
            })

        # Média geral (F) = média das Médias 1/2/3
        medias_usadas = [MEDIAS_COLS[i] for i in range(len(nota_cols_por_grupo))
                         if i < len(MEDIAS_COLS)]
        if medias_usadas:
            args = ";".join(f"{c}{row}" for c in medias_usadas)
            ranges.append({
                "range": f"{p}F{row}",
                "values": [[f'=SEERRO(MÉDIA({args});"")']],
            })

    return ranges


# ---------------------------------------------------------------------------
# Validação de dados — menu suspenso nas colunas Ocorrencias
# ---------------------------------------------------------------------------

SITUACOES_ATIVAS = {"", "ativo", "matriculado", "cursando"}


def _requests_ocultar_inativos(ws_id, alunos):
    """
    Retorna requests updateDimensionProperties para ocultar linhas de alunos
    com situação diferente de ativo/matriculado (ex: transferido, desistente).
    Alunos começam na linha 4 (índice 3, 0-based).
    """
    requests = []
    for idx, aluno in enumerate(alunos):
        situacao = aluno.get("situacao", "").strip().lower()
        if situacao in SITUACOES_ATIVAS:
            continue
        row_idx = 3 + idx  # linha 4 = índice 3 (0-based)
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId":    ws_id,
                    "dimension":  "ROWS",
                    "startIndex": row_idx,
                    "endIndex":   row_idx + 1,
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser",
            }
        })
    return requests


def _requests_validacao(ws_id, colunas, num_alunos):
    """
    Retorna requests setDataValidation para cada coluna 'Ocorrencias',
    apontando para o intervalo de ocorrências na aba Penalidades.
    Aplica da linha 4 (índice 3) até o último aluno.
    """
    ultima_linha = 4 + num_alunos  # endRowIndex é exclusivo
    intervalo_pen = f"Penalidades!$A$2:$A${1 + len(PENALIDADES)}"
    requests = []

    for i, col in enumerate(colunas):
        if col["tipo"] != "ocorrencias":
            continue
        col_idx = COL_AULAS_INICIO + i - 1  # 0-based

        requests.append({
            "setDataValidation": {
                "range": {
                    "sheetId":          ws_id,
                    "startRowIndex":    3,           # linha 4 (0-based)
                    "endRowIndex":      ultima_linha,
                    "startColumnIndex": col_idx,
                    "endColumnIndex":   col_idx + 1,
                },
                "rule": {
                    "condition": {
                        "type":   "ONE_OF_RANGE",
                        "values": [{"userEnteredValue": f"={intervalo_pen}"}],
                    },
                    "showCustomUi": True,
                    "strict":       False,
                },
            }
        })

    return requests


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def gerar_diario(turma_data, config, pasta_id):
    """
    Gera um diário de classe no Google Sheets.

    Args:
        turma_data: dict com escola, turma, disciplina, alunos
        config:     dict com num_aulas, data_inicio (date ou 'YYYY-MM-DD'),
                    frequencia_semanal, avaliacoes (lista de dicts com valor_maximo)
        pasta_id:   ID da pasta no Google Drive

    Returns:
        URL da planilha criada.
    """
    creds  = _get_creds()
    gc     = gspread.authorize(creds)
    drive  = build("drive",  "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    nome = _gerar_nome(turma_data)
    print(f"Criando planilha: {nome}")

    # Cria o arquivo na pasta correta via Drive API
    arquivo = drive.files().create(
        body={
            "name": nome,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [pasta_id],
        },
        fields="id",
    ).execute()
    sheet_id = arquivo["id"]

    planilha = gc.open_by_key(sheet_id)
    ws_default = planilha.sheet1  # aba criada automaticamente pelo Google

    # Cria as abas necessárias
    ws_pen  = planilha.add_worksheet("Penalidades",  rows=20,  cols=2)
    ws_tri1 = planilha.add_worksheet("1 Trimestre",  rows=200, cols=60)
    ws_tri2 = planilha.add_worksheet("2 Trimestre",  rows=200, cols=60)
    ws_tri3 = planilha.add_worksheet("3 Trimestre",  rows=200, cols=60)
    planilha.del_worksheet(ws_default)

    # Preenche Penalidades
    pen_data = [["Ocorrencia", "Penalidade"]] + [[o, p] for o, p in PENALIDADES]
    ws_pen.update(pen_data, "A1", value_input_option="USER_ENTERED")
    print("  Penalidades preenchidas")

    # Monta todos os ranges dos 3 trimestres e envia em um único batch
    all_ranges = []
    for tri_num in [1, 2, 3]:
        nome_aba = f"{tri_num} Trimestre"
        all_ranges.extend(_trimestre_ranges(nome_aba, turma_data, config, tri_num))

    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": all_ranges},
    ).execute()
    print(f"  {len(all_ranges)} ranges gravados nos 3 trimestres")

    # Adiciona menu suspenso nas colunas Ocorrencias dos 3 trimestres
    colunas = _colunas_aulas(config.get("num_aulas", 0), config.get("avaliacoes", []))
    num_alunos = len(turma_data.get("alunos", []))
    alunos_ord = sorted(turma_data.get("alunos", []), key=lambda a: a["numero"])
    format_requests = []
    for ws in [ws_tri1, ws_tri2, ws_tri3]:
        format_requests.extend(_requests_validacao(ws.id, colunas, num_alunos))
        format_requests.extend(_requests_ocultar_inativos(ws.id, alunos_ord))

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": format_requests},
    ).execute()
    print(f"  {len(format_requests)} requests de formatação aplicados")

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    print(f"Planilha criada: {url}")
    return url
