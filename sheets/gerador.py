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
                   "values": [["Aluno", "Situação", "ATV 1", "ATV 2", "ATV 3", "Média", "Nº"]]})

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
        situacao_texto = aluno.get("situacao", "").strip() or "Regular"
        ranges.append({"range": f"{p}B{row}", "values": [[situacao_texto]]})

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
# Sincronização de alunos
# ---------------------------------------------------------------------------

def _ler_alunos_planilha(planilha_id):
    """
    Lê os alunos da primeira aba de trimestre.
    Retorna dict: {numero: {nome, situacao, row_1based}}
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        result = {}
        for i, linha in enumerate(linhas[3:], start=4):   # row 4 = alunos
            if not linha or not linha[0].strip():
                continue
            num_str = linha[6].strip() if len(linha) > 6 else ""
            if not num_str.isdigit():
                continue
            result[int(num_str)] = {
                "nome":      linha[0].strip(),
                "situacao":  linha[1].strip() if len(linha) > 1 else "",
                "row":       i,   # 1-based
            }
        return result
    return {}


def comparar_alunos(planilha_id, alunos_json):
    """
    Compara dados.json com a planilha e retorna o que precisa ser sincronizado.
    Returns: {novos, alterados, removidos}
      novos     — alunos em dados.json que não estão na planilha
      alterados — alunos em ambos, mas situação diferente
                  cada item tem {numero, nome, situacao_antiga, situacao_nova, acao}
                  acao: "ocultar" (situação inativa) | "atualizar" (ainda ativa)
      removidos — alunos na planilha que não estão mais em dados.json
    """
    sheet = _ler_alunos_planilha(planilha_id)
    json_por_num = {a["numero"]: a for a in alunos_json}

    novos, alterados, removidos = [], [], []

    for num, a in json_por_num.items():
        if num not in sheet:
            novos.append({"numero": num, "nome": a["nome"],
                          "situacao": a.get("situacao", "")})
        else:
            sit_nova   = a.get("situacao", "").strip()
            sit_antiga = sheet[num]["situacao"]
            if sit_antiga == "Regular":
                sit_antiga = ""
            if sit_nova != sit_antiga:
                acao = ("ocultar"
                        if sit_nova and sit_nova.lower() not in SITUACOES_ATIVAS
                        else "atualizar")
                alterados.append({
                    "numero":         num,
                    "nome":           a["nome"],
                    "situacao_antiga": sit_antiga or "Regular",
                    "situacao_nova":   sit_nova  or "Regular",
                    "acao":           acao,
                })

    for num, a in sheet.items():
        if num not in json_por_num:
            removidos.append({"numero": num, "nome": a["nome"]})

    return {"novos": novos, "alterados": alterados, "removidos": removidos}


def adicionar_aluno(planilha_id, aluno):
    """
    Adiciona uma nova linha de aluno (nome, situação, número) em todas as abas
    de trimestre, logo após o último aluno existente.
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    situacao_texto = aluno.get("situacao", "").strip() or "Regular"

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        ultima = 3   # 1-based, row 3 = headers
        for i, linha in enumerate(linhas[3:], start=4):
            if linha and linha[0].strip():
                ultima = i
        nova = ultima + 1
        ws.update([[aluno["nome"]]],           f"A{nova}", value_input_option="USER_ENTERED")
        ws.update([[situacao_texto]],          f"B{nova}", value_input_option="USER_ENTERED")
        ws.update([[int(aluno["numero"])]],    f"G{nova}", value_input_option="USER_ENTERED")


def atualizar_situacao(planilha_id, numero_chamada, nova_situacao):
    """Atualiza a coluna B (Situação) do aluno em todas as abas de trimestre."""
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    texto = nova_situacao.strip() or "Regular"

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        for i, linha in enumerate(linhas[3:], start=4):
            if len(linha) > 6 and linha[6].strip().isdigit():
                if int(linha[6].strip()) == numero_chamada:
                    ws.update([[texto]], f"B{i}", value_input_option="USER_ENTERED")
                    break


def ocultar_aluno(planilha_id, numero_chamada):
    """Oculta a linha do aluno em todas as abas de trimestre."""
    creds  = _get_creds()
    gc     = gspread.authorize(creds)
    sheets = build("sheets", "v4", credentials=creds)
    sh     = gc.open_by_key(planilha_id)

    requests = []
    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        for i, linha in enumerate(linhas[3:], start=4):   # i é 1-based
            if len(linha) > 6 and linha[6].strip().isdigit():
                if int(linha[6].strip()) == numero_chamada:
                    requests.append({
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId":    ws.id,
                                "dimension":  "ROWS",
                                "startIndex": i - 1,   # 0-based
                                "endIndex":   i,
                            },
                            "properties": {"hiddenByUser": True},
                            "fields": "hiddenByUser",
                        }
                    })
                    break

    if requests:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=planilha_id,
            body={"requests": requests},
        ).execute()


# ---------------------------------------------------------------------------
# Leitura de ocorrências
# ---------------------------------------------------------------------------

# Ocorrências que NÃO geram comentário (presença normal ou ausência)
_IGNORAR = {"", "Fez a atividade", "Falta"}

# Mapeamento ocorrência → frase de comentário para o diário
_COMENTARIOS = {
    "Não fez a atividade": "Não realizou a atividade proposta",
    "Não terminou":        "Não concluiu a atividade proposta",
    "Entregou com atraso": "Entregou a atividade com atraso",
    "Muita Conversa":      "Conversa excessiva durante a aula",
    "Celular":             "Uso de celular durante a aula",
    "Dormindo":            "Dormindo durante a aula",
    "Evasão(Gaseio)":      "Evasão de sala de aula (gaseio)",
}


def ler_ocorrencias_planilha(planilha_id, data_str):
    """
    Lê a planilha e retorna os alunos que têm ocorrências relevantes
    na coluna correspondente à data informada.

    Args:
        planilha_id: ID do Google Sheets
        data_str:    string "DD/MM/AAAA"

    Returns:
        lista de dicts {numero, nome, ocorrencia, comentario}
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    # Datas na planilha estão no formato "DD/MM" (sem ano), na linha 1
    data_curta = data_str[:5]

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue

        linhas = ws.get_all_values()
        if len(linhas) < 4:
            continue

        row1 = linhas[0]   # datas (DD/MM) nas colunas de nota
        row3 = linhas[2]   # cabeçalhos: Aluno, Situação, ATV1..., Nota, Ocorrencias...

        # Encontra a coluna da nota para a data alvo
        nota_idx = next(
            (i for i, v in enumerate(row1) if v.strip() == data_curta), None
        )
        if nota_idx is None:
            continue

        # Coluna de Ocorrências = coluna seguinte à Nota
        ocorr_idx = nota_idx + 1
        # Confirmação: verifica cabeçalho
        if len(row3) > ocorr_idx and row3[ocorr_idx].strip() != "Ocorrencias":
            # Procura "Ocorrencias" nas próximas 2 colunas
            for j in range(nota_idx + 1, min(nota_idx + 3, len(row3))):
                if row3[j].strip() == "Ocorrencias":
                    ocorr_idx = j
                    break

        resultado = []
        for linha in linhas[3:]:           # linhas de alunos (row 4+)
            if not linha or not linha[0].strip():
                continue
            nome   = linha[0].strip()
            numero = linha[6].strip() if len(linha) > 6 else ""
            ocorr  = linha[ocorr_idx].strip() if len(linha) > ocorr_idx else ""

            if ocorr in _IGNORAR:
                continue

            resultado.append({
                "numero":     int(numero) if numero.isdigit() else numero,
                "nome":       nome,
                "ocorrencia": ocorr,
                "comentario": _COMENTARIOS.get(ocorr, ocorr),
            })

        return resultado   # data encontrada nesta aba

    return []   # data não encontrada em nenhuma aba


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
        format_requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": ws.id,
                    "gridProperties": {"frozenColumnCount": COL_FIXAS},
                },
                "fields": "gridProperties.frozenColumnCount",
            }
        })

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": format_requests},
    ).execute()
    print(f"  {len(format_requests)} requests de formatação aplicados")

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    print(f"Planilha criada: {url}")
    return {"url": url, "id": sheet_id}
