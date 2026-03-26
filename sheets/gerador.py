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

# Colunas fixas: A=Aluno B=Situação C=AV1 D=AV2 E=AV3 F=Média G=Nº
COL_FIXAS        = 7
COL_AULAS_INICIO = COL_FIXAS + 1   # H = 8
MEDIAS_COLS      = ["C", "D", "E"]

SITUACOES_ATIVAS = {"", "ativo", "matriculado", "cursando"}


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


def _fmt(n):
    """Formata número para uso em fórmula do Sheets (padrão pt-BR: vírgula decimal)."""
    s = f"{float(n):.10g}"
    return s.replace(".", ",")


def _parsear_turma(nome_turma):
    partes = [p.strip() for p in nome_turma.split(" - ")]
    if len(partes) >= 3:
        return partes[-3], partes[-2], partes[-1]
    if len(partes) == 2:
        return partes[0], "", partes[1]
    return nome_turma, "", ""


def _limpar(texto):
    return "".join(c for c in texto if c.isalnum() or c in ("_", "-")).strip()


def _gerar_nome(turma_data):
    escola    = _limpar(turma_data["escola"].split()[0])
    serie, turno, letra = _parsear_turma(turma_data["turma"])
    serie_limpa  = _limpar(serie.replace("ª", "").replace(" ", ""))
    disciplina   = _limpar(turma_data["disciplina"].replace(" ", "_"))
    return f"{escola}_{turno}_{serie_limpa}_{letra}_{disciplina}"


def _gerar_datas_por_semana(data_inicio, num_semanas, freq):
    """
    Retorna dict {semana(int): [date, ...]} com `freq` datas úteis por semana,
    avançando consecutivamente a partir de data_inicio.
    """
    if not data_inicio or freq <= 0:
        return {}
    if isinstance(data_inicio, str):
        data_inicio = date.fromisoformat(data_inicio)

    total = num_semanas * freq
    all_dates, atual = [], data_inicio
    while len(all_dates) < total:
        if atual.weekday() < 5:
            all_dates.append(atual)
        atual += timedelta(days=1)

    return {s: all_dates[(s - 1) * freq: s * freq]
            for s in range(1, num_semanas + 1)}


def _colunas_v2(num_semanas, freq, avaliacoes, modo):
    """
    Gera a lista de descritores de coluna para um trimestre.

    Modos
    -----
    completo / diario : por aula → 2 cols (ocorr + eng); na semana da AV → 1 col av
    simples           : apenas colunas de AV (sem registro diário)
    """
    av_por_semana = {av["semana"]: i for i, av in enumerate(avaliacoes)
                     if av.get("semana")}
    colunas       = []
    aula_global   = 1

    for semana in range(1, num_semanas + 1):
        if modo != "simples":
            for _ in range(freq):
                colunas.append({"tipo": "ocorr", "semana": semana, "aula": aula_global})
                colunas.append({"tipo": "eng",   "semana": semana, "aula": aula_global})
                aula_global += 1

        if semana in av_por_semana:
            colunas.append({"tipo": "av", "semana": semana,
                            "av_idx": av_por_semana[semana]})

    return colunas


# ---------------------------------------------------------------------------
# Montagem dos ranges de cada trimestre
# ---------------------------------------------------------------------------

def _trimestre_ranges(nome_aba, turma_data, config, tri_num):
    modo        = config.get("modo", "diario")
    num_semanas = config.get("num_semanas", 14)
    freq        = config.get("frequencia_semanal", 2) if modo != "simples" else 0
    data_inicio = config.get("data_inicio")
    avaliacoes  = config.get("avaliacoes", [])
    alunos      = sorted(turma_data.get("alunos", []), key=lambda a: a["numero"])

    datas_por_semana = _gerar_datas_por_semana(data_inicio, num_semanas, freq)
    colunas          = _colunas_v2(num_semanas, freq, avaliacoes, modo)

    titulo = f"{turma_data['disciplina']} - {tri_num} TRI - {turma_data['turma']}"
    p      = f"'{nome_aba}'!"
    ranges = []

    # ------------------------------------------------------------------
    # Linha 1 — título + datas / labels das AVs
    # ------------------------------------------------------------------
    ranges.append({"range": f"{p}A1", "values": [[titulo]]})

    for i, col in enumerate(colunas):
        cl = _col_letter(COL_AULAS_INICIO + i)
        if col["tipo"] == "ocorr":
            semana       = col["semana"]
            aula         = col["aula"]
            idx_na_semana = (aula - 1) % freq if freq > 0 else 0
            datas_sem    = datas_por_semana.get(semana, [])
            if idx_na_semana < len(datas_sem):
                ranges.append({"range": f"{p}{cl}1",
                               "values": [[datas_sem[idx_na_semana].strftime("%d/%m")]]})
        elif col["tipo"] == "av":
            av_nome = avaliacoes[col["av_idx"]]["nome"]
            ranges.append({"range": f"{p}{cl}1", "values": [[f"* {av_nome}"]]})

    # ------------------------------------------------------------------
    # Linha 2 — "Tema da Aula" na primeira coluna de dados
    # ------------------------------------------------------------------
    if colunas:
        first_cl = _col_letter(COL_AULAS_INICIO)
        ranges.append({"range": f"{p}{first_cl}2", "values": [["Tema da Aula"]]})

    # ------------------------------------------------------------------
    # Linha 3 — cabeçalhos fixos + cabeçalhos das colunas de aula
    # ------------------------------------------------------------------
    av_headers = [av["nome"] for av in avaliacoes[:3]]
    while len(av_headers) < 3:
        av_headers.append("")
    ranges.append({"range": f"{p}A3",
                   "values": [["Aluno", "Situação"] + av_headers + ["Nota", "Nº"]]})

    av_col_letters    = {}   # av_idx  → col_letter
    eng_cols_por_sem  = {}   # semana  → [col_letters das eng]

    for i, col in enumerate(colunas):
        cl = _col_letter(COL_AULAS_INICIO + i)
        if col["tipo"] == "ocorr":
            ranges.append({"range": f"{p}{cl}3", "values": [["Ocorrência"]]})
        elif col["tipo"] == "eng":
            ranges.append({"range": f"{p}{cl}3", "values": [["Engaj."]]})
            eng_cols_por_sem.setdefault(col["semana"], []).append(cl)
        elif col["tipo"] == "av":
            av = avaliacoes[col["av_idx"]]
            ranges.append({"range": f"{p}{cl}3", "values": [[f"{av['nome']} (0-10)"]]})
            av_col_letters[col["av_idx"]] = cl

    # Períodos de engajamento por AV (semanas desde a AV anterior até esta)
    prev_semana     = 0
    av_eng_periods  = {}
    for av_idx, av in enumerate(avaliacoes):
        av_sem = av.get("semana") or num_semanas
        cols   = []
        for s in range(prev_semana + 1, av_sem + 1):
            cols.extend(eng_cols_por_sem.get(s, []))
        av_eng_periods[av_idx] = cols
        prev_semana = av_sem

    # ------------------------------------------------------------------
    # Linhas 4+ — alunos
    # ------------------------------------------------------------------
    for idx, aluno in enumerate(alunos):
        row = 4 + idx
        ranges.append({"range": f"{p}G{row}", "values": [[aluno["numero"]]]})
        ranges.append({"range": f"{p}A{row}", "values": [[aluno["nome"]]]})
        situacao = aluno.get("situacao", "").strip() or "Regular"
        ranges.append({"range": f"{p}B{row}", "values": [[situacao]]})

        # Fórmula de engajamento: 100 + PROCV(ocorrência)
        for i, col in enumerate(colunas):
            if col["tipo"] == "eng":
                eng_cl   = _col_letter(COL_AULAS_INICIO + i)
                ocorr_cl = _col_letter(COL_AULAS_INICIO + i - 1)
                formula  = (
                    f'=SE(ÉCÉL.VAZIA({ocorr_cl}{row});"";'
                    f'100+PROCV({ocorr_cl}{row};Penalidades!A:B;2;0))'
                )
                ranges.append({"range": f"{p}{eng_cl}{row}", "values": [[formula]]})

        # Fórmulas das AVs (colunas C, D, E)
        for av_idx, av in enumerate(avaliacoes):
            if av_idx >= len(MEDIAS_COLS):
                break
            av_cl   = av_col_letters.get(av_idx)
            med_col = MEDIAS_COLS[av_idx]
            if not av_cl:
                continue

            valor_max = av["valor_maximo"]
            peso_eng  = av.get("peso_engajamento", 0.0)
            peso_av   = av.get("peso_avaliacao",  valor_max)
            eng_cols  = av_eng_periods.get(av_idx, [])

            if modo == "completo" and peso_eng > 0 and eng_cols:
                eng_args   = ";".join(f"{c}{row}" for c in eng_cols)
                eng_part   = f"SEERRO(MÉDIA({eng_args});0)/100*{_fmt(peso_eng)}"
                av_part    = f"{av_cl}{row}/10*{_fmt(peso_av)}"
                formula    = (
                    f'=SEERRO(SE(ÉCÉL.VAZIA({av_cl}{row});"";'
                    f'{eng_part}+{av_part});"")'
                )
            else:
                formula = (
                    f'=SEERRO(SE(ÉCÉL.VAZIA({av_cl}{row});"";'
                    f'{av_cl}{row}/10*{_fmt(valor_max)});"")'
                )
            ranges.append({"range": f"{p}{med_col}{row}", "values": [[formula]]})

        # Nota final (F) = soma das colunas C, D, E
        avs_usadas = [MEDIAS_COLS[i] for i in range(min(len(avaliacoes), 3))]
        if avs_usadas:
            args = ";".join(f"{c}{row}" for c in avs_usadas)
            ranges.append({"range": f"{p}F{row}",
                           "values": [[f'=SEERRO(SOMA({args});"")']]} )

    return ranges


# ---------------------------------------------------------------------------
# Formatação e validação
# ---------------------------------------------------------------------------

def _requests_ocultar_inativos(ws_id, alunos):
    requests = []
    for idx, aluno in enumerate(alunos):
        situacao = aluno.get("situacao", "").strip().lower()
        if situacao in SITUACOES_ATIVAS:
            continue
        row_idx = 3 + idx
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
    """Menu suspenso nas colunas 'Ocorrência'."""
    ultima_linha  = 4 + num_alunos
    intervalo_pen = f"Penalidades!$A$2:$A${1 + len(PENALIDADES)}"
    requests      = []

    for i, col in enumerate(colunas):
        if col["tipo"] != "ocorr":
            continue
        col_idx = COL_AULAS_INICIO + i - 1   # 0-based
        requests.append({
            "setDataValidation": {
                "range": {
                    "sheetId":          ws_id,
                    "startRowIndex":    3,
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
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        result = {}
        for i, linha in enumerate(linhas[3:], start=4):
            if not linha or not linha[0].strip():
                continue
            num_str = linha[6].strip() if len(linha) > 6 else ""
            if not num_str.isdigit():
                continue
            result[int(num_str)] = {
                "nome":     linha[0].strip(),
                "situacao": linha[1].strip() if len(linha) > 1 else "",
                "row":      i,
            }
        return result
    return {}


def comparar_alunos(planilha_id, alunos_json):
    sheet        = _ler_alunos_planilha(planilha_id)
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
                    "numero":          num,
                    "nome":            a["nome"],
                    "situacao_antiga": sit_antiga or "Regular",
                    "situacao_nova":   sit_nova  or "Regular",
                    "acao":            acao,
                })

    for num, a in sheet.items():
        if num not in json_por_num:
            removidos.append({"numero": num, "nome": a["nome"]})

    return {"novos": novos, "alterados": alterados, "removidos": removidos}


def adicionar_aluno(planilha_id, aluno):
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    situacao_texto = aluno.get("situacao", "").strip() or "Regular"

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        ultima = 3
        for i, linha in enumerate(linhas[3:], start=4):
            if linha and linha[0].strip():
                ultima = i
        nova = ultima + 1
        ws.update([[aluno["nome"]]],        f"A{nova}", value_input_option="USER_ENTERED")
        ws.update([[situacao_texto]],       f"B{nova}", value_input_option="USER_ENTERED")
        ws.update([[int(aluno["numero"])]], f"G{nova}", value_input_option="USER_ENTERED")


def atualizar_situacao(planilha_id, numero_chamada, nova_situacao):
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
    creds  = _get_creds()
    gc     = gspread.authorize(creds)
    sheets = build("sheets", "v4", credentials=creds)
    sh     = gc.open_by_key(planilha_id)

    requests = []
    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        for i, linha in enumerate(linhas[3:], start=4):
            if len(linha) > 6 and linha[6].strip().isdigit():
                if int(linha[6].strip()) == numero_chamada:
                    requests.append({
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId":    ws.id,
                                "dimension":  "ROWS",
                                "startIndex": i - 1,
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

_IGNORAR = {"", "Fez a atividade", "Falta"}

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
    Lê a planilha e retorna alunos com ocorrências relevantes na data informada.
    Suporta formato novo (data na col Ocorrência) e antigo (data na col Nota).
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    data_curta = data_str[:5]   # "DD/MM"

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue

        linhas = ws.get_all_values()
        if len(linhas) < 4:
            continue

        row1 = linhas[0]
        row3 = linhas[2]

        nota_idx = next(
            (i for i, v in enumerate(row1) if v.strip() == data_curta), None
        )
        if nota_idx is None:
            continue

        # Determina coluna de ocorrências (novo: data na ocorr; antigo: data na nota)
        header = row3[nota_idx].strip() if len(row3) > nota_idx else ""
        if header in ("Ocorrência", "Ocorrencias"):
            ocorr_idx = nota_idx
        else:
            ocorr_idx = nota_idx + 1
            for j in range(nota_idx + 1, min(nota_idx + 3, len(row3))):
                if row3[j].strip() in ("Ocorrência", "Ocorrencias"):
                    ocorr_idx = j
                    break

        resultado = []
        for linha in linhas[3:]:
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

        return resultado

    return []


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def gerar_diario(turma_data, config, pasta_id):
    """
    Gera um diário de classe no Google Sheets.

    Args:
        turma_data : dict com escola, turma, disciplina, alunos
        config     : dict com modo, frequencia_semanal, avaliacoes
                     (num_semanas e data_inicio são auto-preenchidos por trimestre)
        pasta_id   : ID da pasta no Google Drive

    Returns:
        {"url": ..., "id": ...}
    """
    from calendario.calendario_pr import TRIMESTRES, calcular_semanas_trimestre

    creds  = _get_creds()
    gc     = gspread.authorize(creds)
    drive  = build("drive",  "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    nome = _gerar_nome(turma_data)
    print(f"Criando planilha: {nome}")

    arquivo = drive.files().create(
        body={
            "name":     nome,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents":  [pasta_id],
        },
        fields="id",
    ).execute()
    sheet_id = arquivo["id"]

    planilha   = gc.open_by_key(sheet_id)
    ws_default = planilha.sheet1

    # Calcula largura máxima necessária para as abas
    modo    = config.get("modo", "diario")
    freq    = config.get("frequencia_semanal", 2) if modo != "simples" else 0
    n_avs   = len(config.get("avaliacoes", []))
    ano     = date.today().year
    max_sem = max(calcular_semanas_trimestre(t, ano) for t in [1, 2, 3])
    max_cols = COL_FIXAS + max_sem * freq * 2 + n_avs + 10

    ws_pen  = planilha.add_worksheet("Penalidades",  rows=20,  cols=2)
    ws_tri1 = planilha.add_worksheet("1 Trimestre",  rows=200, cols=max(80, max_cols))
    ws_tri2 = planilha.add_worksheet("2 Trimestre",  rows=200, cols=max(80, max_cols))
    ws_tri3 = planilha.add_worksheet("3 Trimestre",  rows=200, cols=max(80, max_cols))
    planilha.del_worksheet(ws_default)

    # Penalidades
    pen_data = [["Ocorrencia", "Penalidade"]] + [[o, p] for o, p in PENALIDADES]
    ws_pen.update(pen_data, "A1", value_input_option="USER_ENTERED")
    print("  Penalidades preenchidas")

    # Gera ranges com datas corretas para cada trimestre
    tri_info = TRIMESTRES.get(ano, {})
    all_ranges  = []
    tri_colunas = {}

    for tri_num, ws in [(1, ws_tri1), (2, ws_tri2), (3, ws_tri3)]:
        info        = tri_info.get(tri_num, {})
        num_semanas = calcular_semanas_trimestre(tri_num, ano)
        tri_config  = dict(config)
        tri_config["data_inicio"]  = info.get("inicio")
        tri_config["num_semanas"]  = num_semanas

        nome_aba = f"{tri_num} Trimestre"
        all_ranges.extend(_trimestre_ranges(nome_aba, turma_data, tri_config, tri_num))
        tri_colunas[tri_num] = _colunas_v2(
            num_semanas, freq, config.get("avaliacoes", []), modo
        )

    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": all_ranges},
    ).execute()
    print(f"  {len(all_ranges)} ranges gravados nos 3 trimestres")

    # Formatação: validação, ocultar inativos, congelar colunas
    alunos_ord = sorted(turma_data.get("alunos", []), key=lambda a: a["numero"])
    num_alunos = len(alunos_ord)
    format_requests = []

    for tri_num, ws in [(1, ws_tri1), (2, ws_tri2), (3, ws_tri3)]:
        colunas = tri_colunas[tri_num]
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
