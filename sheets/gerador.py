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

# Colunas fixas: A=Aluno B=Situação C=ATV1 D=REC1 E=ATV2 F=REC2 G=ATV3 H=REC3 I=Nota J=Nº
COL_FIXAS        = 10
COL_AULAS_INICIO = COL_FIXAS + 1   # K = 11
MEDIAS_COLS      = ["C", "D", "E", "F", "G", "H"]   # ATV1,REC1,ATV2,REC2,ATV3,REC3

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
            av_idx = av_por_semana[semana]
            colunas.append({"tipo": "av",  "semana": semana, "av_idx": av_idx})
            colunas.append({"tipo": "rec", "semana": semana, "av_idx": av_idx})

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
    # C=ATV1, D=REC1, E=ATV2, F=REC2, G=ATV3, H=REC3, I=Nota, J=Nº
    av_names = [av["nome"] for av in avaliacoes[:3]]
    while len(av_names) < 3:
        av_names.append("")
    fixed_headers = ["Aluno", "Situação"]
    for n in av_names:
        fixed_headers.append(n)           # ATV N
        fixed_headers.append(f"REC {n[-1]}" if n else "")  # REC N
    fixed_headers += ["Nota", "Nº"]
    ranges.append({"range": f"{p}A3", "values": [fixed_headers]})

    av_col_letters     = {}   # av_idx  → col_letter  (coluna da nota ATV na área dinâmica)
    rec_col_letters    = {}   # av_idx  → col_letter  (coluna da nota REC na área dinâmica)
    eng_cols_por_sem   = {}   # semana  → [col_letters das eng]

    for i, col in enumerate(colunas):
        cl = _col_letter(COL_AULAS_INICIO + i)
        if col["tipo"] == "ocorr":
            ranges.append({"range": f"{p}{cl}3", "values": [["Ocorrência"]]})
        elif col["tipo"] == "eng":
            ranges.append({"range": f"{p}{cl}3", "values": [["Engaj."]]})
            eng_cols_por_sem.setdefault(col["semana"], []).append(cl)
        elif col["tipo"] == "av":
            av = avaliacoes[col["av_idx"]]
            ranges.append({"range": f"{p}{cl}3", "values": [[f"ATV {col['av_idx']+1} (0-10)"]]})
            av_col_letters[col["av_idx"]] = cl
        elif col["tipo"] == "rec":
            ranges.append({"range": f"{p}{cl}3", "values": [[f"REC {col['av_idx']+1} (0-10)"]]})
            rec_col_letters[col["av_idx"]] = cl

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
        num_col = _col_letter(COL_FIXAS)   # J quando COL_FIXAS=10
        ranges.append({"range": f"{p}{num_col}{row}", "values": [[aluno["numero"]]]})
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

        # Fórmulas das AVs (colunas C, E, G = ATV1, ATV2, ATV3)
        # Colunas D, F, H (REC1, REC2, REC3) ficam em branco para preenchimento manual
        for av_idx, av in enumerate(avaliacoes):
            av_cl    = av_col_letters.get(av_idx)
            # med_col index: ATV1→C(0), ATV2→E(2), ATV3→G(4)
            med_col_idx = av_idx * 2
            if med_col_idx >= len(MEDIAS_COLS):
                break
            med_col = MEDIAS_COLS[med_col_idx]
            if not av_cl:
                continue

            peso_eng  = av.get("peso_engajamento", 0.0)
            eng_cols  = av_eng_periods.get(av_idx, [])

            if modo == "completo" and peso_eng > 0 and eng_cols:
                eng_args = ";".join(f"{c}{row}" for c in eng_cols)
                # eng: média/10 * peso_eng  (100% eng com peso 2 → 20)
                # av:  nota lançada pelo professor já no valor correto
                eng_part = f"SEERRO(MÉDIA({eng_args});0)/10*{_fmt(peso_eng)}"
                formula  = (
                    f'=SEERRO(SE(ÉCÉL.VAZIA({av_cl}{row});"";'
                    f'ARRED({eng_part}+{av_cl}{row};0));"")'
                )
            else:
                # Nota direta — professor já lança o valor correto
                formula = (
                    f'=SEERRO(SE(ÉCÉL.VAZIA({av_cl}{row});"";'
                    f'ARRED({av_cl}{row};0));"")'
                )
            ranges.append({"range": f"{p}{med_col}{row}", "values": [[formula]]})

        # Nota final (I) = soma de ATV1+ATV2+ATV3 (C, E, G)
        # Considera REC: se REC preenchida, usa REC no lugar da ATV correspondente
        partes_nota = []
        for av_idx in range(min(len(avaliacoes), 3)):
            atv_col = MEDIAS_COLS[av_idx * 2]      # C, E, G
            rec_col = MEDIAS_COLS[av_idx * 2 + 1]  # D, F, H
            # SE REC preenchida usa REC, senão usa ATV
            partes_nota.append(
                f'SE(ÉCÉL.VAZIA({rec_col}{row});SEERRO({atv_col}{row};0);{rec_col}{row})'
            )
        if partes_nota:
            soma_formula = "+".join(partes_nota)
            ranges.append({"range": f"{p}I{row}",
                           "values": [[f'=SEERRO({soma_formula};"")']]})

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


def _requests_cores_cabecalho(ws_id, colunas, num_alunos):
    """Coloração das colunas ATV (azul claro) e REC (laranja claro) na linha 3 e dados."""
    requests = []
    ultima_linha = 4 + num_alunos

    _AZUL   = {"red": 0.776, "green": 0.871, "blue": 0.953}  # #C6DEFF aprox
    _LARANJA = {"red": 0.988, "green": 0.871, "blue": 0.706}  # #FCDED4 aprox

    for i, col in enumerate(colunas):
        if col["tipo"] not in ("av", "rec"):
            continue
        col_idx = COL_AULAS_INICIO + i - 1   # 0-based
        cor = _AZUL if col["tipo"] == "av" else _LARANJA
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId":          ws_id,
                    "startRowIndex":    2,          # linha 3 (header)
                    "endRowIndex":      ultima_linha,
                    "startColumnIndex": col_idx,
                    "endColumnIndex":   col_idx + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": cor
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
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
# Leitura de notas
# ---------------------------------------------------------------------------

_NOMES_ABA = {1: "1 Trimestre", 2: "2 Trimestre", 3: "3 Trimestre"}

# Mapeamento tipo_av → nome do radio no RCO
# Chave: o que o frontend envia; valor: o que preencher_formulario_avaliacao espera
AV_TIPOS = {
    "AV1": "AV1", "AV2": "AV2", "AV3": "AV3",
    "Recuperação 1": "Recuperação", "Recuperação 2": "Recuperação",
    "Recuperação 3": "Recuperação",
}

_SITUACOES_INATIVAS = {"transferido", "desistente", "cancelado", "evadido", "afastado"}


def ler_notas_planilha(planilha_id, trimestre, coluna_av):
    """
    Lê a coluna de nota calculada (C, D ou E) referente a `coluna_av` na aba do trimestre.

    coluna_av: nome da AV conforme cabeçalho do Sheets — ex. "AV1", "AV2", "AV3"
               ou prefixo que combine com o header (ex. "AV1" bate com "AV1 (0-10)").

    Retorna lista de dicts:
        numero   (int)   — número de chamada
        nome     (str)
        nota     (str)   — valor calculado da planilha (ex. "2.5") ou "" se vazio
        nota_rco (str)   — nota × 10 como inteiro com zero-padding (ex. "25"), "" se sem nota
        inativo  (bool)
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    nome_aba = _NOMES_ABA.get(int(trimestre))
    ws = next((w for w in sh.worksheets() if w.title == nome_aba), None)
    if ws is None:
        raise ValueError(f"Aba '{nome_aba}' não encontrada na planilha.")

    linhas = ws.get_all_values()
    if len(linhas) < 4:
        return []

    row3 = linhas[2]   # índice 2 = linha 3 (cabeçalhos)

    # Descobre a coluna que corresponde a coluna_av nas colunas fixas C–H (índices 2–7)
    # Aceita tanto o formato novo ("ATV 1") quanto o antigo ("AV1", "AV2", "AV3")
    _ALIASES = {
        "ATV 1": ("ATV 1", "AV1"),
        "REC 1": ("REC 1",),
        "ATV 2": ("ATV 2", "AV2"),
        "REC 2": ("REC 2",),
        "ATV 3": ("ATV 3", "AV3"),
        "REC 3": ("REC 3",),
    }
    prefixos = [p.upper() for p in _ALIASES.get(coluna_av, (coluna_av,))]

    av_col_idx = None
    for idx in range(2, min(10, len(row3))):   # C=2 … J=9
        header = row3[idx].strip().upper()
        if any(header.startswith(p) for p in prefixos):
            av_col_idx = idx
            break

    if av_col_idx is None:
        raise ValueError(
            f"Coluna '{coluna_av}' não encontrada na aba '{nome_aba}'. "
            f"Cabeçalhos disponíveis: {row3[2:10]}"
        )

    # Número de chamada: novo formato → coluna J (índice 9); antigo → coluna G (índice 6)
    # Detecta pelo cabeçalho "Nº" na linha 3
    NUM_COL_IDX = next(
        (i for i, h in enumerate(row3) if h.strip() in ("Nº", "N°", "Nº")),
        COL_FIXAS - 1   # fallback: J
    )

    resultado = []
    for linha in linhas[3:]:
        if not linha or not linha[0].strip():
            continue
        nome   = linha[0].strip()
        sit    = linha[1].strip().lower() if len(linha) > 1 else ""
        num_s  = linha[NUM_COL_IDX].strip() if len(linha) > NUM_COL_IDX else ""
        if not num_s.isdigit():
            continue

        inativo = sit in _SITUACOES_INATIVAS

        nota_raw = linha[av_col_idx].strip() if len(linha) > av_col_idx else ""
        nota_raw = nota_raw.replace(",", ".")   # normaliza decimal

        if nota_raw:
            try:
                nota_rco = str(int(round(float(nota_raw))))
            except ValueError:
                nota_rco = ""
        else:
            nota_rco = ""

        resultado.append({
            "numero":   int(num_s),
            "nome":     nome,
            "nota":     nota_raw,
            "nota_rco": nota_rco,
            "inativo":  inativo,
        })

    return resultado


# ---------------------------------------------------------------------------
# Datas de aula
# ---------------------------------------------------------------------------

def get_datas_aula(planilha_id, trimestre):
    """
    Lê a linha 1 da aba do trimestre e retorna as datas de aula preenchidas.
    Retorna lista de strings "DD/MM/AAAA", ordenadas cronologicamente.
    """
    from datetime import datetime

    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    nome_aba = _NOMES_ABA.get(int(trimestre))
    ws = next((w for w in sh.worksheets() if w.title == nome_aba), None)
    if ws is None:
        raise ValueError(f"Aba '{nome_aba}' não encontrada na planilha.")

    row1 = ws.row_values(1)

    datas = []
    for cell in row1[COL_AULAS_INICIO - 1:]:   # a partir da coluna H (índice 7)
        v = cell.strip()
        if not v:
            continue
        # Aceita "DD/MM" (gerado pelo gerador) ou "DD/MM/AAAA"
        if len(v) == 5 and v[2] == "/":
            v = v + f"/{datetime.today().year}"
        try:
            datetime.strptime(v, "%d/%m/%Y")
            datas.append(v)
        except ValueError:
            continue

    # Ordena cronologicamente e remove duplicatas
    datas = sorted(set(datas), key=lambda d: datetime.strptime(d, "%d/%m/%Y"))
    return datas


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
        format_requests.extend(_requests_cores_cabecalho(ws.id, colunas, num_alunos))
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
        # Colunas C–J (ATV1/REC1/ATV2/REC2/ATV3/REC3/Nota/Nº) → 45 px
        format_requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId":    ws.id,
                    "dimension":  "COLUMNS",
                    "startIndex": 2,   # C
                    "endIndex":   10,  # até J inclusive
                },
                "properties": {"pixelSize": 45},
                "fields": "pixelSize",
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
