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

# Colunas fixas: A=Aluno B=Situação C=ATV1 D=REC1 … X=ATVn Y=RECn Z=Nota AA=Nº
# O número exato depende de quantas avaliações foram configuradas (N avs → 2+N*2 colunas fixas)
# Use _col_fixas(n_avs) e _col_aulas_inicio(n_avs) em vez das constantes abaixo.
# As constantes abaixo são mantidas apenas para compatibilidade com funções que não recebem n_avs.
COL_FIXAS        = 10   # padrão legado (3 avs)
COL_AULAS_INICIO = COL_FIXAS + 1


def _col_fixas(n_avs):
    """Número total de colunas fixas para N avaliações: A, B + N pares ATV/REC + Nota + Nº."""
    return 2 + n_avs * 2 + 2


def _col_aulas_inicio(n_avs):
    """Índice 1-based da primeira coluna dinâmica (logo após as colunas fixas)."""
    return _col_fixas(n_avs) + 1


def _medias_cols(n_avs):
    """Lista de letras de coluna para ATV1,REC1,ATV2,REC2,...,ATVn,RECn (colunas C em diante)."""
    cols = []
    for i in range(n_avs):
        cols.append(_col_letter(3 + i * 2))      # ATV i+1
        cols.append(_col_letter(3 + i * 2 + 1))  # REC i+1
    return cols


def _nota_col(n_avs):
    """Letra da coluna Nota (penúltima coluna fixa)."""
    return _col_letter(_col_fixas(n_avs) - 1)


def _num_col(n_avs):
    """Letra da coluna Nº (última coluna fixa)."""
    return _col_letter(_col_fixas(n_avs))

SITUACOES_ATIVAS = {"", "ativo", "matriculado", "cursando"}

# Cabeçalhos reconhecidos para cada coluna (case-insensitive, strip)
_HEADERS_NOME    = {"nome", "aluno", "alunos"}
_HEADERS_SIT     = {"situação", "situacao", "sit.", "sit", "status"}
_HEADERS_NUM     = {"nº", "n°", "num", "número", "numero", "nro", "chamada"}


def _detectar_colunas_alunos(row3):
    """
    A partir da linha de cabeçalhos (row3), detecta os índices das colunas
    de nome, situação e número de chamada.

    Retorna dict com chaves 'nome', 'situacao', 'numero'.
    Qualquer coluna não encontrada retorna None.
    """
    nome_idx = sit_idx = num_idx = None
    for i, h in enumerate(row3):
        h_norm = h.strip().lower()
        if nome_idx is None and h_norm in _HEADERS_NOME:
            nome_idx = i
        if sit_idx is None and h_norm in _HEADERS_SIT:
            sit_idx = i
        if num_idx is None and h_norm in _HEADERS_NUM:
            num_idx = i
    return {"nome": nome_idx, "situacao": sit_idx, "numero": num_idx}


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

    n_avs           = len(avaliacoes)
    col_aulas_ini   = _col_aulas_inicio(n_avs)
    medias_cols     = _medias_cols(n_avs)   # [ATV1, REC1, ATV2, REC2, ...]
    nota_col        = _nota_col(n_avs)
    num_col         = _num_col(n_avs)

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
        cl = _col_letter(col_aulas_ini + i)
        if col["tipo"] == "ocorr":
            semana        = col["semana"]
            aula          = col["aula"]
            idx_na_semana = (aula - 1) % freq if freq > 0 else 0
            datas_sem     = datas_por_semana.get(semana, [])
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
        first_cl = _col_letter(col_aulas_ini)
        ranges.append({"range": f"{p}{first_cl}2", "values": [["Tema da Aula"]]})

    # ------------------------------------------------------------------
    # Linha 3 — cabeçalhos fixos (dinâmicos) + cabeçalhos das colunas de aula
    # ------------------------------------------------------------------
    fixed_headers = ["Aluno", "Situação"]
    for i, av in enumerate(avaliacoes):
        fixed_headers.append(f"ATV {i+1}")
        fixed_headers.append(f"REC {i+1}")
    fixed_headers += ["Nota", "Nº"]
    ranges.append({"range": f"{p}A3", "values": [fixed_headers]})

    av_col_letters   = {}   # av_idx → col_letter (coluna ATV na área dinâmica)
    rec_col_letters  = {}   # av_idx → col_letter (coluna REC na área dinâmica)
    eng_cols_por_sem = {}   # semana → [col_letters das eng]

    for i, col in enumerate(colunas):
        cl = _col_letter(col_aulas_ini + i)
        if col["tipo"] == "ocorr":
            ranges.append({"range": f"{p}{cl}3", "values": [["Ocorrência"]]})
        elif col["tipo"] == "eng":
            ranges.append({"range": f"{p}{cl}3", "values": [["Engaj."]]})
            eng_cols_por_sem.setdefault(col["semana"], []).append(cl)
        elif col["tipo"] == "av":
            ranges.append({"range": f"{p}{cl}3", "values": [[f"ATV {col['av_idx']+1} (0-10)"]]})
            av_col_letters[col["av_idx"]] = cl
        elif col["tipo"] == "rec":
            ranges.append({"range": f"{p}{cl}3", "values": [[f"REC {col['av_idx']+1} (0-10)"]]})
            rec_col_letters[col["av_idx"]] = cl

    # Períodos de engajamento por AV (semanas desde a AV anterior até esta)
    prev_semana    = 0
    av_eng_periods = {}
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
        ranges.append({"range": f"{p}{num_col}{row}", "values": [[aluno["numero"]]]})
        ranges.append({"range": f"{p}A{row}",         "values": [[aluno["nome"]]]})
        situacao = aluno.get("situacao", "").strip() or "Regular"
        ranges.append({"range": f"{p}B{row}", "values": [[situacao]]})

        # Fórmula de engajamento
        for i, col in enumerate(colunas):
            if col["tipo"] == "eng":
                eng_cl   = _col_letter(col_aulas_ini + i)
                ocorr_cl = _col_letter(col_aulas_ini + i - 1)
                formula  = (
                    f'=SE(ÉCÉL.VAZIA({ocorr_cl}{row});"";'
                    f'100+PROCV({ocorr_cl}{row};Penalidades!A:B;2;0))'
                )
                ranges.append({"range": f"{p}{eng_cl}{row}", "values": [[formula]]})

        # Fórmulas das AVs nas colunas fixas (ATV1→C, ATV2→E, ...)
        for av_idx, av in enumerate(avaliacoes):
            av_cl   = av_col_letters.get(av_idx)
            med_col = medias_cols[av_idx * 2]   # coluna ATV na área fixa
            if not av_cl:
                continue

            peso_eng = av.get("peso_engajamento", 0.0)
            eng_cols = av_eng_periods.get(av_idx, [])

            if modo == "completo" and peso_eng > 0 and eng_cols:
                eng_args = ";".join(f"{c}{row}" for c in eng_cols)
                eng_part = f"SEERRO(MÉDIA({eng_args});0)/10*{_fmt(peso_eng)}"
                formula  = (
                    f'=SEERRO(SE(ÉCÉL.VAZIA({av_cl}{row});"";'
                    f'ARRED({eng_part}+{av_cl}{row};0));"")'
                )
            else:
                formula = (
                    f'=SEERRO(SE(ÉCÉL.VAZIA({av_cl}{row});"";'
                    f'ARRED({av_cl}{row};0));"")'
                )
            ranges.append({"range": f"{p}{med_col}{row}", "values": [[formula]]})

        # Nota final = soma de todas as AVs, substituindo por REC quando preenchida
        partes_nota = []
        for av_idx in range(n_avs):
            atv_col = medias_cols[av_idx * 2]       # ATV i
            rec_col = medias_cols[av_idx * 2 + 1]   # REC i
            partes_nota.append(
                f'SE(ÉCÉL.VAZIA({rec_col}{row});SEERRO({atv_col}{row};0);{rec_col}{row})'
            )
        if partes_nota:
            soma_formula = "+".join(partes_nota)
            ranges.append({"range": f"{p}{nota_col}{row}",
                           "values": [[f'=SEERRO({soma_formula};"")']]})\

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


def _requests_cores_cabecalho(ws_id, colunas, num_alunos, n_avs):
    """Coloração das colunas ATV (azul claro) e REC (laranja claro) na linha 3 e dados."""
    requests = []
    ultima_linha  = 4 + num_alunos
    col_aulas_ini = _col_aulas_inicio(n_avs)

    _AZUL   = {"red": 0.776, "green": 0.871, "blue": 0.953}  # #C6DEFF aprox
    _LARANJA = {"red": 0.988, "green": 0.871, "blue": 0.706}  # #FCDED4 aprox

    for i, col in enumerate(colunas):
        if col["tipo"] not in ("av", "rec"):
            continue
        col_idx = col_aulas_ini + i - 1   # 0-based
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


def _requests_validacao(ws_id, colunas, num_alunos, n_avs):
    """Menu suspenso nas colunas 'Ocorrência'."""
    ultima_linha  = 4 + num_alunos
    intervalo_pen = "Penalidades!$A$2:$A$200"
    col_aulas_ini = _col_aulas_inicio(n_avs)
    requests      = []

    for i, col in enumerate(colunas):
        if col["tipo"] != "ocorr":
            continue
        col_idx = col_aulas_ini + i - 1   # 0-based
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


def _aliases_av(coluna_av):
    """
    Retorna lista de prefixos (uppercase) aceitos para a coluna_av.
    Suporta qualquer N: "ATV 4" → ["ATV 4"], "REC 4" → ["REC 4"],
    e mantém compatibilidade com o formato antigo "AV1"/"AV2"/"AV3".
    """
    mapa_legado = {
        "ATV 1": ("ATV 1", "AV1"),
        "REC 1": ("REC 1",),
        "ATV 2": ("ATV 2", "AV2"),
        "REC 2": ("REC 2",),
        "ATV 3": ("ATV 3", "AV3"),
        "REC 3": ("REC 3",),
    }
    return [p.upper() for p in mapa_legado.get(coluna_av, (coluna_av,))]


def get_avaliacoes_planilha(planilha_id, trimestre):
    """
    Lê a linha 3 da aba do trimestre e retorna a lista de avaliações disponíveis.

    Retorna lista de dicts: [{av: "ATV 1", rec: "REC 1"}, {av: "ATV 2", rec: "REC 2"}, ...]
    com quantas avaliações existirem na planilha (não limitado a 3).
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    nome_aba = _NOMES_ABA.get(int(trimestre))
    ws = next((w for w in sh.worksheets() if w.title == nome_aba), None)
    if ws is None:
        return []

    row3 = ws.row_values(3)

    # Detecta coluna Nº
    num_col_idx = next(
        (i for i, h in enumerate(row3) if h.strip() in ("Nº", "N°")),
        9
    )

    avaliacoes = []
    i = 2
    while i < num_col_idx:
        h = row3[i].strip().upper() if i < len(row3) else ""
        if h.startswith("ATV "):
            n = h.split()[1]   # "ATV 3 (0-10)" → "3"
            av_key  = f"ATV {n}"
            rec_key = f"REC {n}"
            avaliacoes.append({"av": av_key, "rec": rec_key})
        i += 1

    return avaliacoes


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

    # Detecta coluna Nº pelo header (sem assumir posição fixa)
    NUM_COL_IDX = next(
        (i for i, h in enumerate(row3) if h.strip() in ("Nº", "N°")),
        9   # fallback: J (planilhas com 3 avs)
    )

    # Busca a coluna da avaliação em toda a linha 3 (sem restringir intervalo)
    prefixos = _aliases_av(coluna_av)

    av_col_idx = None
    for idx, header in enumerate(row3):
        if any(header.strip().upper().startswith(p) for p in prefixos):
            av_col_idx = idx
            break

    if av_col_idx is None:
        raise ValueError(
            f"Coluna '{coluna_av}' não encontrada na aba '{nome_aba}'. "
            f"Cabeçalhos disponíveis: {row3}"
        )

    cols = _detectar_colunas_alunos(row3)
    nome_col_idx = cols["nome"] if cols["nome"] is not None else 0
    sit_col_idx  = cols["situacao"]

    resultado = []
    for linha in linhas[3:]:
        if not linha or not linha[nome_col_idx].strip():
            continue
        nome  = linha[nome_col_idx].strip()
        sit   = ""
        if sit_col_idx is not None and len(linha) > sit_col_idx:
            sit = linha[sit_col_idx].strip().lower()
        num_s = linha[NUM_COL_IDX].strip() if len(linha) > NUM_COL_IDX else ""
        if not num_s.isdigit():
            continue

        inativo = sit in _SITUACOES_INATIVAS

        nota_raw = linha[av_col_idx].strip() if len(linha) > av_col_idx else ""
        nota_raw = nota_raw.replace(",", ".")   # normaliza decimal

        if nota_raw:
            try:
                v = float(nota_raw)
                # Se o valor já é inteiro >= 10 (escala 0-100), usa direto.
                # Se é decimal < 10 (escala 0-10), converte ×10.
                if v < 10 and not nota_raw.isdigit():
                    nota_rco = str(int(round(v * 10)))
                else:
                    nota_rco = str(int(round(v)))
            except ValueError:
                nota_rco = "0"
        else:
            nota_rco = "0"

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
    row3 = ws.row_values(3)

    # Detecta onde começam as colunas dinâmicas: logo após a coluna "Nº"
    num_col_idx = next(
        (i for i, h in enumerate(row3) if h.strip() in ("Nº", "N°")),
        9   # fallback: J (planilhas com 3 avs)
    )
    aulas_inicio_idx = num_col_idx + 1   # 0-based

    datas = []
    for cell in row1[aulas_inicio_idx:]:   # a partir da primeira coluna dinâmica
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

    ABAS_TRIMESTRE = {"1 Trimestre", "2 Trimestre", "3 Trimestre"}
    for ws in sh.worksheets():
        if ws.title not in ABAS_TRIMESTRE:
            continue
        linhas = ws.get_all_values()
        if len(linhas) < 3:
            continue

        row3 = linhas[2]
        cols = _detectar_colunas_alunos(row3)
        num_col_idx  = cols["numero"]
        nome_col_idx = cols["nome"]
        sit_col_idx  = cols["situacao"]

        if num_col_idx is None:
            continue
        # Nome: usa coluna detectada ou fallback para coluna 0
        if nome_col_idx is None:
            nome_col_idx = 0

        result = {}
        for i, linha in enumerate(linhas[3:], start=4):
            if not linha or not linha[nome_col_idx].strip():
                continue
            num_str = linha[num_col_idx].strip() if len(linha) > num_col_idx else ""
            if not num_str.isdigit():
                continue
            sit = ""
            if sit_col_idx is not None and len(linha) > sit_col_idx:
                sit = linha[sit_col_idx].strip()
            result[int(num_str)] = {
                "nome":     linha[nome_col_idx].strip(),
                "situacao": sit,
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
        if len(linhas) < 3:
            continue
        cols = _detectar_colunas_alunos(linhas[2])
        num_col_idx  = cols["numero"]
        nome_col_idx = cols["nome"]   if cols["nome"]    is not None else 0
        sit_col_idx  = cols["situacao"]

        # Encontra última linha com conteúdo na coluna nome
        ultima = 3
        for i, linha in enumerate(linhas[3:], start=4):
            if linha and len(linha) > nome_col_idx and linha[nome_col_idx].strip():
                ultima = i
        nova = ultima + 1

        col_nome = gspread.utils.rowcol_to_a1(nova, nome_col_idx + 1)[:-len(str(nova))]
        ws.update([[aluno["nome"]]], f"{col_nome}{nova}", value_input_option="USER_ENTERED")
        if sit_col_idx is not None:
            col_sit = gspread.utils.rowcol_to_a1(nova, sit_col_idx + 1)[:-len(str(nova))]
            ws.update([[situacao_texto]], f"{col_sit}{nova}", value_input_option="USER_ENTERED")
        if num_col_idx is not None:
            col_num = gspread.utils.rowcol_to_a1(nova, num_col_idx + 1)[:-len(str(nova))]
            ws.update([[int(aluno["numero"])]], f"{col_num}{nova}", value_input_option="USER_ENTERED")


def atualizar_situacao(planilha_id, numero_chamada, nova_situacao):
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)
    texto = nova_situacao.strip() or "Regular"

    for ws in sh.worksheets():
        if ws.title == "Penalidades":
            continue
        linhas = ws.get_all_values()
        if len(linhas) < 3:
            continue
        cols = _detectar_colunas_alunos(linhas[2])
        num_col_idx = cols["numero"]
        sit_col_idx = cols["situacao"]
        if num_col_idx is None or sit_col_idx is None:
            continue
        col_sit = gspread.utils.rowcol_to_a1(1, sit_col_idx + 1)[:-1]
        for i, linha in enumerate(linhas[3:], start=4):
            if len(linha) > num_col_idx and linha[num_col_idx].strip().isdigit():
                if int(linha[num_col_idx].strip()) == numero_chamada:
                    ws.update([[texto]], f"{col_sit}{i}", value_input_option="USER_ENTERED")
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
        if len(linhas) < 3:
            continue
        cols = _detectar_colunas_alunos(linhas[2])
        num_col_idx = cols["numero"]
        if num_col_idx is None:
            continue
        for i, linha in enumerate(linhas[3:], start=4):
            if len(linha) > num_col_idx and linha[num_col_idx].strip().isdigit():
                if int(linha[num_col_idx].strip()) == numero_chamada:
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
    "Muita Conversa":      "Comportamento inadequado durante a aula",
    "Celular":             "Uso de celular durante a aula",
    "Dormindo":            "Dormindo durante a aula",
    "Evasão(Gaseio)":      "Evasão durante a aula",
}

_NOMES_ABA_TRI = {
    "1 Trimestre": "1T",
    "2 Trimestre": "2T",
    "3 Trimestre": "3T",
}


def get_ocorrencias_periodo(planilha_id, data_inicio, data_fim):
    """
    Lê todas as abas de trimestre e retorna ocorrências relevantes no período.

    data_inicio, data_fim: "DD/MM/AAAA"

    Retorna lista de dicts:
        {data, trimestre, numero_chamada, nome, ocorrencia, comentario}
    """
    from datetime import datetime

    def _parse(d):
        return datetime.strptime(d, "%d/%m/%Y")

    dt_ini = _parse(data_inicio)
    dt_fim = _parse(data_fim)

    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    resultado = []

    for ws in sh.worksheets():
        tri_label = _NOMES_ABA_TRI.get(ws.title)
        if not tri_label:
            continue

        linhas = ws.get_all_values()
        if len(linhas) < 4:
            continue

        row1 = linhas[0]
        row3 = linhas[2]

        # Detecta coluna do número de chamada pelo header "Nº"
        num_col_idx = next(
            (i for i, h in enumerate(row3) if h.strip() in ("Nº", "N°")),
            6   # fallback: coluna G (planilhas antigas)
        )

        # Encontra colunas de Ocorrência cujas datas estão no período
        ocorr_cols = []   # lista de (col_idx, data_str)
        for idx, header in enumerate(row3):
            if header.strip() not in ("Ocorrência", "Ocorrencias"):
                continue
            data_cell = row1[idx].strip() if idx < len(row1) else ""
            if not data_cell or len(data_cell) != 5:
                continue
            data_completa = data_cell + f"/{datetime.today().year}"
            try:
                dt = _parse(data_completa)
            except ValueError:
                continue
            if dt_ini <= dt <= dt_fim:
                ocorr_cols.append((idx, data_completa))

        for ocorr_idx, data_str in ocorr_cols:
            for linha in linhas[3:]:
                if not linha or not linha[0].strip():
                    continue
                nome   = linha[0].strip()
                numero = linha[num_col_idx].strip() if len(linha) > num_col_idx else ""
                ocorr  = linha[ocorr_idx].strip() if len(linha) > ocorr_idx else ""

                if ocorr in _IGNORAR:
                    continue

                resultado.append({
                    "data":           data_str,
                    "trimestre":      tri_label,
                    "numero_chamada": int(numero) if numero.isdigit() else None,
                    "nome":           nome,
                    "ocorrencia":     ocorr,
                    "comentario":     _COMENTARIOS.get(ocorr, ocorr),
                })

    return resultado


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

        # Detecta coluna do número de chamada pelo header "Nº"
        num_col_idx = next(
            (i for i, h in enumerate(row3) if h.strip() in ("Nº", "N°")),
            6   # fallback: coluna G (planilhas antigas)
        )

        cols = _detectar_colunas_alunos(row3)
        nome_col_idx_oc = cols["nome"] if cols["nome"] is not None else 0

        resultado = []
        for linha in linhas[3:]:
            if not linha or not linha[nome_col_idx_oc].strip():
                continue
            nome   = linha[nome_col_idx_oc].strip()
            numero = linha[num_col_idx].strip() if len(linha) > num_col_idx else ""
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
def reaplicar_validacao(planilha_id):
    """
    Reaplicar o dropdown de ocorrências em todas as abas de trimestre da planilha,
    usando o intervalo aberto Penalidades!$A$2:$A$200 para suportar novas entradas.
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    ABAS = {"1 Trimestre", "2 Trimestre", "3 Trimestre"}
    requests = []

    for ws in sh.worksheets():
        if ws.title not in ABAS:
            continue
        linhas = ws.get_all_values()
        if len(linhas) < 3:
            continue

        row3 = linhas[2]
        num_alunos = sum(
            1 for l in linhas[3:]
            if l and any(c.strip() for c in l)
        )
        ultima_linha = 4 + num_alunos

        # Detecta colunas de ocorrência pelo cabeçalho (contém "ocorr" ou "ocorrência")
        for col_idx, header in enumerate(row3):
            h = header.strip().lower()
            if "ocorr" in h or "ocorrencia" in h or "ocorrência" in h:
                requests.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId":          ws.id,
                            "startRowIndex":    3,
                            "endRowIndex":      ultima_linha,
                            "startColumnIndex": col_idx,
                            "endColumnIndex":   col_idx + 1,
                        },
                        "rule": {
                            "condition": {
                                "type":   "ONE_OF_RANGE",
                                "values": [{"userEnteredValue": "=Penalidades!$A$2:$A$200"}],
                            },
                            "showCustomUi": True,
                            "strict":       False,
                        },
                    }
                })

    if requests:
        sh.batch_update({"requests": requests})

    return len(requests)


# Aba Resumo
# ---------------------------------------------------------------------------

def diagnosticar_planilha(planilha_id):
    """
    Analisa uma planilha externa e retorna quais módulos foram encontrados.

    Retorna dict:
        nome_planilha  (str)
        abas           (list of str)  — abas encontradas
        modulos        (list of dict) — {nome, encontrado, detalhe}
    """
    creds = _get_creds()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(planilha_id)

    nome_planilha = sh.title
    abas = [ws.title for ws in sh.worksheets() if ws.title != "Penalidades"]

    modulos = []

    for aba_titulo in abas:
        ws     = sh.worksheet(aba_titulo)
        linhas = ws.get_all_values()

        if len(linhas) < 3:
            modulos.append({
                "nome":      f"Aba '{aba_titulo}'",
                "encontrado": False,
                "detalhe":   "Menos de 3 linhas — sem cabeçalho na linha 3",
            })
            continue

        row3 = linhas[2]
        cols = _detectar_colunas_alunos(row3)

        # Módulo: número de chamada
        modulos.append({
            "nome":       f"[{aba_titulo}] Nº de chamada",
            "encontrado": cols["numero"] is not None,
            "detalhe":    f"Coluna {cols['numero']+1}" if cols["numero"] is not None
                          else "Cabeçalho 'Nº'/'N°' não encontrado na linha 3",
        })

        # Módulo: nome do aluno
        modulos.append({
            "nome":       f"[{aba_titulo}] Nome do aluno",
            "encontrado": cols["nome"] is not None,
            "detalhe":    f"Coluna {cols['nome']+1}" if cols["nome"] is not None
                          else "Cabeçalho 'Nome'/'Aluno' não encontrado na linha 3",
        })

        # Módulo: situação
        modulos.append({
            "nome":       f"[{aba_titulo}] Situação",
            "encontrado": cols["situacao"] is not None,
            "detalhe":    f"Coluna {cols['situacao']+1}" if cols["situacao"] is not None
                          else "Cabeçalho 'Situação'/'Status' não encontrado (opcional)",
        })

        # Módulo: AVs — detecta todas as colunas com prefixo ATV/REC/AV
        import re as _re
        avs_encontradas = [
            h.strip() for h in row3
            if _re.match(r'(ATV|REC|AV)\s*\d+', h.strip(), _re.IGNORECASE)
        ]
        modulos.append({
            "nome":       f"[{aba_titulo}] Avaliações (ATV/REC)",
            "encontrado": len(avs_encontradas) > 0,
            "detalhe":    ", ".join(avs_encontradas) if avs_encontradas
                          else "Nenhuma coluna ATV/REC encontrada na linha 3",
        })

        # Módulo: ocorrências
        tem_ocorr = any(
            h.strip() in ("Ocorrência", "Ocorrencias", "Ocorrência")
            for h in row3
        )
        modulos.append({
            "nome":       f"[{aba_titulo}] Ocorrências",
            "encontrado": tem_ocorr,
            "detalhe":    "Coluna 'Ocorrência' encontrada" if tem_ocorr
                          else "Coluna 'Ocorrência' não encontrada",
        })

    return {
        "nome_planilha": nome_planilha,
        "abas":          abas,
        "modulos":       modulos,
    }


def adicionar_aba_resumo(planilha_id, notas_finais):
    """
    Cria ou atualiza a aba "Resumo" na planilha com os dados do RCO.

    notas_finais: lista de dicts retornada por buscar_notas_finais_rco()
        {numero, nome, situacao, soma}  ← formato real (1 trimestre por vez)
      OU lista expandida com 1T/2T/3T quando chamado com dados de múltiplos trimestres.

    Retorna {aprovados, reprovados, parciais, total_alunos}.
    """
    creds  = _get_creds()
    gc     = gspread.authorize(creds)
    sheets = build("sheets", "v4", credentials=creds)
    sh     = gc.open_by_key(planilha_id)

    # Cria ou limpa a aba Resumo
    ws_resumo = next((w for w in sh.worksheets() if w.title == "Resumo"), None)
    if ws_resumo is None:
        ws_resumo = sh.add_worksheet("Resumo", rows=200, cols=10)
        # Move para índice 0 (primeira aba)
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=planilha_id,
            body={"requests": [{"updateSheetProperties": {
                "properties": {"sheetId": ws_resumo.id, "index": 0},
                "fields": "index",
            }}]},
        ).execute()
    else:
        ws_resumo.clear()

    # ------------------------------------------------------------------
    # Monta dados
    # ------------------------------------------------------------------
    cabecalho = ["Nº", "Nome", "Situação", "1º Trimestre", "2º Trimestre",
                 "3º Trimestre", "Total", "Situação Final"]

    def _situacao_final(t1, t2, t3):
        notas = [n for n in (t1, t2, t3) if n is not None]
        if not notas:
            return "Parcial"
        total = sum(notas)
        if len(notas) < 3:
            return "Parcial"
        if total >= 18 and all(n >= 6 for n in notas):
            return "Aprovado"
        return "Reprovado"

    linhas  = [cabecalho]
    stats   = {"aprovados": 0, "reprovados": 0, "parciais": 0}

    for a in notas_finais:
        t1  = a.get("1T")
        t2  = a.get("2T")
        t3  = a.get("3T")
        # Compatibilidade com formato de 1 trimestre (campo "soma")
        if t1 is None and t2 is None and t3 is None and "soma" in a:
            t1 = a["soma"]

        notas_disp = [n for n in (t1, t2, t3) if n is not None]
        total      = sum(notas_disp) if notas_disp else None
        sit_final  = _situacao_final(t1, t2, t3)

        stats[{"Aprovado": "aprovados", "Reprovado": "reprovados",
               "Parcial": "parciais"}[sit_final]] += 1

        linhas.append([
            a.get("numero", ""),
            a.get("nome", ""),
            a.get("situacao", ""),
            t1 if t1 is not None else "",
            t2 if t2 is not None else "",
            t3 if t3 is not None else "",
            total if total is not None else "",
            sit_final,
        ])

    ws_resumo.update(linhas, "A1", value_input_option="USER_ENTERED")

    # ------------------------------------------------------------------
    # Formatação via batchUpdate
    # ------------------------------------------------------------------
    sid      = ws_resumo.id
    n_alunos = len(notas_finais)

    def _rgb(hex_str):
        h = hex_str.lstrip("#")
        return {
            "red":   int(h[0:2], 16) / 255,
            "green": int(h[2:4], 16) / 255,
            "blue":  int(h[4:6], 16) / 255,
        }

    requests = []

    # Cabeçalho: fundo #1a2744, texto branco, negrito
    requests.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": 8},
        "cell": {"userEnteredFormat": {
            "backgroundColor": _rgb("1a2744"),
            "textFormat": {"foregroundColor": _rgb("ffffff"), "bold": True},
            "horizontalAlignment": "CENTER",
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
    }})

    # Linhas alternadas
    for i in range(n_alunos):
        cor = "ffffff" if i % 2 == 0 else "f8fafc"
        requests.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": 1 + i, "endRowIndex": 2 + i,
                      "startColumnIndex": 0, "endColumnIndex": 8},
            "cell": {"userEnteredFormat": {"backgroundColor": _rgb(cor)}},
            "fields": "userEnteredFormat.backgroundColor",
        }})

    # Cor da coluna Situação Final (col H = índice 7) por valor
    _COR_SIT = {
        "Aprovado":  ("d1fae5", "065f46"),
        "Reprovado": ("fee2e2", "991b1b"),
        "Parcial":   ("fef3c7", "92400e"),
    }
    for i, a in enumerate(notas_finais):
        t1 = a.get("1T")
        t2 = a.get("2T")
        t3 = a.get("3T")
        if t1 is None and t2 is None and t3 is None and "soma" in a:
            t1 = a["soma"]
        sit = _situacao_final(t1, t2, t3)
        bg, fg = _COR_SIT[sit]
        requests.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": 1 + i, "endRowIndex": 2 + i,
                      "startColumnIndex": 7, "endColumnIndex": 8},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _rgb(bg),
                "textFormat": {"foregroundColor": _rgb(fg), "bold": True},
                "horizontalAlignment": "CENTER",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
        }})

    # Larguras das colunas: Nº=50 Nome=220 Sit=100 3×Tri=110 Total=80 SitFinal=130
    larguras = [50, 220, 100, 110, 110, 110, 80, 130]
    for col_idx, px in enumerate(larguras):
        requests.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": col_idx, "endIndex": col_idx + 1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }})

    # Congela linha 1
    requests.append({"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount",
    }})

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=planilha_id,
        body={"requests": requests},
    ).execute()

    return {
        "total_alunos": n_alunos,
        **stats,
    }


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
    max_cols = _col_fixas(n_avs) + max_sem * freq * 2 + n_avs + 10

    ws_pen  = planilha.add_worksheet("Penalidades",  rows=200, cols=2)
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

    col_fixas_n = _col_fixas(n_avs)
    for tri_num, ws in [(1, ws_tri1), (2, ws_tri2), (3, ws_tri3)]:
        colunas = tri_colunas[tri_num]
        format_requests.extend(_requests_validacao(ws.id, colunas, num_alunos, n_avs))
        format_requests.extend(_requests_cores_cabecalho(ws.id, colunas, num_alunos, n_avs))
        format_requests.extend(_requests_ocultar_inativos(ws.id, alunos_ord))
        format_requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": ws.id,
                    "gridProperties": {
                        "frozenColumnCount": col_fixas_n,
                        "frozenRowCount":    3,
                    },
                },
                "fields": "gridProperties.frozenColumnCount,gridProperties.frozenRowCount",
            }
        })
        # Colunas C até Nº (todas as colunas fixas exceto A e B) → 45 px
        format_requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId":    ws.id,
                    "dimension":  "COLUMNS",
                    "startIndex": 2,            # C (0-based)
                    "endIndex":   col_fixas_n,  # até Nº inclusive
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
