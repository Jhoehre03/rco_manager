"""
Calendário escolar do Paraná (SEED-PR).
Fonte: Resolução n.º 6.494/2025-GS/SEED — Calendário Escolar 2026.
"""
from datetime import date, timedelta

# ── Períodos dos trimestres ──────────────────────────────────────────────────
TRIMESTRES = {
    2026: {
        1: {"inicio": date(2026, 2,  5), "fim": date(2026, 5, 14), "dias_letivos": 63},
        2: {"inicio": date(2026, 5, 18), "fim": date(2026, 9,  4), "dias_letivos": 68},
        3: {"inicio": date(2026, 9,  7), "fim": date(2026, 12,18), "dias_letivos": 70},
    }
}

# ── Dias não letivos 2026 ────────────────────────────────────────────────────
# Feriados nacionais + recesso de Páscoa + jornada pedagógica de início do ano
_DIAS_NAO_LETIVOS_2026: set[date] = {
    # Jornada pedagógica (3 primeiros dias úteis do ano letivo)
    date(2026, 2,  5), date(2026, 2,  6), date(2026, 2,  9),
    # Feriados
    date(2026, 4,  3),   # Sexta-feira da Paixão
    date(2026, 4, 21),   # Tiradentes
    date(2026, 5,  1),   # Dia do Trabalho
    date(2026, 6,  4),   # Corpus Christi
    date(2026, 6,  5),   # Emenda de Corpus Christi (ponte)
    date(2026, 9,  7),   # Independência do Brasil
    date(2026, 10,28),   # Dia do Servidor Público Estadual (Paraná)
    date(2026, 10,12),   # N. Sra. Aparecida
    date(2026, 11, 2),   # Finados
    date(2026, 11,15),   # Proclamação da República (cai domingo, mas reservado)
    date(2026, 11,20),   # Zumbi e Consciência Negra
    date(2026, 12,25),   # Natal
    # Recesso de Páscoa (segunda e terça após o domingo de Páscoa)
    date(2026, 4,  6), date(2026, 4,  7),
}

# Recesso de julho: 13 a 24/07/2026
_RECESSOS_2026: list[tuple[date, date]] = [
    (date(2026, 7, 13), date(2026, 7, 24)),
]


def _e_dia_letivo(d: date, ano: int) -> bool:
    """Retorna True se 'd' é um dia útil letivo (não é fim de semana, feriado ou recesso)."""
    if d.weekday() >= 5:          # sábado ou domingo
        return False

    nao_letivos = {2026: _DIAS_NAO_LETIVOS_2026}.get(ano, set())
    if d in nao_letivos:
        return False

    recessos = {2026: _RECESSOS_2026}.get(ano, [])
    for inicio, fim in recessos:
        if inicio <= d <= fim:
            return False

    return True


def calcular_aulas(trimestre: int, dias_semana: list[int], ano: int | None = None) -> int:
    """
    Calcula o número de aulas no trimestre para um professor que leciona
    nos dias da semana indicados.

    Args:
        trimestre:   1, 2 ou 3
        dias_semana: lista de inteiros 0=segunda…4=sexta
                     (ex: [0, 2, 4] = segunda, quarta e sexta)
        ano:         ano letivo (padrão: ano atual)

    Returns:
        Número de aulas (dias letivos que coincidem com os dias do professor).
    """
    if ano is None:
        ano = date.today().year

    dados = TRIMESTRES.get(ano, {}).get(trimestre)
    if not dados:
        return 0

    dias_set = set(dias_semana)
    count = 0
    atual = dados["inicio"]
    fim   = dados["fim"]

    while atual <= fim:
        if atual.weekday() in dias_set and _e_dia_letivo(atual, ano):
            count += 1
        atual += timedelta(days=1)

    return count


def obter_trimestre_atual(ano: int | None = None) -> int | None:
    """Retorna o número (1, 2 ou 3) do trimestre atual, ou None se estiver fora do período letivo."""
    if ano is None:
        ano = date.today().year

    hoje  = date.today()
    dados = TRIMESTRES.get(ano, {})
    for t, info in dados.items():
        if info["inicio"] <= hoje <= info["fim"]:
            return t
    return None


def calcular_semanas_trimestre(trimestre: int, ano: int | None = None) -> int:
    """Retorna o número de semanas no trimestre (ceil dos dias / 7)."""
    import math
    if ano is None:
        ano = date.today().year
    dados = TRIMESTRES.get(ano, {}).get(trimestre)
    if not dados:
        return 0
    return math.ceil((dados["fim"] - dados["inicio"]).days / 7)


def obter_info_trimestre(trimestre: int, ano: int | None = None) -> dict:
    """
    Retorna um dict com informações do trimestre:
      - inicio_str: "DD/MM/AAAA"
      - fim_str:    "DD/MM/AAAA"
      - dias_letivos: int
    """
    if ano is None:
        ano = date.today().year

    dados = TRIMESTRES.get(ano, {}).get(trimestre)
    if not dados:
        return {}

    return {
        "inicio_str":   dados["inicio"].strftime("%d/%m/%Y"),
        "fim_str":      dados["fim"].strftime("%d/%m/%Y"),
        "dias_letivos": dados["dias_letivos"],
    }
