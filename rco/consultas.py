"""
Consultas de leitura à API do RCO.
Todas as funções usam rco_client (RateLimiter + audit log).
Escrita continua via Selenium — este módulo é somente leitura.
"""
import datetime
from rco.auth import rco_client


# ── Turmas ────────────────────────────────────────────────────────────────────

def get_turmas_do_dia(data: str) -> list[dict]:
    """
    Retorna as turmas do educador para uma data específica.

    Args:
        data: formato "YYYY-MM-DD"

    Returns:
        Lista de dicts com as chaves relevantes já extraídas:
        [{"codClasse", "codPeriodoAvaliacao", "codPeriodoLetivo", "turma", "escola", "disciplina"}, ...]
    """
    dados = rco_client.get(f"/educador/grade/dia/v2/{data}")
    if not isinstance(dados, list):
        return []

    vistos = set()
    resultado = []
    for item in dados:
        classe = item.get("classe", {})
        turma  = classe.get("turma", {})
        estab  = turma.get("estabelecimento", {})
        disc   = classe.get("disciplina", {})
        # codPeriodoAvaliacao está na raiz do item, não dentro de classe
        pa     = item.get("periodoAvaliacao", {})
        pl     = turma.get("periodoLetivo", {})

        cod_classe = classe.get("codClasse")
        cod_pa     = pa.get("codPeriodoAvaliacao")

        # Deduplica: mesma turma pode aparecer várias vezes (uma por aula do dia)
        chave = (cod_classe, cod_pa)
        if chave in vistos:
            continue
        vistos.add(chave)

        resultado.append({
            "codClasse":           cod_classe,
            "codPeriodoAvaliacao": cod_pa,
            "codPeriodoLetivo":    pl.get("codPeriodoLetivo"),
            "turma":               turma.get("descrTurma", ""),
            "escola":              estab.get("nomeCompletoEstab", ""),
            "disciplina":          disc.get("descrDisciplina", ""),
        })
    return resultado


def get_todas_turmas() -> list[dict]:
    """
    Retorna todas as turmas do educador via /educador/estabelecimentos/v2/{data}.
    Uma única chamada retorna todos os estabelecimentos, turmas e períodos avaliativos.

    Returns:
        [{"codClasse", "codPeriodoAvaliacao", "codPeriodoLetivo", "turma", "escola",
          "disciplina", "trimestre"}, ...]
        Uma entrada por (codClasse, codPeriodoAvaliacao) — ou seja, por turma/trimestre.
    """
    hoje = datetime.date.today().isoformat()
    dados = rco_client.get(f"/educador/estabelecimentos/v2/{hoje}")
    if not isinstance(dados, list):
        return []

    resultado = []
    for estab in dados:
        escola = estab.get("nomeCompletoEstab", "")
        for pl in estab.get("periodoLetivos", []):
            cod_pl = pl.get("codPeriodoLetivo")
            for livro in pl.get("livros", []):
                classe = livro.get("classe", {})
                turma  = classe.get("turma", {})
                disc   = classe.get("disciplina", {})
                cod_classe = classe.get("codClasse")

                for cal in livro.get("calendarioAvaliacaos", []):
                    pa = cal.get("periodoAvaliacao", {})
                    resultado.append({
                        "codClasse":           cod_classe,
                        "codPeriodoAvaliacao": pa.get("codPeriodoAvaliacao"),
                        "codPeriodoLetivo":    cod_pl,
                        "turma":               turma.get("descrTurma", ""),
                        "escola":              escola,
                        "disciplina":          disc.get("nomeDisciplina", ""),
                        "trimestre":           pa.get("descrPeriodoAvaliacao", ""),
                    })
    return resultado


# ── Avaliações da turma ───────────────────────────────────────────────────────

def get_config_avaliacoes(cod_classe: int, cod_periodo: int) -> dict:
    """
    Retorna a configuração de avaliações da turma:
    regra de cálculo (ex: Somatória) e quantidade de AVs.

    Returns:
        {"codAvaliacaoClasse", "regraCalculo", "qtdeAvaliacao"}
    """
    return rco_client.get(
        "/classe/v1/avaliacaoClasses",
        params={"codClasse": cod_classe, "codPeriodoAvaliacao": cod_periodo},
    )


def get_avaliacoes_parciais(
    cod_classe: int,
    cod_periodo: int,
    cod_regra_calculo: int,
    qtde_avaliacao: int,
) -> list[dict]:
    """
    Retorna as AVs configuradas na turma: nome, data, peso e recuperações.

    Returns:
        [{"codAvaliacaoParcialClasse", "descrAvaliacaoParcial", "pesoDecimal",
          "dataAvaliacaoParcial", "recuperacaos": [...]}, ...]
    """
    dados = rco_client.get(
        "/classe/v1/avaliacaoParcialClasses",
        params={
            "codClasse":          cod_classe,
            "codPeriodoAvaliacao": cod_periodo,
            "codRegraCalculo":    cod_regra_calculo,
            "qtdeAvaliacao":      qtde_avaliacao,
            "page":               1,
            "perPage":            50,
        },
    )
    return dados if isinstance(dados, list) else []


# ── Notas dos alunos ──────────────────────────────────────────────────────────

def get_notas_alunos(cod_classe: int, cod_periodo: int) -> list[dict]:
    """
    Retorna as notas lançadas por aluno na turma.

    O JSON da API usa chaves dinâmicas "nota{codAvaliacaoParcialClasse}"
    para cada AV. Esta função normaliza para uma estrutura fixa:

    Returns:
        [
          {
            "codMatrizAluno": int,
            "numChamada": int,
            "nome": str,
            "situacao": str,
            "notas": {"55881012": "1.7", "55977955": "1.7"},
            "final": str,
          },
          ...
        ]
    """
    dados = rco_client.get(
        "/classe/v1/relatorios/avaliacaoParcialAlunos",
        params={"codClasse": cod_classe, "codPeriodoAvaliacao": cod_periodo},
    )
    if not isinstance(dados, list):
        return []

    resultado = []
    for aluno in dados:
        notas = {
            k.replace("nota", ""): v
            for k, v in aluno.items()
            if k.startswith("nota")
        }
        resultado.append({
            "codMatrizAluno": aluno.get("codMatrizAluno"),
            "numChamada":     aluno.get("numChamada"),
            "nome":           aluno.get("nome", ""),
            "situacao":       aluno.get("descrAbrevSituacaoMatricula", ""),
            "notas":          notas,
            "final":          aluno.get("final", ""),
        })
    return resultado


def get_notas_finais(cod_classe: int, cod_periodo: int) -> list[dict]:
    """
    Retorna a nota final do trimestre por aluno (campo "final" calculado pelo RCO).
    Equivalente ao que o Selenium lia na coluna "Somatória" da aba Alunos.

    Returns:
        [{"numero": int, "nome": str, "situacao": str, "soma": float|None}, ...]
    """
    dados = rco_client.get(
        "/classe/v1/relatorios/avaliacaoParcialAlunos",
        params={"codClasse": cod_classe, "codPeriodoAvaliacao": cod_periodo},
    )
    if not isinstance(dados, list):
        return []

    def _parse(valor):
        if not valor or valor == "-":
            return None
        try:
            return float(str(valor).replace(",", "."))
        except ValueError:
            return None

    return [
        {
            "numero":   a.get("numChamada"),
            "nome":     a.get("nome", ""),
            "situacao": a.get("descrAbrevSituacaoMatricula", ""),
            "soma":     _parse(a.get("final")),
        }
        for a in dados
        if str(a.get("numChamada", "")).isdigit() or a.get("numChamada") is not None
    ]


def get_datas_aula(cod_classe: int, cod_periodo: int) -> list[str]:
    """
    Retorna as datas de aulas dadas no trimestre, formato "DD/MM/YYYY", ordem decrescente.
    Substitui o Selenium que varria o calendário visual do RCO.
    """
    dados = rco_client.get(
        "/educador/grade/aula/v2",
        params={
            "codClasse":           cod_classe,
            "codPeriodoAvaliacao": cod_periodo,
            "page":                1,
            "perPage":             200,
        },
    )
    if not isinstance(dados, list):
        return []

    datas = set()
    for aula in dados:
        data_raw = (aula.get("dataAula") or "")[:10]  # "2026-03-04"
        if data_raw:
            p = data_raw.split("-")
            if len(p) == 3:
                datas.add(f"{p[2]}/{p[1]}/{p[0]}")

    # Ordena corretamente: converte DD/MM/YYYY → YYYY-MM-DD para sort, depois volta
    return sorted(datas, key=lambda d: d[6:]+d[3:5]+d[:2], reverse=True)


# ── Alunos e frequência ───────────────────────────────────────────────────────

def get_alunos(cod_classe: int, cod_periodo: int, cod_periodo_letivo: int) -> list[dict]:
    """
    Retorna a lista de alunos matriculados na turma com situação de matrícula.
    Substitui get_alunos() do Selenium.

    Returns:
        [{"codMatrizAluno", "numChamada", "nome", "situacao"}, ...]
    """
    dados = rco_client.get(
        "/classe/v3/relatorios/frequenciaAulas",
        params={
            "codClasse":          cod_classe,
            "codPeriodoAvaliacao": cod_periodo,
            "codPeriodoLetivo":   cod_periodo_letivo,
            "page":               1,
            "perPage":            200,  # turmas grandes
        },
    )
    if not isinstance(dados, list):
        return []

    return [
        {
            "codMatrizAluno": a.get("codMatrizAluno"),
            "numChamada":     a.get("numChamada"),
            "nome":           a.get("nome", ""),
            "situacao":       a.get("descrAbrevSituacaoMatricula", ""),
        }
        for a in dados
    ]


def get_frequencia_alunos(
    cod_classe: int, cod_periodo: int, cod_periodo_letivo: int
) -> list[dict]:
    """
    Retorna alunos com frequência detalhada por aula.
    Cada aluno tem "faltas": dict {codAula: "C"|"F"}.

    Returns:
        [
          {
            "codMatrizAluno": int,
            "numChamada": int,
            "nome": str,
            "situacao": str,
            "frequencia": {"368842847": "C", "369225307": "F", ...},
            "total_faltas": int,
          },
          ...
        ]
    """
    dados = rco_client.get(
        "/classe/v3/relatorios/frequenciaAulas",
        params={
            "codClasse":          cod_classe,
            "codPeriodoAvaliacao": cod_periodo,
            "codPeriodoLetivo":   cod_periodo_letivo,
            "page":               1,
            "perPage":            200,
        },
    )
    if not isinstance(dados, list):
        return []

    CAMPOS_FIXOS = {"codMatrizAluno", "numChamada", "nome", "descrAbrevSituacaoMatricula"}
    resultado = []
    for a in dados:
        freq = {k: v for k, v in a.items() if k not in CAMPOS_FIXOS}
        resultado.append({
            "codMatrizAluno": a.get("codMatrizAluno"),
            "numChamada":     a.get("numChamada"),
            "nome":           a.get("nome", ""),
            "situacao":       a.get("descrAbrevSituacaoMatricula", ""),
            "frequencia":     freq,
            "total_faltas":   sum(1 for v in freq.values() if v == "F"),
        })
    return resultado


# ── Aulas dadas ───────────────────────────────────────────────────────────────

def get_aulas_dadas(cod_classe: int, cod_periodo: int) -> list[dict]:
    """
    Retorna as aulas registradas na turma.

    Returns:
        [{"codAula", "numAula", "dataAula", "qtdeAlunos", "situacao"}, ...]
    """
    dados = rco_client.get(
        "/educador/grade/aula/v2",
        params={
            "codClasse":          cod_classe,
            "codPeriodoAvaliacao": cod_periodo,
            "page":               1,
            "perPage":            50,
        },
    )
    if not isinstance(dados, list):
        return []

    return [
        {
            "codAula":    a.get("codAula"),
            "numAula":    a.get("numAula"),
            "dataAula":   a.get("dataAula", "")[:10],  # remove horário
            "qtdeAlunos": a.get("qtdeC"),
            "situacao":   a.get("situacao"),
        }
        for a in dados
    ]


# ── Utilitário: snapshot completo de uma turma ────────────────────────────────

def get_snapshot_turma(
    cod_classe: int, cod_periodo: int, cod_periodo_letivo: int = None
) -> dict:
    """
    Retorna um snapshot completo de uma turma em uma única chamada consolidada:
    alunos, configuração de AVs, lista de AVs, notas e aulas dadas.

    Args:
        cod_periodo_letivo: necessário para buscar alunos e frequência.
                            Se None, alunos e frequência são omitidos.

    Returns:
        {
          "alunos":    list de get_alunos(),
          "config":    dict de get_config_avaliacoes(),
          "avaliacoes": list de get_avaliacoes_parciais(),
          "notas":     list de get_notas_alunos(),
          "aulas":     list de get_aulas_dadas(),
        }
    """
    config    = get_config_avaliacoes(cod_classe, cod_periodo)
    cod_regra = config.get("regraCalculo", {}).get("codigo", 3)
    qtde_av   = config.get("qtdeAvaliacao", 2)

    resultado = {
        "config":     config,
        "avaliacoes": get_avaliacoes_parciais(cod_classe, cod_periodo, cod_regra, qtde_av),
        "notas":      get_notas_alunos(cod_classe, cod_periodo),
        "aulas":      get_aulas_dadas(cod_classe, cod_periodo),
    }

    if cod_periodo_letivo:
        resultado["alunos"] = get_alunos(cod_classe, cod_periodo, cod_periodo_letivo)

    return resultado
