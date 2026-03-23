import json
from datetime import date
from sheets.gerador import gerar_diario

PASTA_ID = "1MsRODhlMhWxRkKPni5jAJlJnqi5TOLJr"

# Carrega a primeira turma do banco
with open("dados.json", encoding="utf-8") as f:
    dados = json.load(f)

primeira_escola = dados["escolas"][0]
primeira_turma  = primeira_escola["turmas"][0]

turma_data = {
    "escola":     primeira_escola["nome"],
    "turma":      primeira_turma["turma"],
    "disciplina": primeira_turma["disciplina"],
    "alunos":     primeira_turma["alunos"],
}

config = {
    "num_aulas":           10,
    "data_inicio":         date(2026, 2, 6),
    "frequencia_semanal":  2,
    "avaliacoes": [
        {"valor_maximo": 3.0},
        {"valor_maximo": 3.0},
        {"valor_maximo": 4.0},
    ],
}

print(f"Escola:     {turma_data['escola']}")
print(f"Turma:      {turma_data['turma']}")
print(f"Disciplina: {turma_data['disciplina']}")
print(f"Alunos:     {len(turma_data['alunos'])}")
print()

url = gerar_diario(turma_data, config, PASTA_ID)
print(f"\nAcesse: {url}")
