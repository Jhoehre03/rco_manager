import json
from rco.auth import conectar_chrome
from database import entrar_turma, carregar
from rco.notas import navegar_avaliacao, preencher_formulario_avaliacao, preencher_notas, obter_datas_aula

TURMA      = "3ª Série - Noite - A"
DISCIPLINA = "FISICA"
TRIMESTRE  = "1º Tri"

# ------------------------------------------------------------------
# 1. Conecta ao Chrome em modo debug
# ------------------------------------------------------------------
browser = conectar_chrome()
print("Conectado:", browser.title)

# ------------------------------------------------------------------
# 2. Descobre o nome da escola que contém a turma
# ------------------------------------------------------------------
dados = carregar()
nome_escola = None
alunos_turma = []

for escola in dados["escolas"]:
    for t in escola["turmas"]:
        if t["turma"] == TURMA and t["disciplina"] == DISCIPLINA:
            nome_escola  = escola["nome"]
            alunos_turma = t["alunos"]
            break
    if nome_escola:
        break

if not nome_escola:
    raise SystemExit(
        f"Turma '{TURMA}' / '{DISCIPLINA}' não encontrada no dados.json. "
        "Rode o teste_auth.py primeiro."
    )

print(f"Escola encontrada: {nome_escola} — {len(alunos_turma)} alunos")

# ------------------------------------------------------------------
# 3. Navega para a tela de turmas e entra na turma
# ------------------------------------------------------------------
print("\nNavegando para /livro...")
browser.get("https://rco.paas.pr.gov.br/livro")

print(f"Entrando em: {TURMA} | {DISCIPLINA} | {TRIMESTRE}")
ok = entrar_turma(browser, nome_escola, TURMA, DISCIPLINA, TRIMESTRE)
if not ok:
    raise SystemExit("Falha ao entrar na turma. Verifique se o RCO está na tela correta.")
print("OK — turma aberta")

# ------------------------------------------------------------------
# 4. Navega para Avaliação
# ------------------------------------------------------------------
navegar_avaliacao(browser)
print("OK — página de avaliação carregada")

# ------------------------------------------------------------------
# 5. Busca datas disponíveis e preenche o formulário
# ------------------------------------------------------------------
datas = obter_datas_aula(browser)
if not datas:
    raise SystemExit("Nenhuma data com aula encontrada no calendário.")

# Usa a mais recente (datas já vêm em ordem decrescente)
data_iso = datas[0]
data_br = f"{data_iso[8:10]}/{data_iso[5:7]}/{data_iso[0:4]}"
print(f"Datas disponíveis: {datas}")
print(f"Usando: {data_br}")

preencher_formulario_avaliacao(browser, tipo="AV1", data=data_br, valor="30")
print("OK — formulário enviado, tabela de alunos visível")

# ------------------------------------------------------------------
# 6. Monta notas de teste (25) para os primeiros 5 alunos
# ------------------------------------------------------------------
ativos = [a for a in alunos_turma if not a.get("situacao", "").strip()]
primeiros_cinco = ativos[:5]
notas_teste = [
    {"nome_normalizado": a["nome_normalizado"], "nota": "25"}
    for a in primeiros_cinco
]

print("\nPreenchendo notas para:")
for item in notas_teste:
    print(f"  {item['nome_normalizado']} → {item['nota']}")

preencher_notas(browser, notas_teste)

# ------------------------------------------------------------------
# 7. Concluído
# ------------------------------------------------------------------
print("\nPronto! Confira as notas no RCO e marque os conteúdos manualmente.")
