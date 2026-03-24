from rco.auth import conectar_chrome
from database import entrar_turma, carregar
from rco.notas import lancar_comentarios_aula
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

TURMA      = "3ª Série - Noite - A"
DISCIPLINA = "FISICA"
TRIMESTRE  = "1º Tri"
DATA_AULA  = "11/02/2026"

# ------------------------------------------------------------------
# 1. Conecta ao Chrome em modo debug
# ------------------------------------------------------------------
browser = conectar_chrome()
print("Conectado:", browser.title)

# ------------------------------------------------------------------
# 2. Descobre escola e alunos no dados.json
# ------------------------------------------------------------------
dados = carregar()
nome_escola  = None
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

ativos = [a for a in alunos_turma if not a.get("situacao", "").strip()]
print(f"Escola: {nome_escola} — {len(ativos)} alunos ativos")

# ------------------------------------------------------------------
# 3. Entra na turma
# ------------------------------------------------------------------
print("\nNavegando para /livro...")
browser.get("https://rco.paas.pr.gov.br/livro")

print(f"Entrando em: {TURMA} | {DISCIPLINA} | {TRIMESTRE}")
ok = entrar_turma(browser, nome_escola, TURMA, DISCIPLINA, TRIMESTRE)
if not ok:
    raise SystemExit("Falha ao entrar na turma.")
print("OK — turma aberta")

# ------------------------------------------------------------------
# 4. Navega para Frequência
# ------------------------------------------------------------------
wait = WebDriverWait(browser, 15)
link_freq = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//a[contains(@href,'/aula') and contains(.,'Frequência')]")
))
browser.execute_script("arguments[0].click()", link_freq)
wait.until(EC.url_contains("/aula"))
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
print("OK — tela de frequência carregada")

# ------------------------------------------------------------------
# 5. Lança comentários nos 2 primeiros alunos ativos
# ------------------------------------------------------------------
comentarios = [
    {
        "numero":     ativos[0]["numero"],
        "nome":       ativos[0]["nome"],
        "comentario": "Teste de comentário automático",
    },
    {
        "numero":     ativos[1]["numero"],
        "nome":       ativos[1]["nome"],
        "comentario": "Segundo teste de comentário",
    },
]

print(f"\nLançando comentários na aula {DATA_AULA}:")
lancar_comentarios_aula(browser, DATA_AULA, comentarios)

print("\nPronto! Confira os comentários no RCO.")
