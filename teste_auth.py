from rco.auth import conectar_chrome
from rco.escolas import get_escolas, get_escolas_turmas

browser = conectar_chrome()
print("Conectado:", browser.title)

print("\nEscolas:")
for e in get_escolas(browser):
    print(" -", e)

print("\nTurmas por escola:")
turmas = get_escolas_turmas(browser)
escola_atual = ""
for t in turmas:
    if t["escola"] != escola_atual:
        escola_atual = t["escola"]
        print(f"\n  {escola_atual}")
    print(f"    - {t['turma']} | {t['disciplina']}")