from rco.auth import conectar_chrome
from database import atualizar_banco, carregar

browser = conectar_chrome()
print("Conectado:", browser.title)

print("\nQual trimestre buscar?")
print("  1 - 1º Tri")
print("  2 - 2º Tri")
print("  3 - 3º Tri")

opcao = input("\nDigite 1, 2 ou 3: ").strip()
trimestres = {"1": "1º Tri", "2": "2º Tri", "3": "3º Tri"}
trimestre = trimestres.get(opcao, "1º Tri")

print("\nAtualizando banco de dados...")
atualizar_banco(browser, trimestre)

print("\nConteúdo do banco:")
dados = carregar()
print(f"Atualizado em: {dados['ultima_atualizacao']}")
for escola in dados["escolas"]:
    print(f"\n  {escola['nome']}")
    for turma in escola["turmas"]:
        print(f"    - {turma['turma']} | {turma['disciplina']} | {len(turma['alunos'])} alunos")