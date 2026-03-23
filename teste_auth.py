from rco.auth import conectar_chrome
from database import atualizar_banco, carregar

browser = conectar_chrome()
print("Conectado:", browser.title)

print("\nAtualizando banco de dados...")
atualizar_banco(browser)

print("\nConteúdo do banco:")
dados = carregar()
print(f"Atualizado em: {dados['ultima_atualizacao']}")
for escola in dados["escolas"]:
    print(f"\n  {escola['nome']}")
    for turma in escola["turmas"]:
        print(f"    - {turma['turma']} | {turma['disciplina']}")