import json
import os
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

ARQUIVO = "dados.json"


def carregar():
    if not os.path.exists(ARQUIVO):
        return {"escolas": [], "ultima_atualizacao": None}
    with open(ARQUIVO, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar(dados):
    dados["ultima_atualizacao"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    with open(ARQUIVO, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    print(f"Dados salvos em {ARQUIVO}")


def atualizar_banco(browser):
    from rco.escolas import get_escolas_turmas

    print("Navegando para a tela de turmas...")
    browser.get("https://rco.paas.pr.gov.br/livro")

    print("Buscando escolas e turmas...")
    turmas = get_escolas_turmas(browser)

    dados = carregar()
    dados["escolas"] = []

    escolas_dict = {}
    for t in turmas:
        escola = t["escola"]
        if escola not in escolas_dict:
            escolas_dict[escola] = []
        escolas_dict[escola].append({
            "turma": t["turma"],
            "disciplina": t["disciplina"],
            "alunos": []
        })

    for nome_escola, turmas_lista in escolas_dict.items():
        dados["escolas"].append({
            "nome": nome_escola,
            "turmas": turmas_lista
        })

    salvar(dados)
    print(f"Banco atualizado: {len(turmas)} turmas encontradas")
    return dados