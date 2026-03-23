import json
import os
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

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


def entrar_turma(browser, nome_escola, nome_turma, disciplina, trimestre):
    wait = WebDriverWait(browser, 15)

    try:
        wait.until(EC.invisibility_of_element_located(
            (By.CSS_SELECTOR, "div.position-absolute.bg-light")
        ))
    except Exception:
        pass

    time.sleep(2)

    # Primeiro acha o card da escola correta
    cards = browser.find_elements(By.CSS_SELECTOR, "div.card")
    card_escola = None

    for card in cards:
        headers = card.find_elements(By.CSS_SELECTOR, "div > div.card-header")
        if headers:
            texto_header = browser.execute_script(
                "return arguments[0].textContent", headers[0]
            ).strip().split("\n")[0].strip()
            if nome_escola in texto_header:
                card_escola = card
                break

    if not card_escola:
        print(f"      Escola não encontrada: {nome_escola}")
        return False

    # Dentro do card da escola, acha o body da turma correta
    bodies = card_escola.find_elements(
        By.XPATH,
        ".//div[contains(@class,'card-body') and .//div[contains(@class,'d-flex') and contains(@class,'font-weight-bold')]]"
    )

    for body in bodies:
        divs = body.find_elements(By.XPATH, "div")
        textos = [
            browser.execute_script("return arguments[0].textContent", d).strip()
            for d in divs
        ]
        textos = [t for t in textos if t and "Tri" not in t]

        # mesma estrutura que get_escolas_turmas: [nivel, turma, disciplina]
        if len(textos) < 3:
            continue

        if textos[1] == nome_turma and textos[2] == disciplina:
            botoes = body.find_elements(By.CSS_SELECTOR, "a.btn-outline-primary")
            if not botoes:
                continue

            for botao in botoes:
                texto_botao = browser.execute_script(
                    "return arguments[0].textContent", botao
                ).strip()
                if trimestre in texto_botao:
                    browser.execute_script("arguments[0].click()", botao)
                    try:
                        wait.until(EC.invisibility_of_element_located(
                            (By.CSS_SELECTOR, "div.position-absolute.bg-light")
                        ))
                    except Exception:
                        time.sleep(4)
                    return True

    print(f"      Botão do trimestre '{trimestre}' não encontrado em: {nome_turma} | {disciplina}")
    return False


def atualizar_banco(browser, trimestre="1º Tri"):
    from rco.escolas import get_escolas_turmas, get_alunos

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
        chave = (t["turma"], t["disciplina"])
        if not any((x["turma"], x["disciplina"]) == chave for x in escolas_dict[escola]):
            escolas_dict[escola].append({
                "turma": t["turma"],
                "disciplina": t["disciplina"],
                "alunos": []
            })

    for nome_escola, turmas_lista in escolas_dict.items():
        print(f"\n  Escola: {nome_escola}")
        for turma in turmas_lista:
            print(f"    Buscando: {turma['turma']} | {turma['disciplina']}...")

            browser.get("https://rco.paas.pr.gov.br/livro")
            try:
                WebDriverWait(browser, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.card"))
                )
            except Exception:
                time.sleep(5)

            entrou = entrar_turma(
                browser, nome_escola, turma["turma"], turma["disciplina"], trimestre
            )

            if entrou:
                url_atual = browser.current_url
                try:
                    titulo_el = browser.find_element(By.CSS_SELECTOR, "h4, h5, .card-header")
                    titulo = browser.execute_script(
                        "return arguments[0].textContent", titulo_el
                    ).strip().split("\n")[0].strip()
                except Exception:
                    titulo = "não encontrado"
                print(f"      URL: {url_atual}")
                print(f"      Título na página: {titulo}")
                alunos = get_alunos(browser)
                turma["alunos"] = alunos
                print(f"      {len(alunos)} alunos encontrados")
            else:
                print(f"      Não foi possível entrar na turma")

    for nome_escola, turmas_lista in escolas_dict.items():
        dados["escolas"].append({
            "nome": nome_escola,
            "turmas": turmas_lista
        })

    salvar(dados)
    print(f"\nBanco atualizado: {len(turmas)} turmas, trimestre {trimestre}")
    return dados