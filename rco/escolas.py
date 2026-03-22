from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_escolas_turmas(browser):
    """
    Retorna lista de escolas com suas turmas e disciplinas.
    """
    wait = WebDriverWait(browser, 10)

    wait.until(EC.presence_of_all_elements_located(
        (By.CLASS_NAME, "card-header")
    ))

    resultado = []

    # Cada card representa uma escola
    cards_escola = browser.find_elements(By.CSS_SELECTOR, "div.card")

    for card in cards_escola:
        # Pega o nome da escola no card-header direto
        try:
            header = card.find_element(By.CSS_SELECTOR, "div.card-header")
            nome_escola = header.text.strip()
        except:
            continue

        # Filtra entradas que não são escolas (ex: "2026-1")
        if not nome_escola or nome_escola[:4].replace("-", "").isdigit():
            continue

        # Pega as turmas dentro do card
        turmas_elementos = card.find_elements(By.CSS_SELECTOR, "div.p-1")

        for turma_el in turmas_elementos:
            divs = turma_el.find_elements(By.CLASS_NAME, "d-flex")
            textos = [d.text.strip() for d in divs if d.text.strip()]

            if len(textos) >= 2:
                resultado.append({
                    "escola": nome_escola,
                    "turma": textos[1] if len(textos) > 2 else textos[0],
                    "disciplina": textos[2] if len(textos) > 2 else textos[1],
                })

    return resultado


def get_escolas(browser):
    """
    Retorna só os nomes das escolas.
    """
    wait = WebDriverWait(browser, 10)

    wait.until(EC.presence_of_all_elements_located(
        (By.CLASS_NAME, "card-header")
    ))

    cards_escola = browser.find_elements(By.CSS_SELECTOR, "div.card")
    escolas = []

    for card in cards_escola:
        try:
            header = card.find_element(By.CSS_SELECTOR, "div.card-header")
            nome_escola = header.text.strip()
        except:
            continue

        if not nome_escola or nome_escola[:4].replace("-", "").isdigit():
            continue

        if nome_escola not in escolas:
            escolas.append(nome_escola)

    return escolas