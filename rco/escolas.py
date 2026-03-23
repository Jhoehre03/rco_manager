from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unicodedata
import time


def get_texto(browser, elemento):
    """Lê texto via JavaScript — necessário para o RCO."""
    return browser.execute_script(
        "return arguments[0].textContent", elemento
    ).strip()


def get_escolas_turmas(browser):
    wait = WebDriverWait(browser, 15)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "card-header")))
    time.sleep(5)

    resultado = []

    turmas_bodies = browser.find_elements(
        By.XPATH,
        "//div[contains(@class,'card-body') and .//div[contains(@class,'d-flex') and contains(@class,'font-weight-bold')]]"
    )

    for body in turmas_bodies:
        try:
            escola = get_texto(browser, body.find_element(
                By.XPATH,
                "./ancestor::div[contains(@class,'card')][last()]//div[@class='card-header'][1]"
            )).split("\n")[0].strip()
        except:
            escola = "Desconhecida"

        divs = body.find_elements(By.XPATH, "div")
        textos = [get_texto(browser, d) for d in divs]
        textos = [t for t in textos if t and "Tri" not in t]

        if len(textos) >= 3:
            resultado.append({
                "escola": escola,
                "nivel": textos[0],
                "turma": textos[1],
                "disciplina": textos[2],
            })

    return resultado


def get_escolas(browser):
    turmas = get_escolas_turmas(browser)
    escolas = []
    for t in turmas:
        if t["escola"] not in escolas:
            escolas.append(t["escola"])
    return escolas


def get_alunos(browser):
    wait = WebDriverWait(browser, 10)
    wait.until(EC.presence_of_element_located(
        (By.ID, "table-transition-alunos")
    ))

    linhas = browser.find_elements(
        By.CSS_SELECTOR, "#table-transition-alunos tbody tr"
    )

    alunos = []
    for linha in linhas:
        numero = linha.get_attribute("data-pk")
        try:
            nome = linha.find_element(
                By.CSS_SELECTOR, "div.text-nowrap"
            )
            nome_texto = linha.find_element(
                By.CSS_SELECTOR, "div.text-nowrap"
            ).get_attribute("textContent").strip()
        except:
            continue

        if nome_texto and numero:
            alunos.append({
                "numero": int(numero),
                "nome": nome_texto,
                "nome_normalizado": normalizar(nome_texto)
            })

    return alunos


def normalizar(nome):
    return unicodedata.normalize('NFKD', nome)\
           .encode('ASCII', 'ignore')\
           .decode('ASCII')\
           .upper()\
           .strip()