from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import unicodedata
import time


def _normalizar(nome):
    return (
        unicodedata.normalize("NFKD", nome)
        .encode("ASCII", "ignore")
        .decode("ASCII")
        .upper()
        .strip()
    )


def _aguardar_sem_overlay(browser, timeout=15):
    wait = WebDriverWait(browser, timeout)
    try:
        wait.until(EC.invisibility_of_element_located(
            (By.CSS_SELECTOR, "div.position-absolute.bg-light")
        ))
    except Exception:
        pass


def _abrir_calendario(browser):
    """Abre o datepicker e aguarda o calendário aparecer."""
    wait = WebDriverWait(browser, 15)
    btn = wait.until(EC.element_to_be_clickable((By.ID, "dataAvaliacaoParcial")))
    browser.execute_script("arguments[0].scrollIntoView({block:'center'})", btn)
    browser.execute_script("arguments[0].click()", btn)
    wait.until(lambda d: d.find_element(By.ID, "dataAvaliacaoParcial")
                          .get_attribute("aria-expanded") == "true")
    time.sleep(0.3)


def _fechar_calendario(browser):
    """Fecha o datepicker clicando fora dele."""
    try:
        fora = browser.find_element(By.CSS_SELECTOR, ".card-header, h4, h3, h2")
        browser.execute_script("arguments[0].click()", fora)
        time.sleep(0.3)
    except Exception:
        pass


def _navegar_mes(browser, mes_alvo):
    """
    Navega o calendário até o mês alvo.
    mes_alvo: string 'YYYY-MM'
    """
    for _ in range(24):
        grid = browser.find_element(By.CSS_SELECTOR, "[data-month]")
        mes_atual = grid.get_attribute("data-month")
        if mes_atual == mes_alvo:
            return
        if mes_atual > mes_alvo:
            btns = browser.find_elements(
                By.CSS_SELECTOR, "button[title='Previous month']"
            )
            if btns and "disabled" not in (btns[0].get_attribute("class") or ""):
                browser.execute_script("arguments[0].click()", btns[0])
                time.sleep(0.35)
            else:
                break
        else:
            btns = browser.find_elements(
                By.CSS_SELECTOR, "button[title='Next month']"
            )
            if btns and "disabled" not in (btns[0].get_attribute("class") or ""):
                browser.execute_script("arguments[0].click()", btns[0])
                time.sleep(0.35)
            else:
                break


def _selecionar_dia(browser, data_iso):
    """
    Navega até o mês certo e clica no dia.
    data_iso: 'YYYY-MM-DD'
    Levanta ValueError se o dia estiver desabilitado.
    """
    wait = WebDriverWait(browser, 10)
    mes_alvo = data_iso[:7]  # 'YYYY-MM'
    _navegar_mes(browser, mes_alvo)

    cell = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, f"div[data-date='{data_iso}']")
    ))
    if cell.get_attribute("aria-disabled") == "true":
        raise ValueError(
            f"Data {data_iso} não disponível no calendário — sem aula neste dia."
        )
    # Clique nativo via ActionChains para triggerar o handler Vue
    ActionChains(browser).move_to_element(cell).click().perform()

    # Aguarda a label do datepicker mostrar a data selecionada (não mais o placeholder)
    wait.until(lambda d: d.execute_script(
        "return document.getElementById('dataAvaliacaoParcial__value_').textContent"
    ) != "Clique aqui para selecionar a data")


def obter_datas_aula(browser):
    """
    Lê o calendário do formulário de avaliação e retorna todos os dias
    disponíveis (sem aria-disabled), navegando para trás até o limite.

    Deve ser chamado quando a página /avaliacao já estiver carregada.
    Retorna lista de strings 'YYYY-MM-DD', ordem decrescente (mais recente primeiro).
    """
    _aguardar_sem_overlay(browser)
    _abrir_calendario(browser)

    datas = []
    meses_visitados = set()

    for _ in range(12):
        grid = browser.find_element(By.CSS_SELECTOR, "[data-month]")
        mes_atual = grid.get_attribute("data-month")
        if mes_atual in meses_visitados:
            break
        meses_visitados.add(mes_atual)

        cells = browser.find_elements(By.CSS_SELECTOR, "div[data-date]")
        for cell in cells:
            if cell.get_attribute("aria-disabled") != "true":
                datas.append(cell.get_attribute("data-date"))

        btns = browser.find_elements(By.CSS_SELECTOR, "button[title='Previous month']")
        if not btns or "disabled" in (btns[0].get_attribute("class") or ""):
            break
        browser.execute_script("arguments[0].click()", btns[0])
        time.sleep(0.4)

    _fechar_calendario(browser)
    return sorted(set(datas), reverse=True)


def navegar_avaliacao(browser):
    """Clica no link 'Avaliação' no menu lateral e aguarda a página carregar."""
    wait = WebDriverWait(browser, 15)

    link = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(@href,'/avaliacao') and contains(.,'Avalia')]")
    ))
    browser.execute_script("arguments[0].click()", link)

    wait.until(EC.url_contains("/avaliacao"))
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='radio']")))


def preencher_formulario_avaliacao(browser, tipo, data, valor):
    """
    Preenche o formulário de avaliação e avança para a tabela de alunos.

    Args:
        tipo:  "AV1" ou "Recuperação"
        data:  string "DD/MM/AAAA"
        valor: string sem vírgula, ex "30" para 3,0
    """
    wait = WebDriverWait(browser, 15)
    _aguardar_sem_overlay(browser)

    # Seleciona o radio button
    radio_value = "1" if tipo == "AV1" else "2"
    radio = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, f"input[type='radio'][value='{radio_value}']")
    ))
    browser.execute_script("arguments[0].click()", radio)
    time.sleep(0.5)

    # Converte DD/MM/AAAA → YYYY-MM-DD e seleciona no calendário
    partes = data.split("/")
    data_iso = f"{partes[2]}-{partes[1]}-{partes[0]}"
    _abrir_calendario(browser)
    _selecionar_dia(browser, data_iso)

    # Preenche o valor no peso
    campo_valor = wait.until(EC.presence_of_element_located((By.ID, "pesoDecimal")))
    browser.execute_script("arguments[0].value = arguments[1]", campo_valor, valor)
    browser.execute_script(
        "arguments[0].dispatchEvent(new Event('input', {bubbles: true}))", campo_valor
    )
    time.sleep(0.3)

    # Clica em Avançar (btn-primary no card-footer)
    btn_avancar = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, ".card-footer .btn-primary")
    ))
    browser.execute_script("arguments[0].click()", btn_avancar)

    # Aguarda tabela de alunos aparecer
    wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "input[id^='notaDecimal-']")
    ))


def preencher_notas(browser, notas):
    """
    Preenche as notas na tabela de alunos.

    Args:
        notas: lista de dicts com {nome_normalizado, nota}
              nome_normalizado: nome sem acento, maiúsculo
              nota: string com o valor a digitar
    """
    indice = {item["nome_normalizado"]: item["nota"] for item in notas}

    linhas = browser.find_elements(By.CSS_SELECTOR, "tbody tr")

    for linha in linhas:
        try:
            nome_el = linha.find_element(By.CSS_SELECTOR, "td.text-truncate")
            nome_raw = browser.execute_script(
                "return arguments[0].textContent", nome_el
            ).strip()
        except Exception:
            continue

        nome_norm = _normalizar(nome_raw)
        nota = indice.get(nome_norm)
        if nota is None:
            continue

        try:
            campo = linha.find_element(By.CSS_SELECTOR, "input[id^='notaDecimal-']")
        except Exception:
            continue

        browser.execute_script("arguments[0].value = arguments[1]", campo, str(nota))
        browser.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}))", campo
        )
        browser.execute_script(
            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}))", campo
        )
