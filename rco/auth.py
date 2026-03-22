from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


def conectar_chrome():
    """
    Conecta ao Chrome já aberto em modo debug na porta 9222.
    """
    options = Options()
    options.debugger_address = "localhost:9222"
    browser = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    return browser


def fazer_login(cpf, senha):
    """
    Abre o Chrome normalmente e faz login no RCO.
    Usar só quando a sessão expirar.
    """
    options = Options()
    browser = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    url = 'https://auth-cs.identidadedigital.pr.gov.br/centralautenticacao/login.html?response_type=token&client_id=f340f1b1f65b6df5b5e3f94d95b11daf&redirect_uri=https%3A%2F%2Frco.paas.pr.gov.br&scope=emgpr.mobile%20emgpr.v1.ocorrencia.post&state=null&urlCert=https://certauth-cs.identidadedigital.pr.gov.br&dnsCidadao=https://cidadao-cs.identidadedigital.pr.gov.br/centralcidadao&loginPadrao=btnCentral&labelCentral=CPF,Login%20Sentinela&modulosDeAutenticacao=btnSentinela,btnSms,btnCpf,btnCentral&urlLogo=https%3A%2F%2Fwww.registrodeclasse.seed.pr.gov.br%2Frcdig%2Fimages%2Flogo_sistema.png&acesso=2100&tokenFormat=jwt&exibirLinkAutoCadastro=true&exibirLinkRecuperarSenha=true&exibirLinkAutoCadastroCertificado=false&captcha=false'

    browser.get(url)
    wait = WebDriverWait(browser, 10)

    wait.until(EC.element_to_be_clickable((By.ID, "btnCentral"))).click()
    wait.until(EC.presence_of_element_located((By.NAME, "attribute_central"))).send_keys(cpf)
    browser.find_element(By.NAME, "password").send_keys(senha)
    browser.find_element(By.ID, "btn-central-acessar").click()

    try:
        wait.until(EC.presence_of_element_located((By.ID, "mensagemServidor")))
        browser.quit()
        return None, "CPF ou senha inválidos"
    except:
        pass

    try:
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "pt-3")))
        botoes = browser.find_elements(By.CLASS_NAME, "pt-3")
        botoes[0].click()
        return browser, "ok"
    except:
        browser.quit()
        return None, "Erro ao acessar o sistema"