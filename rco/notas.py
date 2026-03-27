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


def preencher_formulario_avaliacao(browser, tipo, data, valor, rec_de=None):
    """
    Preenche o formulário de avaliação e avança para a tabela de alunos.

    Args:
        tipo:   "AV1" ou "Recuperação"
        data:   string "DD/MM/AAAA"
        valor:  string sem vírgula, ex "30" para 3,0
        rec_de: quando tipo="Recuperação", texto da AV a marcar no checkbox,
                ex "AV1" — marca o checkbox cujo label contém esse texto
    """
    wait = WebDriverWait(browser, 15)
    _aguardar_sem_overlay(browser)

    # Seleciona o radio button: ATV N → "1" (AV1), REC N → "2" (Recuperação)
    radio_value = "2" if tipo.upper().startswith("REC") else "1"
    radio = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, f"input[type='radio'][value='{radio_value}']")
    ))
    browser.execute_script("arguments[0].click()", radio)
    time.sleep(0.5)

    # Se for Recuperação, marca o checkbox da AV correspondente
    if tipo.upper().startswith("REC") and rec_de:
        try:
            # Aguarda o grupo de checkboxes aparecer
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "[name='grupoRecuperadas']")
            ))
            # Encontra o label cujo texto contém rec_de (ex: "AV1")
            # O grupo tem role="group" e contém os checkboxes de recuperação
            grupo = browser.find_element(
                By.CSS_SELECTOR, "[data-vv-name='grupoRecuperadas']"
            )
            labels = grupo.find_elements(By.CSS_SELECTOR, "label.custom-control-label")
            for label in labels:
                if rec_de.upper() in label.text.upper():
                    browser.execute_script("arguments[0].click()", label)
                    time.sleep(0.3)
                    break
        except Exception as e:
            print(f"[WARN] Não foi possível marcar checkbox de recuperação: {e}")

    # Converte DD/MM/AAAA → YYYY-MM-DD e seleciona no calendário
    partes = data.split("/")
    data_iso = f"{partes[2]}-{partes[1]}-{partes[0]}"
    _abrir_calendario(browser)
    _selecionar_dia(browser, data_iso)

    # Preenche o valor no peso (não existe na tela de Recuperação)
    if not tipo.upper().startswith("REC"):
        campo_valor = wait.until(EC.presence_of_element_located((By.ID, "pesoDecimal")))
        browser.execute_script("arguments[0].value = arguments[1]", campo_valor, valor)
        browser.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}))", campo_valor
        )
        time.sleep(0.3)

    # Clica em Avançar (btn-primary no card-footer)
    time.sleep(0.3)
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

        browser.execute_script("arguments[0].value = arguments[1]", campo, str(nota).zfill(2))
        browser.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}))", campo
        )
        browser.execute_script(
            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}))", campo
        )


def abrir_frequencia_dia(browser, data):
    """
    Na lista de aulas da tela de frequência, clica no botão Alterar da linha
    correspondente à data informada.

    Args:
        data: string "DD/MM/AAAA"
    """
    wait = WebDriverWait(browser, 15)
    _aguardar_sem_overlay(browser)

    # Cada linha do tbody representa uma aula; a data fica na primeira célula.
    # O botão de editar é a[title="Alterar"] na mesma linha.
    btn = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        f"//tbody/tr[td[contains(.,'{data}')]]//a[@title='Alterar']"
    )))
    browser.execute_script("arguments[0].click()", btn)

    # Aguarda tela "Alterar Aula" carregar (tabela de alunos aparece)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr")))


def lancar_comentario_aluno(browser, numero, comentario):
    """
    Abre o modal de observações do aluno pelo número de chamada,
    preenche o comentário e confirma.

    Args:
        numero:     número de chamada do aluno (int ou string)
        comentario: texto a lançar
    """
    wait = WebDriverWait(browser, 15)

    # Encontra a linha pelo número de chamada (coluna span com o número)
    linha = wait.until(EC.presence_of_element_located((
        By.XPATH,
        f"//tbody/tr[td/span[normalize-space(text())='{numero}']]"
    )))

    # Clica no ícone de observações da linha
    link_obs = linha.find_element(By.CSS_SELECTOR, "a[title='Observações']")
    browser.execute_script("arguments[0].click()", link_obs)

    # Aguarda o modal abrir — id do conteúdo segue padrão obs-{id_aluno}___BV_modal_content_
    modal_content = wait.until(EC.visibility_of_element_located(
        (By.CSS_SELECTOR, "div[id^='obs-'][id$='___BV_modal_content_']")
    ))

    # Extrai o id_aluno do atributo id do modal
    modal_id = modal_content.get_attribute("id")
    id_aluno = modal_id.removeprefix("obs-").removesuffix("___BV_modal_content_")

    # Preenche o textarea com send_keys para triggerar a reatividade Vue
    textarea = wait.until(EC.element_to_be_clickable((By.ID, f"observacao-{id_aluno}-1")))
    textarea.click()
    textarea.clear()
    textarea.send_keys(comentario)
    time.sleep(0.5)

    # Aguarda o botão OK habilitar — footer é <footer>, não <div>
    btn_ok = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, f"[id='obs-{id_aluno}___BV_modal_footer_'] .btn-primary")
    ))
    browser.execute_script("arguments[0].click()", btn_ok)

    # Aguarda o modal fechar
    wait.until(EC.invisibility_of_element_located(
        (By.CSS_SELECTOR, f"[id='obs-{id_aluno}___BV_modal_content_']")
    ))


def salvar_frequencia(browser):
    """Clica no botão 'Alterar' fora de qualquer modal para salvar a frequência."""
    wait = WebDriverWait(browser, 15)

    btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH,
         "//button[contains(@class,'btn-primary') and "
         "(contains(.,'Alterar') or contains(.,'Salvar'))]"
         "[not(ancestor::div[contains(@class,'modal')])]")
    ))
    browser.execute_script("arguments[0].click()", btn)


def lancar_comentarios_aula(browser, data, comentarios):
    """
    Abre a frequência do dia e lança comentários para cada aluno.

    Args:
        data:        string "DD/MM/AAAA"
        comentarios: lista de dicts {numero, nome, comentario}
                     numero: número de chamada do aluno
    """
    abrir_frequencia_dia(browser, data)

    # Aguarda a tabela de alunos carregar completamente
    wait = WebDriverWait(browser, 15)
    wait.until(EC.presence_of_element_located(
        (By.XPATH, "//tbody/tr[td/span]")
    ))
    time.sleep(1)

    for item in comentarios:
        print(f"  Lançando comentário: {item['nome']}")
        lancar_comentario_aluno(browser, item["numero"], item["comentario"])

    salvar_frequencia(browser)
    print(f"Frequência do dia {data} salva.")


def debug_avaliacao(browser):
    """
    Navega para /avaliacao e imprime o HTML completo da página para
    mapear a estrutura antes de implementar buscar_notas_finais_rco.
    """
    wait = WebDriverWait(browser, 15)

    # Garante que está na página de avaliação
    if "/avaliacao" not in browser.current_url:
        link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@href,'/avaliacao') and contains(.,'Avalia')]")
        ))
        browser.execute_script("arguments[0].click()", link)
        wait.until(EC.url_contains("/avaliacao"))

    time.sleep(2)

    html = browser.page_source
    print("=" * 80)
    print("URL:", browser.current_url)
    print("=" * 80)
    print(html)
    print("=" * 80)

    # Também lista todos os links/abas visíveis na página
    print("\n--- Links e abas visíveis ---")
    els = browser.find_elements(By.CSS_SELECTOR, "a, button, [role='tab'], .nav-link")
    for el in els:
        txt  = browser.execute_script("return arguments[0].textContent", el).strip()
        href = el.get_attribute("href") or el.get_attribute("class") or ""
        if txt:
            print(f"  [{el.tag_name}] '{txt}' → {href[:80]}")

    # Lista todas as tabelas
    print("\n--- Tabelas encontradas ---")
    tabelas = browser.find_elements(By.CSS_SELECTOR, "table")
    for i, t in enumerate(tabelas):
        linhas = t.find_elements(By.CSS_SELECTOR, "tr")
        print(f"  Tabela {i}: {len(linhas)} linhas")
        for j, linha in enumerate(linhas[:5]):   # primeiras 5 linhas
            txt = browser.execute_script("return arguments[0].textContent", linha).strip()
            print(f"    Linha {j}: {txt[:120]}")


def buscar_notas_finais_rco(browser):
    """
    Na página /avaliacao, clica na aba 'Alunos' e lê a somatória de cada aluno.

    A tabela tem colunas: Nº | Nome | Situação | Somatória
    A "Somatória" é a nota acumulada do trimestre atual (ou '-' se não houver).

    Retorna lista de dicts:
        numero   (int)
        nome     (str)
        situacao (str)   — ex. "Transf", "" (ativo)
        soma     (float|None) — somatória do trimestre, None se '-'
    """
    wait = WebDriverWait(browser, 15)

    # Navega para /avaliacao se ainda não estiver lá
    if "/avaliacao" not in browser.current_url:
        link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@href,'/avaliacao') and contains(.,'Avalia')]")
        ))
        browser.execute_script("arguments[0].click()", link)
        wait.until(EC.url_contains("/avaliacao"))
        time.sleep(1.5)

    # Clica na aba "Alunos"
    aba = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[@role='tab' and contains(.,'Alunos')]")
    ))
    browser.execute_script("arguments[0].click()", aba)
    time.sleep(1.5)

    # Aguarda a tabela dentro do painel ativo
    wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, ".tab-pane.active tbody tr")
    ))

    def _parse_soma(texto):
        texto = texto.strip().replace(",", ".")
        if texto in ("-", ""):
            return None
        try:
            return float(texto)
        except ValueError:
            return None

    # Detecta índice da coluna "Somatória" pelos cabeçalhos (dinâmico)
    ths = browser.find_elements(By.CSS_SELECTOR, ".tab-pane.active thead th")
    cabecalhos = [
        browser.execute_script("return arguments[0].textContent", th).strip().lower()
        for th in ths
    ]
    idx_soma = next(
        (i for i, h in enumerate(cabecalhos) if "somat" in h),
        len(cabecalhos) - 1   # fallback: última coluna
    )

    linhas = browser.find_elements(By.CSS_SELECTOR, ".tab-pane.active tbody tr")
    resultado = []
    for linha in linhas:
        cells = linha.find_elements(By.CSS_SELECTOR, "td")
        if len(cells) < 3:
            continue

        def _txt(idx):
            if idx >= len(cells):
                return ""
            return browser.execute_script(
                "return arguments[0].textContent", cells[idx]
            ).strip()

        numero_s = _txt(0)
        if not numero_s.isdigit():
            continue

        resultado.append({
            "numero":   int(numero_s),
            "nome":     _txt(1),
            "situacao": _txt(2),
            "soma":     _parse_soma(_txt(idx_soma)),
        })

    return resultado
