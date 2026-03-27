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


def get_config():
    """Retorna o dict 'config' armazenado em dados.json ou {} se não existir."""
    dados = carregar()
    return dados.get("config", {})


def salvar_config(config):
    """Persiste o dict 'config' em dados.json sem alterar ultima_atualizacao."""
    dados = carregar()
    dados["config"] = config
    with open(ARQUIVO, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


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


def marcar_nota_lancada(escola, turma, disciplina, trimestre, avaliacao):
    """
    Registra que uma avaliação foi lançada no RCO.
    trimestre: int 1/2/3
    avaliacao: ex "ATV 1", "REC 1"
    """
    dados = carregar()
    tri_key = f"{trimestre}T"
    hoje = datetime.now().strftime("%d/%m/%Y")

    for e in dados.get("escolas", []):
        if e["nome"] == escola:
            for t in e["turmas"]:
                if t["turma"] == turma and t["disciplina"] == disciplina:
                    if "notas_lancadas" not in t:
                        t["notas_lancadas"] = {}
                    if tri_key not in t["notas_lancadas"]:
                        t["notas_lancadas"][tri_key] = {}
                    t["notas_lancadas"][tri_key][avaliacao] = hoje
                    with open(ARQUIVO, "w", encoding="utf-8") as f:
                        json.dump(dados, f, ensure_ascii=False, indent=2)
                    return


def get_notas_lancadas(escola, turma, disciplina):
    """Retorna o dict notas_lancadas da turma ou {} se não existir."""
    dados = carregar()
    for e in dados.get("escolas", []):
        if e["nome"] == escola:
            for t in e["turmas"]:
                if t["turma"] == turma and t["disciplina"] == disciplina:
                    return t.get("notas_lancadas", {})
    return {}


def sincronizar_notas_lancadas(escola, turma, disciplina, trimestre, avaliacoes):
    """
    Substitui notas_lancadas[triKey] com os dados vindos do RCO.
    avaliacoes: lista de {nome, data}  — ex. [{"nome":"ATV 1","data":"15/03/2026"}]
    trimestre: 1, 2 ou 3 (int)
    """
    dados = carregar()
    tri_key = f"{trimestre}T"
    for e in dados.get("escolas", []):
        if e["nome"] == escola:
            for t in e["turmas"]:
                if t["turma"] == turma and t["disciplina"] == disciplina:
                    if "notas_lancadas" not in t:
                        t["notas_lancadas"] = {}
                    t["notas_lancadas"][tri_key] = {
                        a["nome"]: a["data"] for a in avaliacoes
                    }
                    with open(ARQUIVO, "w", encoding="utf-8") as f:
                        json.dump(dados, f, ensure_ascii=False, indent=2)
                    return


def marcar_comentario_lancado(escola, turma, disciplina, data):
    """
    Registra que os comentários de uma data foram lançados no RCO.
    data: "DD/MM/AAAA"
    """
    dados = carregar()
    for e in dados.get("escolas", []):
        if e["nome"] == escola:
            for t in e["turmas"]:
                if t["turma"] == turma and t["disciplina"] == disciplina:
                    if "comentarios_lancados" not in t:
                        t["comentarios_lancados"] = []
                    if data not in t["comentarios_lancados"]:
                        t["comentarios_lancados"].append(data)
                    with open(ARQUIVO, "w", encoding="utf-8") as f:
                        json.dump(dados, f, ensure_ascii=False, indent=2)
                    return


def get_comentarios_lancados(escola, turma, disciplina):
    """Retorna lista de datas já lançadas para a turma ou [] se não existir."""
    dados = carregar()
    for e in dados.get("escolas", []):
        if e["nome"] == escola:
            for t in e["turmas"]:
                if t["turma"] == turma and t["disciplina"] == disciplina:
                    return t.get("comentarios_lancados", [])
    return []


def atualizar_banco(browser, trimestre="1º Tri"):
    return atualizar_banco_progresso(browser, trimestre)


def atualizar_banco_progresso(browser, trimestre="1º Tri", on_progresso=None):
    """
    Atualiza o banco de dados buscando turmas e alunos do RCO.
    on_progresso(i, total, turma, disciplina, ok) é chamado após cada turma processada.
    Retorna dados salvos.
    """
    from rco.escolas import get_escolas_turmas, get_alunos

    print("Navegando para a tela de turmas...")
    browser.get("https://rco.paas.pr.gov.br/livro")

    print("Buscando escolas e turmas...")
    turmas = get_escolas_turmas(browser)

    # Preserva dados existentes (notas_lancadas, planilha_id etc.)
    dados = carregar()
    dados_existentes = {}
    for e in dados.get("escolas", []):
        for t in e["turmas"]:
            chave = (e["nome"], t["turma"], t["disciplina"])
            dados_existentes[chave] = t

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

    # Conta total de turmas para progresso
    total = sum(len(tl) for tl in escolas_dict.values())
    i = 0

    for nome_escola, turmas_lista in escolas_dict.items():
        print(f"\n  Escola: {nome_escola}")
        for turma in turmas_lista:
            i += 1
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
                alunos = get_alunos(browser)
                turma["alunos"] = alunos
                print(f"      {len(alunos)} alunos encontrados")
            else:
                print(f"      Não foi possível entrar na turma")

            # Preserva campos extras da turma existente
            chave_existente = (nome_escola, turma["turma"], turma["disciplina"])
            if chave_existente in dados_existentes:
                existente = dados_existentes[chave_existente]
                for campo in ("planilha_id", "notas_lancadas"):
                    if campo in existente:
                        turma[campo] = existente[campo]

            if on_progresso:
                on_progresso(i, total, turma["turma"], turma["disciplina"], entrou)

    for nome_escola, turmas_lista in escolas_dict.items():
        dados["escolas"].append({
            "nome": nome_escola,
            "turmas": turmas_lista
        })

    salvar(dados)
    print(f"\nBanco atualizado: {len(turmas)} turmas, trimestre {trimestre}")
    return dados