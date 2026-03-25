import os
import subprocess
import time
import webview
from database import carregar, atualizar_banco as _atualizar_banco
from rco.auth import conectar_chrome as _conectar_chrome

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class Api:
    def __init__(self):
        self.browser = None

    def get_turmas(self):
        dados = carregar()
        turmas = []
        for escola in dados.get("escolas", []):
            for t in escola["turmas"]:
                ativos = [a for a in t["alunos"] if not a.get("situacao", "").strip()]
                turmas.append({
                    "escola":      escola["nome"],
                    "turma":       t["turma"],
                    "disciplina":  t["disciplina"],
                    "total_alunos": len(t["alunos"]),
                    "ativos":      len(ativos),
                    "inativos":    len(t["alunos"]) - len(ativos),
                })
        return turmas

    def get_status(self):
        dados = carregar()
        return {
            "turmas_ativas":         sum(len(e["turmas"]) for e in dados.get("escolas", [])),
            "comentarios_pendentes": 0,
            "notas_pendentes":       0,
            "ultima_atualizacao":    dados.get("ultima_atualizacao", "Nunca"),
            "chrome_conectado":      self.browser is not None,
            "google_autorizado":     os.path.exists("token.json"),
        }

    def autenticar_google(self):
        try:
            from sheets.gerador import _get_creds
            _get_creds()   # abre o navegador e aguarda a autorização
            return {"ok": True, "autorizado": os.path.exists("token.json")}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def conectar_chrome(self):
        try:
            subprocess.Popen([
                "C:/Program Files/Google/Chrome/Application/chrome.exe",
                "--remote-debugging-port=9222",
                "--user-data-dir=C:/chrome_debug",
                "https://rco.paas.pr.gov.br",
            ])
            time.sleep(3)
            return {"ok": True, "aguardando_login": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def confirmar_login(self):
        try:
            self.browser = _conectar_chrome()
            return {"ok": True, "titulo": self.browser.title}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_comentarios_planilha(self, escola, turma, disciplina, data):
        """
        Lê a planilha da turma e retorna alunos com ocorrências relevantes.
        data: "DD/MM/AAAA"
        """
        try:
            from sheets.gerador import ler_ocorrencias_planilha
            dados = carregar()
            planilha_id = None
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            planilha_id = t.get("planilha_id")
                            break
            if not planilha_id:
                return {"ok": False, "erro": "Planilha não associada. Gere a planilha primeiro."}
            comentarios = ler_ocorrencias_planilha(planilha_id, data)
            return {"ok": True, "comentarios": comentarios}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def lancar_comentarios(self, escola, turma, disciplina, trimestre, data, comentarios):
        """
        Entra na turma, navega para frequência e lança comentários.
        comentarios: lista de {numero, nome, comentario}
        """
        if not self.browser:
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            from database import entrar_turma
            from rco.notas import lancar_comentarios_aula
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC

            self.browser.get("https://rco.paas.pr.gov.br/livro")
            WebDriverWait(self.browser, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.card"))
            )
            ok = entrar_turma(self.browser, escola, turma, disciplina, trimestre)
            if not ok:
                return {"ok": False, "erro": f"Não foi possível entrar na turma {turma}"}

            wait = WebDriverWait(self.browser, 15)
            link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@href,'/aula') and contains(.,'Frequência')]")
            ))
            self.browser.execute_script("arguments[0].click()", link)
            wait.until(EC.url_contains("/aula"))
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))

            lancar_comentarios_aula(self.browser, data, comentarios)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def atualizar_banco(self, trimestre):
        if not self.browser:
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            _atualizar_banco(self.browser, trimestre)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}


def iniciar():
    api = Api()
    webview.create_window(
        title="RCO Manager",
        url=os.path.join(BASE_DIR, "index.html"),
        js_api=api,
        width=1100,
        height=680,
        resizable=False,
    )
    webview.start()
