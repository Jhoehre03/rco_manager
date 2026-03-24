import os
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
        }

    def conectar_chrome(self):
        try:
            self.browser = _conectar_chrome()
            return {"ok": True, "titulo": self.browser.title}
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
