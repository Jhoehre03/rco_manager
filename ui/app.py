import os
import json
import subprocess
import time
import webview
from database import carregar, atualizar_banco as _atualizar_banco
from rco.auth import conectar_chrome as _conectar_chrome

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
PASTA_ID  = "1MsRODhlMhWxRkKPni5jAJlJnqi5TOLJr"
DADOS_JSON = "dados.json"


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
                    "escola":          escola["nome"],
                    "turma":           t["turma"],
                    "disciplina":      t["disciplina"],
                    "total_alunos":    len(t["alunos"]),
                    "ativos":          len(ativos),
                    "inativos":        len(t["alunos"]) - len(ativos),
                    "planilha_id":     t.get("planilha_id", ""),
                    "notas_lancadas":  t.get("notas_lancadas", {}),
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

    def get_comentarios_periodo(self, data_inicio, data_fim):
        """
        Lê as planilhas de todas as turmas e retorna ocorrências no período.
        data_inicio, data_fim: "DD/MM/AAAA"
        Retorna: {ok, grupos: [{data, turma, disciplina, escola, trimestre, alunos: [...]}]}
        """
        try:
            from sheets.gerador import get_ocorrencias_periodo
            dados = carregar()

            # Coleta todas as turmas com planilha_id
            turmas_com_planilha = []
            for e in dados.get("escolas", []):
                for t in e["turmas"]:
                    pid = t.get("planilha_id", "")
                    if pid:
                        turmas_com_planilha.append({
                            "escola":      e["nome"],
                            "turma":       t["turma"],
                            "disciplina":  t["disciplina"],
                            "planilha_id": pid,
                        })

            if not turmas_com_planilha:
                return {"ok": False, "erro": "Nenhuma turma possui planilha associada."}

            # Agrupa resultados por (data, escola, turma, disciplina)
            from collections import defaultdict
            grupos_idx = defaultdict(list)

            for turma_info in turmas_com_planilha:
                try:
                    ocorrs = get_ocorrencias_periodo(
                        turma_info["planilha_id"], data_inicio, data_fim
                    )
                    for o in ocorrs:
                        chave = (o["data"], turma_info["escola"],
                                 turma_info["turma"], turma_info["disciplina"],
                                 o["trimestre"])
                        grupos_idx[chave].append({
                            "numero":     o["numero_chamada"],
                            "nome":       o["nome"],
                            "ocorrencia": o["ocorrencia"],
                            "comentario": o["comentario"],
                        })
                except Exception:
                    continue  # planilha inacessível — ignora

            grupos = [
                {
                    "data":        chave[0],
                    "escola":      chave[1],
                    "turma":       chave[2],
                    "disciplina":  chave[3],
                    "trimestre":   chave[4],
                    "alunos":      sorted(alunos, key=lambda a: a["numero"] or 0),
                }
                for chave, alunos in sorted(grupos_idx.items())
            ]

            return {"ok": True, "grupos": grupos}
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

    def gerar_planilha(self, escola, turma, disciplina, config):
        """
        Gera a planilha Google Sheets e salva o planilha_id no dados.json.
        config: {num_aulas, data_inicio ("DD/MM/AAAA"), frequencia_semanal,
                 avaliacoes: [{nome, valor_maximo}]}
        """
        try:
            from sheets.gerador import gerar_diario

            dados = carregar()
            turma_data = None
            turma_obj  = None
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            turma_data = {
                                "escola":     escola,
                                "turma":      turma,
                                "disciplina": disciplina,
                                "alunos":     t["alunos"],
                            }
                            turma_obj = t
                            break

            if not turma_data:
                return {"ok": False, "erro": "Turma não encontrada no dados.json"}

            config_gerador = {
                "modo":               config.get("modo", "diario"),
                "frequencia_semanal": int(config.get("frequencia_semanal", 2)),
                "avaliacoes": [
                    {
                        "nome":             a.get("nome", f"AV{i+1}"),
                        "valor_maximo":     float(a["valor_maximo"]),
                        "semana":           int(a["semana"]),
                        "peso_engajamento": float(a.get("peso_engajamento", 0.0)),
                        "peso_avaliacao":   float(a.get("peso_avaliacao",
                                                        float(a["valor_maximo"]))),
                    }
                    for i, a in enumerate(config["avaliacoes"])
                ],
            }

            resultado   = gerar_diario(turma_data, config_gerador, PASTA_ID)
            planilha_id = resultado["id"]
            url         = resultado["url"]

            # Persiste o planilha_id sem alterar ultima_atualizacao
            turma_obj["planilha_id"] = planilha_id
            with open(DADOS_JSON, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)

            return {"ok": True, "link": url, "id": planilha_id}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def preview_sincronizacao(self, escola, turma, disciplina):
        try:
            from sheets.gerador import comparar_alunos
            dados = carregar()
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            pid = t.get("planilha_id", "")
                            if not pid:
                                return {"ok": False, "erro": "Planilha não cadastrada"}
                            diff = comparar_alunos(pid, t["alunos"])
                            return {"ok": True, **diff}
            return {"ok": False, "erro": "Turma não encontrada"}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def sincronizar_alunos(self, escola, turma, disciplina):
        try:
            from sheets.gerador import (comparar_alunos, adicionar_aluno,
                                        atualizar_situacao, ocultar_aluno)
            dados = carregar()
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            pid = t.get("planilha_id", "")
                            if not pid:
                                return {"ok": False, "erro": "Planilha não cadastrada"}

                            diff = comparar_alunos(pid, t["alunos"])
                            n_novos = n_atualizados = n_ocultados = 0

                            for a in diff["novos"]:
                                adicionar_aluno(pid, a)
                                n_novos += 1

                            for a in diff["alterados"]:
                                atualizar_situacao(pid, a["numero"], a["situacao_nova"])
                                if a["acao"] == "ocultar":
                                    ocultar_aluno(pid, a["numero"])
                                    n_ocultados += 1
                                else:
                                    n_atualizados += 1

                            for a in diff["removidos"]:
                                ocultar_aluno(pid, a["numero"])
                                n_ocultados += 1

                            return {"ok": True, "novos": n_novos,
                                    "atualizados": n_atualizados,
                                    "ocultados": n_ocultados}
            return {"ok": False, "erro": "Turma não encontrada"}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def desvincular_planilha(self, escola, turma, disciplina):
        try:
            dados = carregar()
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            pid = t.get("planilha_id", "")
                            if not pid:
                                return {"ok": False, "erro": "Planilha não cadastrada"}
                            t["planilha_url_anterior"] = (
                                f"https://docs.google.com/spreadsheets/d/{pid}"
                            )
                            t["planilha_id"] = ""
                            with open(DADOS_JSON, "w", encoding="utf-8") as f:
                                json.dump(dados, f, ensure_ascii=False, indent=2)
                            return {"ok": True}
            return {"ok": False, "erro": "Turma não encontrada"}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def abrir_planilha(self, escola, turma, disciplina):
        dados = carregar()
        for e in dados.get("escolas", []):
            if e["nome"] == escola:
                for t in e["turmas"]:
                    if t["turma"] == turma and t["disciplina"] == disciplina:
                        pid = t.get("planilha_id", "")
                        if pid:
                            link = f"https://docs.google.com/spreadsheets/d/{pid}"
                            subprocess.Popen(["start", link], shell=True)
                            return {"ok": True, "link": link}
        return {"ok": False, "erro": "Planilha não cadastrada para esta turma"}

    def calcular_num_aulas(self, trimestre, freq_semanal):
        """
        Calcula o número aproximado de aulas no trimestre.
        trimestre:    1, 2 ou 3
        freq_semanal: número de aulas por semana
        Retorna: {ok, num_semanas, num_aulas, info}
        """
        try:
            from calendario.calendario_pr import calcular_semanas_trimestre, obter_info_trimestre
            num_semanas = calcular_semanas_trimestre(int(trimestre))
            freq        = int(freq_semanal) if freq_semanal else 0
            num_aulas   = num_semanas * freq
            info        = obter_info_trimestre(int(trimestre))
            return {"ok": True, "num_semanas": num_semanas, "num_aulas": num_aulas, "info": info}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def gerar_planilhas_em_lote(self, turmas, config):
        """
        Gera planilhas para múltiplas turmas com a mesma configuração.
        turmas: lista de {escola, turma, disciplina}
        config: mesmo formato de gerar_planilha
        Retorna: {ok, resultados: [{turma, disciplina, ok, link?, erro?}]}
        """
        try:
            from sheets.gerador import gerar_diario
            dados = carregar()

            # Monta índice rápido de turmas
            turmas_idx = {}
            for e in dados.get("escolas", []):
                for t in e["turmas"]:
                    turmas_idx[(e["nome"], t["turma"], t["disciplina"])] = (e, t)

            config_gerador = {
                "modo":               config.get("modo", "diario"),
                "frequencia_semanal": int(config.get("frequencia_semanal", 2)),
                "avaliacoes": [
                    {
                        "nome":             a.get("nome", f"AV{i+1}"),
                        "valor_maximo":     float(a["valor_maximo"]),
                        "semana":           int(a["semana"]),
                        "peso_engajamento": float(a.get("peso_engajamento", 0.0)),
                        "peso_avaliacao":   float(a.get("peso_avaliacao",
                                                        float(a["valor_maximo"]))),
                    }
                    for i, a in enumerate(config["avaliacoes"])
                ],
            }

            resultados = []
            dados_modificados = False

            for item in turmas:
                escola_nome = item["escola"]
                turma_nome  = item["turma"]
                disciplina  = item["disciplina"]
                chave = (escola_nome, turma_nome, disciplina)

                if chave not in turmas_idx:
                    resultados.append({
                        "turma": turma_nome, "disciplina": disciplina,
                        "ok": False, "erro": "Turma não encontrada"
                    })
                    continue

                _, t_obj = turmas_idx[chave]
                turma_data = {
                    "escola":     escola_nome,
                    "turma":      turma_nome,
                    "disciplina": disciplina,
                    "alunos":     t_obj["alunos"],
                }

                try:
                    resultado    = gerar_diario(turma_data, config_gerador, PASTA_ID)
                    planilha_id  = resultado["id"]
                    url          = resultado["url"]
                    t_obj["planilha_id"] = planilha_id
                    dados_modificados = True
                    resultados.append({
                        "turma": turma_nome, "disciplina": disciplina,
                        "ok": True, "link": url
                    })
                except Exception as ex:
                    resultados.append({
                        "turma": turma_nome, "disciplina": disciplina,
                        "ok": False, "erro": str(ex)
                    })

            if dados_modificados:
                with open(DADOS_JSON, "w", encoding="utf-8") as f:
                    json.dump(dados, f, ensure_ascii=False, indent=2)

            return {"ok": True, "resultados": resultados}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_datas_aula(self, escola, turma, disciplina, trimestre):
        if not self.browser:
            return {"ok": False, "chrome": False,
                    "erro": "Chrome não conectado — conecte para ver as datas do RCO"}
        try:
            from database import entrar_turma
            from rco.notas import obter_datas_aula
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC

            tri_str = {1: "1º Tri", 2: "2º Tri", 3: "3º Tri"}.get(int(trimestre), str(trimestre))

            self.browser.get("https://rco.paas.pr.gov.br/livro")
            WebDriverWait(self.browser, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.card"))
            )
            ok = entrar_turma(self.browser, escola, turma, disciplina, tri_str)
            if not ok:
                return {"ok": False, "erro": f"Não foi possível entrar na turma {turma}"}

            from rco.notas import navegar_avaliacao
            navegar_avaliacao(self.browser)

            datas = obter_datas_aula(self.browser)
            # obter_datas_aula retorna "YYYY-MM-DD"; converte para "DD/MM/AAAA"
            datas_fmt = [
                f"{d[8:10]}/{d[5:7]}/{d[:4]}" for d in datas
            ]
            return {"ok": True, "datas": datas_fmt}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_avaliacoes_planilha(self, escola, turma, disciplina, trimestre):
        """
        Lê a linha 3 da aba do trimestre e retorna as avaliações disponíveis.
        Retorna: {ok, avaliacoes: [{av: "ATV 1", rec: "REC 1"}, ...]}
        """
        try:
            from sheets.gerador import get_avaliacoes_planilha
            dados = carregar()
            planilha_id = None
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            planilha_id = t.get("planilha_id")
                            break
            if not planilha_id:
                return {"ok": False, "erro": "Planilha não associada."}
            avs = get_avaliacoes_planilha(planilha_id, trimestre)
            return {"ok": True, "avaliacoes": avs}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_notas_planilha(self, escola, turma, disciplina, trimestre, coluna_av):
        """
        Lê a planilha do Sheets e retorna as notas de uma avaliação.
        trimestre: 1, 2 ou 3
        coluna_av: "ATV 1", "REC 1", etc.
        Retorna: {ok, alunos: [...], ja_lancada: bool, data_lancamento: str|None}
        """
        try:
            from sheets.gerador import ler_notas_planilha
            from database import get_notas_lancadas
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
            alunos = ler_notas_planilha(planilha_id, trimestre, coluna_av)
            nl = get_notas_lancadas(escola, turma, disciplina)
            tri_key = f"{trimestre}T"
            data_lanc = nl.get(tri_key, {}).get(coluna_av)
            return {
                "ok": True,
                "alunos": alunos,
                "ja_lancada":      data_lanc is not None,
                "data_lancamento": data_lanc,
            }
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def lancar_notas(self, escola, turma, disciplina, trimestre, tipo_av, data, valor, notas):
        """
        Entra na turma, navega para avaliação e lança as notas.
        notas: lista de {nome_normalizado, nota} (nota já no formato RCO, ex: "23")
        """
        if not self.browser:
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            from database import entrar_turma
            from rco.notas import navegar_avaliacao, preencher_formulario_avaliacao, preencher_notas
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

            navegar_avaliacao(self.browser)
            # "REC N" → recuperação da "AV N"; qualquer outro → AV normal
            if tipo_av.upper().startswith("REC "):
                n      = tipo_av.split()[-1]   # "REC 4" → "4"
                rec_de = f"AV{n}"
                tipo_rco = "Recuperação"
            else:
                rec_de   = None
                tipo_rco = "AV1"
            preencher_formulario_avaliacao(self.browser, tipo_rco, data, str(valor), rec_de=rec_de)
            preencher_notas(self.browser, notas)

            from database import marcar_nota_lancada
            marcar_nota_lancada(escola, turma, disciplina, int(trimestre[0]), tipo_av)
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
