import os
import json
import subprocess
import time
import webview
from database import (carregar, atualizar_banco as _atualizar_banco,
                      atualizar_banco_progresso, marcar_comentario_lancado,
                      get_comentarios_lancados, sincronizar_notas_lancadas,
                      get_config, salvar_config,
                      get_planilhas_externas, cadastrar_planilha_externa,
                      remover_planilha_externa)
from rco.auth import conectar_chrome as _conectar_chrome

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
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

    def log(self, msg):
        print(f"[JS] {msg}", flush=True)

    def get_status(self):
        dados = carregar()

        # Verifica token da API
        api_conectada   = False
        token_expira_em = None
        try:
            from rco.api_client import _TOKEN_PATH
            import json as _json
            if os.path.exists(_TOKEN_PATH):
                with open(_TOKEN_PATH, "r", encoding="utf-8") as f:
                    t = _json.load(f)
                if t.get("access_token"):
                    api_conectada = True
                    token_expira_em = t.get("salvo_em", "")
        except Exception:
            pass

        return {
            "turmas_ativas":         sum(len(e["turmas"]) for e in dados.get("escolas", [])),
            "comentarios_pendentes": 0,
            "notas_pendentes":       0,
            "ultima_atualizacao":    dados.get("ultima_atualizacao", "Nunca"),
            "chrome_conectado":      self._browser_vivo(),
            "google_autorizado":     os.path.exists("token.json"),
            "api_conectada":         api_conectada,
            "token_salvo_em":        token_expira_em,
        }

    def get_snapshot_turma(self, escola, turma, disciplina):
        """
        Retorna snapshot via API REST: AVs configuradas, notas lançadas e aulas dadas.
        Requer token salvo em rco_token.json (capturado após login).
        """
        try:
            from rco.auth import rco_client
            from rco.consultas import get_snapshot_turma as _snapshot

            # Busca codClasse e codPeriodoAvaliacao do dados.json
            dados = carregar()
            cod_classe = cod_periodo = None
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            cod_classe  = t.get("cod_classe")
                            cod_periodo = t.get("cod_periodo_avaliacao")
                            break

            if not cod_classe or not cod_periodo:
                return {"ok": False, "erro": "codClasse não encontrado — atualize o banco primeiro"}

            if not rco_client.token:
                rco_client.carregar_token_salvo()
            if not rco_client.token:
                return {"ok": False, "erro": "Token da API não disponível — conecte o Chrome e faça login"}

            snapshot = _snapshot(cod_classe, cod_periodo)
            return {"ok": True, **snapshot}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_dados_via_api(self):
        """
        Sincroniza notas_lancadas de todas as turmas usando a API REST.
        Não usa Selenium — consulta direta à API com token salvo.
        Retorna resumo: {turmas_sincronizadas, avaliacoes_encontradas, erros}
        """
        import threading

        try:
            from rco.auth import rco_client
            from rco.consultas import get_config_avaliacoes, get_avaliacoes_parciais

            if not rco_client.token:
                rco_client.carregar_token_salvo()
            if not rco_client.token:
                return {"ok": False, "erro": "Token da API não disponível"}

            dados = carregar()
            turmas_sync  = 0
            avs_total    = 0
            erros        = []
            total        = sum(len(e["turmas"]) for e in dados.get("escolas", []))
            processadas  = 0

            for e in dados.get("escolas", []):
                for t in e["turmas"]:
                    processadas += 1
                    cod_classe  = t.get("cod_classe")
                    cod_periodo = t.get("cod_periodo_avaliacao")

                    pct = round((processadas / total) * 100) if total else 100
                    try:
                        webview.windows[0].evaluate_js(
                            f"window._onApiSyncProgresso && window._onApiSyncProgresso({pct},"
                            f"{json.dumps(t['turma'])},{json.dumps(t['disciplina'])})"
                        )
                    except Exception:
                        pass

                    if not cod_classe or not cod_periodo:
                        erros.append(f"{t['turma']}: sem codClasse")
                        continue

                    try:
                        config = get_config_avaliacoes(cod_classe, cod_periodo)
                        avs    = get_avaliacoes_parciais(
                            cod_classe, cod_periodo,
                            config.get("regraCalculo", {}).get("codigo", 3),
                            config.get("qtdeAvaliacao", 2),
                        )

                        # Monta notas_lancadas a partir das AVs que têm data definida
                        tri_key = f"{cod_periodo}T" if isinstance(cod_periodo, int) else "1T"
                        nl = t.setdefault("notas_lancadas", {}).setdefault(tri_key, {})
                        for av in avs:
                            nome  = av.get("descrAvaliacaoParcial", "").split("\n")[0].strip()
                            data  = av.get("dataAvaliacaoParcial", "")[:10]
                            if nome and data:
                                # Converte YYYY-MM-DD → DD/MM/YYYY
                                partes = data.split("-")
                                if len(partes) == 3:
                                    data = f"{partes[2]}/{partes[1]}/{partes[0]}"
                                nl[nome] = data
                                avs_total += 1

                        turmas_sync += 1
                    except Exception as ex:
                        erros.append(f"{t['turma']}: {ex}")

            with open(DADOS_JSON, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)

            return {
                "ok":                   True,
                "turmas_sincronizadas": turmas_sync,
                "avaliacoes_encontradas": avs_total,
                "erros":                erros,
            }
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def atualizar_tudo(self):
        """
        Atualiza tudo em sequência via API (sem Selenium):
        1. Captura token se não tiver
        2. Busca turmas do dia → salva cod_classe, cod_periodo, cod_periodo_letivo
        3. Para cada turma: busca alunos via API → atualiza dados.json
        4. Para cada turma: busca AVs lançadas → atualiza notas_lancadas
        5. Para cada turma com planilha: compara alunos com Sheets

        Envia progresso via JS: window._onAtualizarProgresso(pct, msg)
        Retorna: {ok, turmas, alunos_atualizados, avs_encontradas, erros}
        """
        import threading, datetime

        def _progresso(pct, msg):
            try:
                import json as _j
                webview.windows[0].evaluate_js(
                    f"window._onAtualizarProgresso && window._onAtualizarProgresso({pct},{_j.dumps(msg)})"
                )
            except Exception:
                pass

        try:
            from rco.auth import rco_client
            from rco.consultas import get_todas_turmas, get_alunos, get_config_avaliacoes, get_avaliacoes_parciais
            from rco.api_client import _TOKEN_PATH

            _progresso(5, "Verificando token...")

            # Captura token se não tiver
            if not rco_client.token:
                rco_client.carregar_token_salvo()
            if not rco_client.token and self._browser_vivo():
                self._capturar_token_e_codigos()
                import time; time.sleep(2)
                rco_client.carregar_token_salvo()
            if not rco_client.token:
                return {"ok": False, "erro": "Token não disponível — conecte o Chrome e faça login"}

            _progresso(10, "Buscando turmas (±7 dias)...")

            turmas_api = get_todas_turmas()

            dados = carregar()
            erros = []
            alunos_atualizados = 0
            avs_encontradas = 0
            turmas_processadas = 0

            # Índice: nome_turma → entry (usa o 1T como período corrente, mas salva todos)
            # get_todas_turmas retorna 3 entradas por turma (1T/2T/3T)
            turmas_api_por_classe = {}
            for ta in turmas_api:
                cod = ta["codClasse"]
                if cod not in turmas_api_por_classe:
                    turmas_api_por_classe[cod] = {"base": ta, "periodos": {}}
                tri = ta.get("trimestre", "")
                turmas_api_por_classe[cod]["periodos"][tri] = ta["codPeriodoAvaliacao"]

            # Atualiza cod_classe de cada turma no dados.json
            for ta in turmas_api:
                nome_api   = ta.get("turma", "").upper()
                escola_api = ta.get("escola", "").upper()
                for e in dados.get("escolas", []):
                    escola_json = e["nome"].upper()
                    # Escola deve bater (uma contém a outra) E turma deve bater
                    if escola_api not in escola_json and escola_json not in escola_api:
                        continue
                    for t in e["turmas"]:
                        if t["turma"] and t["turma"].upper() in nome_api:
                            entry = turmas_api_por_classe.get(ta["codClasse"], {})
                            t["cod_classe"]         = ta["codClasse"]
                            t["cod_periodo_letivo"] = ta["codPeriodoLetivo"]
                            t["periodos_avaliativos"] = entry.get("periodos", {})
                            t["cod_periodo_avaliacao"] = entry.get("periodos", {}).get(
                                "1º Trimestre", ta["codPeriodoAvaliacao"]
                            )

            total_ops = sum(len(e["turmas"]) for e in dados.get("escolas", []))
            op = 0

            for e in dados.get("escolas", []):
                for t in e["turmas"]:
                    op += 1
                    pct = 10 + int((op / total_ops) * 80)
                    turma_label = f"{t['turma']} — {t['disciplina']}"

                    cod_classe  = t.get("cod_classe")
                    cod_periodo = t.get("cod_periodo_avaliacao")
                    cod_pl      = t.get("cod_periodo_letivo")

                    if not cod_classe or not cod_periodo:
                        erros.append(f"{t['turma']}: sem codClasse")
                        continue

                    # 1. Alunos via API
                    try:
                        _progresso(pct, f"Alunos — {turma_label}")
                        alunos_api = get_alunos(cod_classe, cod_periodo, cod_pl) if cod_pl else []
                        if alunos_api:
                            t["alunos"] = [
                                {
                                    "numero":   a["numChamada"],
                                    "nome":     a["nome"],
                                    "situacao": a["situacao"],
                                    "nome_normalizado": __import__("unicodedata").normalize("NFKD", a["nome"])
                                        .encode("ASCII", "ignore").decode("ASCII").upper().strip(),
                                }
                                for a in alunos_api
                            ]
                            alunos_atualizados += len(alunos_api)
                    except Exception as ex:
                        erros.append(f"{t['turma']} alunos: {ex}")

                    # 2. AVs lançadas via API — salva com nomes da planilha (ATV N / REC N)
                    try:
                        _progresso(pct, f"AVs — {turma_label}")
                        config = get_config_avaliacoes(cod_classe, cod_periodo)
                        avs    = get_avaliacoes_parciais(
                            cod_classe, cod_periodo,
                            config.get("regraCalculo", {}).get("codigo", 3),
                            config.get("qtdeAvaliacao", 2),
                        )
                        tri_key = "1T"
                        nl_novo = {}
                        for av in avs:
                            tipo = av.get("codTipoAvaliacaoParcial")
                            num  = av.get("numAvaliacaoParcial")
                            data = (av.get("dataAvaliacaoParcial") or "")[:10]
                            if not data or not num:
                                continue
                            p = data.split("-")
                            if len(p) == 3:
                                data = f"{p[2]}/{p[1]}/{p[0]}"
                            if tipo == 1:
                                nl_novo[f"ATV {num}"] = data
                            elif tipo == 2:
                                nl_novo[f"REC {num}"] = data
                            avs_encontradas += 1
                        # Substitui completamente — evita acumular chaves de fontes antigas
                        t.setdefault("notas_lancadas", {})[tri_key] = nl_novo
                    except Exception as ex:
                        erros.append(f"{t['turma']} avs: {ex}")


            import time as _time
            dados["ultima_atualizacao"] = _time.strftime("%d/%m/%Y %H:%M")
            with open(DADOS_JSON, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)

            _progresso(100, "Concluído!")
            return {
                "ok":                True,
                "turmas":            total_ops,
                "alunos_atualizados": alunos_atualizados,
                "avs_encontradas":   avs_encontradas,
                "erros":             erros,
                "ultima_atualizacao": dados["ultima_atualizacao"],
            }
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_avaliacoes_rco_api(self, escola, turma, disciplina):
        """
        Retorna as AVs configuradas no RCO via API para uma turma específica.
        Usado no modal de Notas para preencher data e valor automaticamente.
        """
        try:
            from rco.auth import rco_client
            from rco.consultas import get_config_avaliacoes, get_avaliacoes_parciais

            if not rco_client.token:
                rco_client.carregar_token_salvo()
            if not rco_client.token:
                return {"ok": False, "erro": "Token não disponível"}

            dados = carregar()
            cod_classe = cod_periodo = None
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            cod_classe  = t.get("cod_classe")
                            cod_periodo = t.get("cod_periodo_avaliacao")
                            break

            if not cod_classe or not cod_periodo:
                return {"ok": False, "erro": "codClasse não encontrado"}

            config = get_config_avaliacoes(cod_classe, cod_periodo)
            avs    = get_avaliacoes_parciais(
                cod_classe, cod_periodo,
                config.get("regraCalculo", {}).get("codigo", 3),
                config.get("qtdeAvaliacao", 2),
            )

            # Mapeia por nomenclatura da planilha (ATV N / REC N) usando
            # codTipoAvaliacaoParcial (1=AV, 2=REC) e numAvaliacaoParcial
            resultado = []
            for av in avs:
                tipo = av.get("codTipoAvaliacaoParcial")
                num  = av.get("numAvaliacaoParcial")
                data = (av.get("dataAvaliacaoParcial") or "")[:10]
                peso = av.get("pesoDecimal", "")
                if not num:
                    continue
                if data:
                    p = data.split("-")
                    if len(p) == 3:
                        data = f"{p[2]}/{p[1]}/{p[0]}"
                nome_planilha = f"ATV {num}" if tipo == 1 else f"REC {num}"
                resultado.append({
                    "nome":            nome_planilha,
                    "data":            data,
                    "peso":            peso,
                    "tem_recuperacao": tipo == 1,
                })

            return {"ok": True, "avaliacoes": resultado}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_notas_rco_api(self, escola, turma, disciplina, trimestre, tipo_av):
        """
        Lê as notas de uma AV via API (sem Selenium).
        tipo_av: "ATV 1", "REC 1", etc.
        Retorna: {ok, alunos: [{numero, nome, nota}]}
        """
        try:
            from rco.auth import rco_client
            from rco.consultas import (get_config_avaliacoes,
                                       get_avaliacoes_parciais,
                                       get_notas_alunos)

            if not rco_client.token:
                rco_client.carregar_token_salvo()
            if not rco_client.token:
                return {"ok": False, "erro": "Token não disponível"}

            dados = carregar()
            cod_classe = cod_periodo = None
            periodos_avaliativos = {}
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            cod_classe           = t.get("cod_classe")
                            periodos_avaliativos = t.get("periodos_avaliativos", {})
                            break

            if not cod_classe:
                return {"ok": False, "erro": "codClasse não encontrado"}

            # Resolve cod_periodo pelo trimestre (ex: "1T" → "1º Trimestre")
            tri_num = int(str(trimestre).replace("T", "").strip())
            tri_key = f"{tri_num}º Trimestre"
            cod_periodo = periodos_avaliativos.get(tri_key)
            if not cod_periodo:
                return {"ok": False, "erro": f"Período avaliativo não encontrado para {tri_key}"}

            # Descobre o codAvaliacaoParcialClasse correspondente ao tipo_av
            config = get_config_avaliacoes(cod_classe, cod_periodo)
            avs    = get_avaliacoes_parciais(
                cod_classe, cod_periodo,
                config.get("regraCalculo", {}).get("codigo", 3),
                config.get("qtdeAvaliacao", 2),
            )

            # Mapeia nome_planilha → codAvaliacaoParcialClasse
            cod_av_alvo = None
            for av in avs:
                tipo = av.get("codTipoAvaliacaoParcial")
                num  = av.get("numAvaliacaoParcial")
                nome_planilha = f"ATV {num}" if tipo == 1 else f"REC {num}"
                if nome_planilha == tipo_av:
                    cod_av_alvo = str(av.get("codAvaliacaoParcialClasse"))
                    break

            if not cod_av_alvo:
                return {"ok": False, "erro": f"AV '{tipo_av}' não encontrada no RCO"}

            # Busca notas dos alunos
            notas_alunos = get_notas_alunos(cod_classe, cod_periodo)

            resultado = []
            for a in notas_alunos:
                nota_raw = a.get("notas", {}).get(cod_av_alvo)
                # Normaliza para o mesmo formato que a planilha usa no campo nota_rco:
                # valor decimal da API (ex "1.7") → inteiro ×10 (ex "17"), "-" se vazio
                if nota_raw is None or str(nota_raw).strip() in ("", "-"):
                    nota = "-"
                else:
                    try:
                        nota = str(round(float(str(nota_raw).replace(",", ".")) * 10))
                    except ValueError:
                        nota = str(nota_raw).strip()
                resultado.append({
                    "numero": a.get("numChamada"),
                    "nome":   a.get("nome", ""),
                    "nota":   nota,
                })

            return {"ok": True, "alunos": resultado}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_planilhas_externas(self):
        try:
            return {"ok": True, "planilhas": get_planilhas_externas()}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def cadastrar_planilha_externa(self, planilha_id, nome):
        try:
            cadastrar_planilha_externa(planilha_id, nome)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def remover_planilha_externa(self, planilha_id):
        try:
            remover_planilha_externa(planilha_id)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def diagnosticar_planilha(self, planilha_id):
        try:
            from sheets.gerador import diagnosticar_planilha
            resultado = diagnosticar_planilha(planilha_id)
            return {"ok": True, **resultado}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_configuracoes(self):
        try:
            return {"ok": True, **get_config()}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def salvar_configuracoes(self, config):
        try:
            salvar_config(config)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def verificar_atualizacao(self):
        try:
            from ui.updater import verificar_atualizacao
            return verificar_atualizacao()
        except Exception as e:
            return {"disponivel": False, "erro": str(e)}

    def instalar_atualizacao(self, url):
        import sys
        import threading

        frozen = getattr(sys, "frozen", False)

        def _progresso(pct):
            try:
                webview.windows[0].evaluate_js(
                    f"window._onUpdateProgresso && window._onUpdateProgresso({pct})"
                )
            except Exception:
                pass

        def _run():
            try:
                if not frozen:
                    webview.windows[0].evaluate_js(
                        "window._onUpdateProgresso && window._onUpdateProgresso(-1)"
                    )
                    return
                from ui.updater import baixar_e_instalar
                baixar_e_instalar(url, _progresso)
            except Exception as e:
                try:
                    msg = str(e).replace("'", "\\'")
                    webview.windows[0].evaluate_js(
                        f"window._onUpdateErro && window._onUpdateErro('{msg}')"
                    )
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()
        return {"ok": True}

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

    def _browser_vivo(self):
        """Retorna True se self.browser tem sessão ativa, False caso contrário."""
        if self.browser is None:
            return False
        try:
            _ = self.browser.title  # lança exceção se sessão morreu
            return True
        except Exception:
            self.browser = None
            return False

    def confirmar_login(self):
        try:
            # Descarta sessão anterior inválida antes de criar nova
            if self.browser is not None and not self._browser_vivo():
                self.browser = None
            self.browser = _conectar_chrome()

            # Captura token JWT do localStorage automaticamente
            self._capturar_token_e_codigos()

            return {"ok": True, "titulo": self.browser.title}
        except Exception as e:
            self.browser = None
            return {"ok": False, "erro": str(e)}

    def _capturar_token_e_codigos(self):
        """
        Após login, captura o token JWT do localStorage e salva codClasse,
        codPeriodoAvaliacao e codPeriodoLetivo de cada turma no dados.json.
        Executa silenciosamente — erros não interrompem o fluxo principal.
        """
        import threading
        def _run():
            try:
                from rco.auth import rco_client
                from rco.api_client import _TOKEN_PATH
                from rco.consultas import get_todas_turmas
                import datetime

                # Aguarda o SPA escrever o token no localStorage (até 15s)
                token_js = None
                for tentativa in range(15):
                    try:
                        token_js = self.browser.execute_script("""
                            var chaves = Object.keys(localStorage).concat(Object.keys(sessionStorage));
                            for (var i = 0; i < chaves.length; i++) {
                                var k = chaves[i];
                                var v = localStorage.getItem(k) || sessionStorage.getItem(k);
                                if (v && v.split('.').length === 3 && v.length > 100)
                                    return {chave: k, valor: v};
                            }
                            return null;
                        """)
                    except Exception:
                        token_js = None
                    if token_js:
                        break
                    time.sleep(1)

                if not token_js:
                    print("[TOKEN] Token não encontrado no localStorage após login.")
                    return

                token = token_js["valor"]
                rco_client.set_token(token)

                os.makedirs(os.path.dirname(_TOKEN_PATH), exist_ok=True)
                with open(_TOKEN_PATH, "w", encoding="utf-8") as f:
                    import json as _json
                    _json.dump({
                        "access_token": token,
                        "chave_storage": token_js["chave"],
                        "salvo_em": __import__("time").strftime("%Y-%m-%d %H:%M:%S"),
                    }, f, indent=2)
                print("[TOKEN] Token capturado e salvo automaticamente.")

                # Busca codClasse de cada turma via API (±7 dias para cobrir toda a grade)
                turmas_api = get_todas_turmas()
                if not turmas_api:
                    print("[API] Nenhuma turma retornada.")
                    return

                dados = carregar()
                atualizados = 0

                # Monta índice por codClasse com todos os períodos
                turmas_api_por_classe = {}
                for ta in turmas_api:
                    cod = ta["codClasse"]
                    if cod not in turmas_api_por_classe:
                        turmas_api_por_classe[cod] = {"base": ta, "periodos": {}}
                    tri = ta.get("trimestre", "")
                    turmas_api_por_classe[cod]["periodos"][tri] = ta["codPeriodoAvaliacao"]

                if not dados.get("escolas"):
                    # ── Primeiro uso: importa todas as turmas direto da API ──
                    import unicodedata as _ud
                    escolas_idx = {}
                    visitados = set()
                    for ta in turmas_api:
                        cod = ta["codClasse"]
                        if cod in visitados:
                            continue
                        visitados.add(cod)

                        escola_nome = ta["escola"]
                        entry = turmas_api_por_classe[cod]

                        # Extrai nome curto da turma (remove prefixo de modalidade)
                        nome_turma = ta["turma"]
                        partes = nome_turma.split(" - ")
                        turma_curta = " - ".join(partes[-3:]) if len(partes) >= 3 else nome_turma

                        nova_turma = {
                            "turma":               turma_curta,
                            "disciplina":          ta["disciplina"],
                            "alunos":              [],
                            "cod_classe":          cod,
                            "cod_periodo_letivo":  ta["codPeriodoLetivo"],
                            "periodos_avaliativos": entry["periodos"],
                            "cod_periodo_avaliacao": entry["periodos"].get(
                                "1º Trimestre", ta["codPeriodoAvaliacao"]
                            ),
                            "notas_lancadas":      {},
                        }

                        if escola_nome not in escolas_idx:
                            escolas_idx[escola_nome] = {"nome": escola_nome, "turmas": []}
                        escolas_idx[escola_nome]["turmas"].append(nova_turma)
                        atualizados += 1

                    dados["escolas"] = list(escolas_idx.values())
                    print(f"[API] Primeiro uso: {atualizados} turma(s) importadas da API.")

                else:
                    # ── Uso normal: atualiza cod_classe nas turmas existentes ──
                    visitados = set()
                    for ta in turmas_api:
                        nome_turma = ta.get("turma", "").upper()
                        cod = ta["codClasse"]
                        if cod in visitados:
                            continue
                        for e in dados.get("escolas", []):
                            for t in e["turmas"]:
                                if t["turma"] and t["turma"].upper() in nome_turma:
                                    entry = turmas_api_por_classe.get(cod, {})
                                    t["cod_classe"]            = cod
                                    t["cod_periodo_letivo"]    = ta["codPeriodoLetivo"]
                                    t["periodos_avaliativos"]  = entry.get("periodos", {})
                                    t["cod_periodo_avaliacao"] = entry.get("periodos", {}).get(
                                        "1º Trimestre", ta["codPeriodoAvaliacao"]
                                    )
                                    atualizados += 1
                                    visitados.add(cod)
                    print(f"[API] {atualizados} turma(s) atualizadas com codClasse.")

                if atualizados:
                    with open(DADOS_JSON, "w", encoding="utf-8") as f:
                        json.dump(dados, f, ensure_ascii=False, indent=2)

                    try:
                        webview.windows[0].evaluate_js(
                            f"window._onApiConectada && window._onApiConectada({atualizados})"
                        )
                    except Exception:
                        pass

            except Exception as e:
                print(f"[TOKEN] Erro ao capturar token/codigos: {e}")

        threading.Thread(target=_run, daemon=True).start()

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

            # Cache de datas lançadas por turma para evitar N leituras do JSON
            lancados_cache = {}
            def _lancados(escola, turma, disc):
                k = (escola, turma, disc)
                if k not in lancados_cache:
                    lancados_cache[k] = set(get_comentarios_lancados(escola, turma, disc))
                return lancados_cache[k]

            grupos = [
                {
                    "data":        chave[0],
                    "escola":      chave[1],
                    "turma":       chave[2],
                    "disciplina":  chave[3],
                    "trimestre":   chave[4],
                    "alunos":      sorted(alunos, key=lambda a: a["numero"] or 0),
                    "ja_lancado":  chave[0] in _lancados(chave[1], chave[2], chave[3]),
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
        if not self._browser_vivo():
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
            marcar_comentario_lancado(escola, turma, disciplina, data)
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

            pasta_id = get_config().get("pasta_drive_id", "")
            if not pasta_id:
                return {"ok": False, "erro": "Pasta do Drive não configurada. Acesse Configurações."}
            resultado   = gerar_diario(turma_data, config_gerador, pasta_id)
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

    def sincronizar_alunos_lote_stream(self, turmas):
        """
        Sincroniza alunos de múltiplas turmas com suas planilhas Google Sheets.
        Envia progresso via JS: _saProgressoEvento(i, total, turma, disciplina, status, detalhe)
        detalhe: {novos:[nomes], atualizados:[nomes], ocultados:[nomes], erro:str}
        turmas: lista de {escola, turma, disciplina}
        """
        try:
            from sheets.gerador import (comparar_alunos, adicionar_aluno,
                                        atualizar_situacao, ocultar_aluno)
            dados = carregar()

            # Índice rápido de turmas
            turmas_idx = {}
            for e in dados.get("escolas", []):
                for t in e["turmas"]:
                    turmas_idx[(e["nome"], t["turma"], t["disciplina"])] = t

            total = len(turmas)

            for i, item in enumerate(turmas, 1):
                if i > 1:
                    time.sleep(2)  # evita quota exceeded (60 req/min Sheets API)
                escola_nome = item["escola"]
                turma_nome  = item["turma"]
                disciplina  = item["disciplina"]
                chave = (escola_nome, turma_nome, disciplina)

                detalhe = {"novos": [], "atualizados": [], "ocultados": [], "erro": ""}
                status  = "ok"

                try:
                    t_obj = turmas_idx.get(chave)
                    if not t_obj:
                        raise Exception("Turma não encontrada no banco")
                    pid = t_obj.get("planilha_id", "")
                    if not pid:
                        raise Exception("Sem planilha vinculada")

                    # Busca alunos frescos via API se cod_classe disponível
                    cod_cls = t_obj.get("cod_classe")
                    cod_pa  = t_obj.get("cod_periodo_avaliacao")
                    cod_pl  = t_obj.get("cod_periodo_letivo")
                    if cod_cls and cod_pa and cod_pl:
                        from rco.consultas import get_alunos as _get_alunos_api
                        alunos_fonte = _get_alunos_api(cod_cls, cod_pa, cod_pl)
                        # Normaliza para o formato esperado por comparar_alunos
                        alunos_fonte = [
                            {"numero": a["numChamada"], "nome": a["nome"], "situacao": a["situacao"]}
                            for a in alunos_fonte
                        ]
                    else:
                        alunos_fonte = t_obj.get("alunos", [])

                    diff = comparar_alunos(pid, alunos_fonte)

                    for a in diff["novos"]:
                        adicionar_aluno(pid, a)
                        detalhe["novos"].append(a["nome"])

                    for a in diff["alterados"]:
                        atualizar_situacao(pid, a["numero"], a["situacao_nova"])
                        if a["acao"] == "ocultar":
                            ocultar_aluno(pid, a["numero"])
                            detalhe["ocultados"].append(a["nome"])
                        else:
                            detalhe["atualizados"].append(a["nome"])

                    for a in diff["removidos"]:
                        ocultar_aluno(pid, a["numero"])
                        detalhe["ocultados"].append(a["nome"])

                except Exception as ex:
                    detalhe["erro"] = str(ex)
                    status = "erro"

                js = (
                    f"_saProgressoEvento({i}, {total}, "
                    f"{json.dumps(turma_nome)}, {json.dumps(disciplina)}, "
                    f"{json.dumps(status)}, {json.dumps(detalhe)})"
                )
                try:
                    webview.windows[0].evaluate_js(js)
                except Exception:
                    pass

            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def vincular_planilha_externa(self, escola, turma, disciplina, planilha_id):
        """Vincula manualmente um planilha_id a uma turma, sem gerar nova planilha."""
        try:
            dados = carregar()
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            t["planilha_id"] = planilha_id
                            with open(DADOS_JSON, "w", encoding="utf-8") as f:
                                json.dump(dados, f, ensure_ascii=False, indent=2)
                            return {"ok": True}
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
            pasta_id = get_config().get("pasta_drive_id", "")
            if not pasta_id:
                return {"ok": False, "erro": "Pasta do Drive não configurada. Acesse Configurações."}
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
                    resultado    = gerar_diario(turma_data, config_gerador, pasta_id)
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

    def gerar_planilhas_em_lote_stream(self, turmas, config):
        """
        Gera planilhas em lote enviando progresso em tempo real via JS.
        Mesmo resultado de gerar_planilhas_em_lote mas com eventos por turma.
        """
        try:
            from sheets.gerador import gerar_diario
            pasta_id = get_config().get("pasta_drive_id", "")
            if not pasta_id:
                return {"ok": False, "erro": "Pasta do Drive não configurada. Acesse Configurações."}
            dados = carregar()

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

            total = len(turmas)
            resultados = []
            dados_modificados = False

            for i, item in enumerate(turmas, 1):
                escola_nome = item["escola"]
                turma_nome  = item["turma"]
                disciplina  = item["disciplina"]
                chave = (escola_nome, turma_nome, disciplina)

                if chave not in turmas_idx:
                    resultado_item = {
                        "turma": turma_nome, "disciplina": disciplina,
                        "ok": False, "erro": "Turma não encontrada"
                    }
                    resultados.append(resultado_item)
                else:
                    _, t_obj = turmas_idx[chave]
                    turma_data = {
                        "escola":     escola_nome,
                        "turma":      turma_nome,
                        "disciplina": disciplina,
                        "alunos":     t_obj["alunos"],
                    }
                    try:
                        resultado    = gerar_diario(turma_data, config_gerador, pasta_id)
                        planilha_id  = resultado["id"]
                        url          = resultado["url"]
                        t_obj["planilha_id"] = planilha_id
                        dados_modificados = True
                        resultado_item = {
                            "turma": turma_nome, "disciplina": disciplina,
                            "ok": True, "link": url
                        }
                    except Exception as ex:
                        resultado_item = {
                            "turma": turma_nome, "disciplina": disciplina,
                            "ok": False, "erro": str(ex)
                        }
                    resultados.append(resultado_item)

                # Envia progresso ao JS
                status_str = "ok" if resultado_item["ok"] else "erro"
                erro_str   = resultado_item.get("erro", "")
                js = (
                    f"_lbProgressoEvento({i}, {total}, "
                    f"{json.dumps(turma_nome)}, {json.dumps(disciplina)}, "
                    f"{json.dumps(status_str)}, {json.dumps(erro_str)})"
                )
                try:
                    webview.windows[0].evaluate_js(js)
                except Exception:
                    pass

            if dados_modificados:
                with open(DADOS_JSON, "w", encoding="utf-8") as f:
                    json.dump(dados, f, ensure_ascii=False, indent=2)

            ok_count   = sum(1 for r in resultados if r["ok"])
            fail_count = total - ok_count
            return {"ok": True, "resultados": resultados,
                    "ok_count": ok_count, "fail_count": fail_count}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_datas_aula(self, escola, turma, disciplina, trimestre):
        try:
            from rco.auth import rco_client
            from rco.consultas import get_datas_aula as _get_datas

            if not rco_client.token:
                rco_client.carregar_token_salvo()
            if not rco_client.token:
                return {"ok": False, "chrome": False,
                        "erro": "Token não disponível — conecte o Chrome e faça login"}

            dados = carregar()
            cod_classe = cod_periodo = None
            _tri_map = {"1º Tri": "1º Trimestre", "2º Tri": "2º Trimestre", "3º Tri": "3º Trimestre"}
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            cod_classe = t.get("cod_classe")
                            periodos   = t.get("periodos_avaliativos", {})
                            tri_int    = int(trimestre)
                            tri_str    = {1: "1º Trimestre", 2: "2º Trimestre", 3: "3º Trimestre"}.get(tri_int, "")
                            cod_periodo = periodos.get(tri_str) or t.get("cod_periodo_avaliacao")
                            break

            if not cod_classe or not cod_periodo:
                return {"ok": False, "erro": "Turma sem cod_classe — clique em Atualizar primeiro."}

            datas = _get_datas(cod_classe, cod_periodo)
            return {"ok": True, "datas": datas}
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
        if not self._browser_vivo():
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            from database import entrar_turma
            from rco.notas import (navegar_avaliacao, preencher_formulario_avaliacao,
                                   lancar_notas_completo)
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
            if tipo_av.upper().startswith("REC "):
                n        = tipo_av.split()[-1]
                rec_de   = f"AV{n}"
                tipo_rco = "Recuperação"
                modo     = "rec"
            else:
                rec_de   = None
                tipo_rco = "AV1"
                modo     = "novo_av"

            preencher_formulario_avaliacao(self.browser, tipo_rco, data, str(valor), rec_de=rec_de)
            lancar_notas_completo(self.browser, notas, modo)

            from database import marcar_nota_lancada
            marcar_nota_lancada(escola, turma, disciplina, int(trimestre[0]), tipo_av)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def get_notas_rco(self, escola, turma, disciplina, trimestre, tipo_av):
        """
        Lê as notas de uma AV já lançada diretamente do RCO (aba Alunos).
        Retorna: {ok, alunos: [{numero, nome, nota}]}
        nota: string como aparece no RCO, ex "30" (=3,0), "-" se vazio
        """
        if not self._browser_vivo():
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            from database import entrar_turma
            from rco.notas import navegar_avaliacao, buscar_notas_av_rco
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
            alunos = buscar_notas_av_rco(self.browser, tipo_av)
            return {"ok": True, "alunos": alunos}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def editar_notas(self, escola, turma, disciplina, trimestre, tipo_av, notas):
        """
        Edita notas de uma AV já lançada no RCO.
        Vai direto ao botão Alterar da AV na aba Avaliações — sem passar pelo formulário.
        notas: lista de {nome_normalizado, nota}
        """
        if not self._browser_vivo():
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            from database import entrar_turma
            from rco.notas import navegar_avaliacao, abrir_edicao_avaliacao, lancar_notas_completo
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
            abrir_edicao_avaliacao(self.browser, tipo_av)
            lancar_notas_completo(self.browser, notas, "editar_av")
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def atualizar_resumo(self, escola, turma, disciplina, trimestre):
        """
        Lê as notas finais do trimestre via API do RCO e atualiza a aba Resumo na planilha.
        trimestre: "1º Tri", "2º Tri" ou "3º Tri"
        """
        try:
            from rco.auth import rco_client
            from rco.consultas import get_notas_finais
            from sheets.gerador import adicionar_aba_resumo

            if not rco_client.token:
                rco_client.carregar_token_salvo()
            if not rco_client.token:
                return {"ok": False, "erro": "Token não disponível — conecte o Chrome e faça login"}

            # Mapeia trimestre string → chave do periodos_avaliativos
            _tri_map = {
                "1º Tri": "1º Trimestre",
                "2º Tri": "2º Trimestre",
                "3º Tri": "3º Trimestre",
            }
            tri_key = _tri_map.get(trimestre, "1º Trimestre")

            dados = carregar()
            planilha_id = None
            cod_classe = cod_periodo = None
            for e in dados.get("escolas", []):
                if e["nome"] == escola:
                    for t in e["turmas"]:
                        if t["turma"] == turma and t["disciplina"] == disciplina:
                            planilha_id = t.get("planilha_id")
                            cod_classe  = t.get("cod_classe")
                            periodos    = t.get("periodos_avaliativos", {})
                            cod_periodo = periodos.get(tri_key) or t.get("cod_periodo_avaliacao")
                            break

            if not planilha_id:
                return {"ok": False, "erro": "Planilha não associada. Gere a planilha primeiro."}
            if not cod_classe or not cod_periodo:
                return {"ok": False, "erro": "Turma sem cod_classe — clique em Atualizar primeiro."}

            notas_finais = get_notas_finais(cod_classe, cod_periodo)
            if not notas_finais:
                return {"ok": False, "erro": "Nenhuma nota encontrada para esta turma/trimestre."}

            resultado = adicionar_aba_resumo(planilha_id, notas_finais)
            return {"ok": True, **resultado}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def atualizar_banco(self, trimestre):
        if not self._browser_vivo():
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            _atualizar_banco(self.browser, trimestre)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def atualizar_banco_stream(self, trimestre):
        """Atualiza o banco enviando progresso em tempo real via JS."""
        if not self._browser_vivo():
            return {"ok": False, "erro": "Chrome não conectado"}

        def on_progresso(i, total, turma, disciplina, ok):
            status = "ok" if ok else "erro"
            js = (
                f"_abProgressoEvento({i}, {total}, "
                f"{json.dumps(turma)}, {json.dumps(disciplina)}, {json.dumps(status)})"
            )
            try:
                webview.windows[0].evaluate_js(js)
            except Exception:
                pass

        try:
            dados = atualizar_banco_progresso(self.browser, trimestre, on_progresso)
            ultima = dados.get("ultima_atualizacao", "")
            return {"ok": True, "ultima_atualizacao": ultima}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def sincronizar_notas_lote_stream(self, turmas, trimestre):
        """
        Para cada turma: entra no RCO, lê as AVs lançadas na aba 'Avaliações'
        e atualiza notas_lancadas no dados.json.
        Envia progresso via JS: _snProgressoEvento(i, total, turma, disciplina, status, info)
        turmas: lista de {escola, turma, disciplina}
        trimestre: 1, 2 ou 3 (int)
        """
        if not self._browser_vivo():
            return {"ok": False, "erro": "Chrome não conectado"}
        try:
            from rco.notas import buscar_avaliacoes_lancadas_rco
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from calendario.calendario_pr import TRIMESTRES
            from datetime import date

            tri_num  = int(trimestre)
            tri_str  = {1: "1º Tri", 2: "2º Tri", 3: "3º Tri"}.get(tri_num, "")
            ano      = date.today().year
            tri_info = TRIMESTRES.get(ano, {}).get(tri_num, {})
            fim_tri  = tri_info.get("fim")
            fechado  = fim_tri is not None and date.today() > fim_tri

            total = len(turmas)

            for i, item in enumerate(turmas, 1):
                escola_nome = item["escola"]
                turma_nome  = item["turma"]
                disciplina  = item["disciplina"]

                try:
                    self.browser.get("https://rco.paas.pr.gov.br/livro")
                    WebDriverWait(self.browser, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.card"))
                    )
                    from database import entrar_turma
                    ok = entrar_turma(self.browser, escola_nome, turma_nome,
                                     disciplina, tri_str)
                    if not ok:
                        raise Exception(f"Não foi possível entrar na turma")

                    avaliacoes = buscar_avaliacoes_lancadas_rco(self.browser)
                    sincronizar_notas_lancadas(escola_nome, turma_nome, disciplina,
                                              tri_num, avaliacoes)

                    info = f"{len(avaliacoes)} AV(s) encontrada(s)"
                    if fechado:
                        info += " · trimestre fechado"
                    status = "ok"
                except Exception as ex:
                    avaliacoes = []
                    info   = str(ex)
                    status = "erro"

                js = (
                    f"_snProgressoEvento({i}, {total}, "
                    f"{json.dumps(turma_nome)}, {json.dumps(disciplina)}, "
                    f"{json.dumps(status)}, {json.dumps(info)})"
                )
                try:
                    webview.windows[0].evaluate_js(js)
                except Exception:
                    pass

            return {"ok": True, "fechado": fechado}
        except Exception as e:
            return {"ok": False, "erro": str(e)}


def _garantir_webview2():
    """Verifica se o WebView2 Runtime está instalado; baixa e instala silenciosamente se não estiver."""
    import sys
    if not getattr(sys, "frozen", False):
        return  # Só verifica no executável empacotado

    try:
        import winreg
        winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients"
            r"\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
        )
        print("[WEBVIEW2] Já instalado.")
        return
    except FileNotFoundError:
        pass

    print("[WEBVIEW2] Não encontrado. Baixando instalador...")
    import urllib.request
    import subprocess
    import tempfile

    url_instalador = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
    tmp = tempfile.NamedTemporaryFile(suffix=".exe", delete=False)
    tmp.close()

    try:
        urllib.request.urlretrieve(url_instalador, tmp.name)
        print("[WEBVIEW2] Instalando silenciosamente...")
        subprocess.run(
            [tmp.name, "/silent", "/install"],
            check=True,
        )
        print("[WEBVIEW2] Instalação concluída.")
    except Exception as e:
        print(f"[WEBVIEW2] Erro na instalação: {e}")
    finally:
        try:
            os.remove(tmp.name)
        except Exception:
            pass


def iniciar():
    _garantir_webview2()
    api = Api()
    webview.create_window(
        title="RCO Manager",
        url=os.path.join(BASE_DIR, "index.html") + f"?v={int(time.time())}",
        js_api=api,
        width=1100,
        height=680,
        resizable=False,
    )

    def _verificar_update():
        import threading

        def _run():
            # Aguarda a janela carregar completamente antes de verificar
            time.sleep(5)
            print("[UPDATE] Verificando atualização...")
            from ui.updater import verificar_atualizacao
            resultado = verificar_atualizacao()
            print(f"[UPDATE] Resultado: {resultado}")

            if resultado.get("disponivel"):
                import json as _json
                payload = _json.dumps({
                    "versao":       resultado.get("versao", ""),
                    "url_download": resultado.get("url_download", ""),
                    "descricao":    resultado.get("descricao", ""),
                })
                js = f"window.notificarAtualizacao && window.notificarAtualizacao({payload})"
                print(f"[UPDATE] Chamando JS com payload JSON")
                try:
                    webview.windows[0].evaluate_js(js)
                except Exception as e:
                    print(f"[UPDATE] Erro evaluate_js: {e}")

        threading.Thread(target=_run, daemon=True).start()

    webview.start(func=_verificar_update, gui='edgechromium')

    # Encerra processos filhos do EdgeWebView2 que ficam órfãos após fechar a janela
    try:
        import subprocess
        subprocess.run(
            ["taskkill", "/F", "/IM", "msedgewebview2.exe"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
    except Exception:
        pass
