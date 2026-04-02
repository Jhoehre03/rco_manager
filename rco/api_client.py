import json
import os
import time
import requests
from urllib.parse import urlparse, parse_qs

from rco.exceptions import TokenExpirado, RateLimitExcedido, ServidorIndisponivel
from rco.rate_limiter import RateLimiter

_TOKEN_PATH = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "RCOManager", "rco_token.json"
)


class RCOClient:
    """
    Cliente HTTP seguro para o RCO.
    Gerencia autenticação, rate limiting, retries e audit log.
    """

    BASE_URL    = "https://apigateway-educacao.paas.pr.gov.br/seed/rcdig/estadual/v1"
    MAX_RETRIES = 3
    TIMEOUT     = 10

    def __init__(self):
        self.session      = requests.Session()
        self.token        = None
        self.rate_limiter = RateLimiter()
        self._audit_log   = []

    # ── Headers ──────────────────────────────────────────────────────────────

    def _headers_padrao(self):
        headers = {
            "User-Agent":      (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/146.0.0.0 Safari/537.36"
            ),
            "Accept":          "application/json, text/plain, */*",
            "Accept-Language": "pt-BR,pt;q=0.9",
            "Content-Type":    "application/json",
            "Origin":          "https://rco.paas.pr.gov.br",
            "Referer":         "https://rco.paas.pr.gov.br/",
            "consumerId":      "RCDIGWEB",
        }
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        return headers

    # ── Requisição base ───────────────────────────────────────────────────────

    def _request(self, method, endpoint, **kwargs):
        url = self.BASE_URL + endpoint
        tentativa = 0

        while tentativa < self.MAX_RETRIES:
            self.rate_limiter.aguardar()

            headers = {**self._headers_padrao(), **kwargs.pop("headers", {})}
            inicio  = time.monotonic()

            try:
                resp = self.session.request(
                    method,
                    url,
                    headers=headers,
                    timeout=self.TIMEOUT,
                    **kwargs,
                )
                elapsed = time.monotonic() - inicio

                self._registrar_audit(method, endpoint, resp.status_code, elapsed)

                if resp.status_code == 401:
                    # 401 é erro de autenticação, não de infraestrutura — não abre circuit breaker
                    raise TokenExpirado(f"Token expirado ou inválido — {endpoint}")

                if resp.status_code == 429:
                    self.rate_limiter.registrar_erro()
                    raise RateLimitExcedido(f"Rate limit excedido — {endpoint}")

                self.rate_limiter.registrar_sucesso()
                return resp

            except (TokenExpirado, RateLimitExcedido):
                raise

            except requests.exceptions.ConnectionError as e:
                self.rate_limiter.registrar_erro()
                tentativa += 1
                if tentativa >= self.MAX_RETRIES:
                    raise ServidorIndisponivel(
                        f"Servidor inacessível após {self.MAX_RETRIES} tentativas: {e}"
                    )
                time.sleep(2)

            except requests.exceptions.Timeout as e:
                self.rate_limiter.registrar_erro()
                tentativa += 1
                if tentativa >= self.MAX_RETRIES:
                    raise ServidorIndisponivel(
                        f"Timeout após {self.MAX_RETRIES} tentativas: {e}"
                    )
                time.sleep(2)

        raise ServidorIndisponivel("Máximo de tentativas atingido.")

    # ── Métodos públicos ──────────────────────────────────────────────────────

    def get(self, endpoint, params=None):
        resp = self._request("GET", endpoint, params=params)
        if "application/json" in resp.headers.get("Content-Type", ""):
            return resp.json()
        return resp.text

    def post(self, endpoint, data=None, json=None):
        resp = self._request("POST", endpoint, data=data, json=json)
        return resp.json()

    def set_token(self, token):
        self.token = token

    # ── Extração de token do browser ─────────────────────────────────────────

    def extrair_token_do_browser(self, browser):
        """
        Lê o token JWT do fragmento # da URL após o login.
        Tenta browser.current_url primeiro; fallback via JavaScript.
        Salva em AppData/Local/RCOManager/rco_token.json.
        Retorna True se encontrou o token, False caso contrário.
        """
        token      = None
        expires_in = None

        # Tentativa 1: URL atual do browser
        try:
            url       = browser.current_url
            fragmento = urlparse(url).fragment
            params    = parse_qs(fragmento)
            token     = params.get("access_token", [None])[0]
            expires_in = params.get("expires_in", [None])[0]
        except Exception as e:
            print(f"[TOKEN] Erro ao ler current_url: {e}")

        # Tentativa 2: via JavaScript
        if not token:
            try:
                hash_js   = browser.execute_script("return window.location.hash")
                fragmento = hash_js.lstrip("#")
                params    = parse_qs(fragmento)
                token     = params.get("access_token", [None])[0]
                expires_in = params.get("expires_in", [None])[0]
            except Exception as e:
                print(f"[TOKEN] Erro ao ler hash via JS: {e}")

        if not token:
            print("[TOKEN] Token não encontrado na URL.")
            return False

        self.set_token(token)
        print(f"[TOKEN] Token extraído com sucesso. expires_in={expires_in}s")

        # Persiste em disco
        try:
            os.makedirs(os.path.dirname(_TOKEN_PATH), exist_ok=True)
            with open(_TOKEN_PATH, "w", encoding="utf-8") as f:
                json.dump({
                    "access_token": token,
                    "expires_in":   expires_in,
                    "salvo_em":     time.strftime("%Y-%m-%d %H:%M:%S"),
                }, f, ensure_ascii=False, indent=2)
            print(f"[TOKEN] Salvo em {_TOKEN_PATH}")
        except Exception as e:
            print(f"[TOKEN] Erro ao salvar token: {e}")

        self._registrar_audit("TOKEN", "/auth/extrair", 200, 0)
        return True

    def carregar_token_salvo(self):
        """
        Carrega o token salvo em disco, se existir.
        Retorna True se carregou, False se não encontrou.
        """
        try:
            with open(_TOKEN_PATH, "r", encoding="utf-8") as f:
                dados = json.load(f)
            token = dados.get("access_token")
            if token:
                self.set_token(token)
                print(f"[TOKEN] Token carregado do disco (salvo em {dados.get('salvo_em')})")
                return True
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"[TOKEN] Erro ao carregar token salvo: {e}")
        return False

    # ── Audit log ─────────────────────────────────────────────────────────────

    def _registrar_audit(self, method, endpoint, status_code, elapsed):
        entrada = {
            "timestamp":   time.strftime("%Y-%m-%d %H:%M:%S"),
            "method":      method,
            "endpoint":    endpoint,
            "status_code": status_code,
            "elapsed_ms":  round(elapsed * 1000),
        }
        self._audit_log.append(entrada)
        try:
            print(
                f"[RCO] {entrada['timestamp']} {method} {endpoint} "
                f"-> {status_code} ({entrada['elapsed_ms']}ms)"
            )
        except Exception:
            pass

    def get_audit_log(self):
        return list(self._audit_log)
