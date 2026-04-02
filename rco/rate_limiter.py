import time
import threading


class RateLimiter:
    """
    Controla a taxa de requisições ao RCO:
    - Mínimo de 2 segundos entre requisições
    - Máximo de 20 requisições por minuto
    - Circuit breaker: bloqueia após 3 erros consecutivos
    """

    MAX_POR_MINUTO   = 60    # servidor suporta 800/min — usamos 60 para margem segura
    DELAY_MINIMO     = 0.5   # segundos entre requisições
    ERROS_CIRCUIT    = 3     # erros consecutivos para abrir o circuit breaker
    CIRCUIT_COOLDOWN = 30.0  # segundos antes de tentar novamente após circuit aberto

    def __init__(self):
        self._lock             = threading.Lock()
        self._ultimo           = 0.0
        self._janela_inicio    = time.monotonic()
        self._contagem_janela  = 0
        self._erros_consecutivos = 0
        self._circuit_aberto   = False
        self._circuit_desde    = 0.0

    def aguardar(self):
        """Bloqueia até que seja seguro fazer a próxima requisição."""
        with self._lock:
            agora = time.monotonic()

            # Circuit breaker
            if self._circuit_aberto:
                if agora - self._circuit_desde < self.CIRCUIT_COOLDOWN:
                    from rco.exceptions import ServidorIndisponivel
                    raise ServidorIndisponivel(
                        f"Circuit breaker aberto. Aguarde {self.CIRCUIT_COOLDOWN:.0f}s."
                    )
                # Cooldown expirou — fecha o circuit e tenta
                self._circuit_aberto = False
                self._erros_consecutivos = 0

            # Delay mínimo entre requisições
            decorrido = agora - self._ultimo
            if decorrido < self.DELAY_MINIMO:
                time.sleep(self.DELAY_MINIMO - decorrido)

            # Limite por minuto — reinicia janela se passou 60s
            agora = time.monotonic()
            if agora - self._janela_inicio >= 60.0:
                self._janela_inicio   = agora
                self._contagem_janela = 0

            if self._contagem_janela >= self.MAX_POR_MINUTO:
                espera = 60.0 - (agora - self._janela_inicio)
                if espera > 0:
                    time.sleep(espera)
                self._janela_inicio   = time.monotonic()
                self._contagem_janela = 0

            self._ultimo = time.monotonic()
            self._contagem_janela += 1

    def registrar_sucesso(self):
        with self._lock:
            self._erros_consecutivos = 0

    def registrar_erro(self):
        with self._lock:
            self._erros_consecutivos += 1
            if self._erros_consecutivos >= self.ERROS_CIRCUIT:
                self._circuit_aberto = True
                self._circuit_desde  = time.monotonic()

    def __repr__(self):
        return (
            f"RateLimiter(max_por_minuto={self.MAX_POR_MINUTO}, "
            f"delay_minimo={self.DELAY_MINIMO}s, "
            f"circuit_breaker={self.ERROS_CIRCUIT} erros, "
            f"circuit_aberto={self._circuit_aberto})"
        )
