class TokenExpirado(Exception):
    """Token JWT expirado ou inválido (HTTP 401)."""


class RateLimitExcedido(Exception):
    """Servidor retornou HTTP 429 — muitas requisições."""


class ServidorIndisponivel(Exception):
    """Servidor inacessível após todas as tentativas."""


class RespostaInvalida(Exception):
    """Resposta inesperada do servidor (formato ou conteúdo inválido)."""
