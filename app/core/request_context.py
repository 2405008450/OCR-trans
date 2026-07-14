from contextvars import ContextVar, Token
from typing import Optional


_client_ip: ContextVar[Optional[str]] = ContextVar("client_ip", default=None)


def set_client_ip(value: Optional[str]) -> Token:
    normalized = str(value).strip()[:64] if value else None
    return _client_ip.set(normalized or None)


def reset_client_ip(token: Token) -> None:
    _client_ip.reset(token)


def get_client_ip() -> Optional[str]:
    return _client_ip.get()
