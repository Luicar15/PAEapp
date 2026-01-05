import hmac
import os
from functools import wraps
from flask import abort, request, session

AUTH_TOKEN = os.getenv("FORM_AUTH_TOKEN")


def _get_bearer_token():
    header = request.headers.get("Authorization", "")
    if header.lower().startswith("bearer "):
        return header[7:].strip()
    return None


def _is_authenticated():
    token = _get_bearer_token()
    if token and AUTH_TOKEN and hmac.compare_digest(token, AUTH_TOKEN):
        session["authenticated"] = True
        return True
    return session.get("authenticated", False)


def require_auth(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not AUTH_TOKEN:
            abort(500, description="La autenticación no está configurada.")
        if not _is_authenticated():
            abort(401, description="No autorizado: se requiere autenticación.")
        return func(*args, **kwargs)
    return wrapper
