import hashlib
import hmac


DEFAULT_USERNAME = "Usuário"
DEFAULT_PASSWORD = "12345"

PEPPER = "dashboard-auditoria-local-2026-anna-leticia"

PASSWORD_SALT = "x_jZgnPty7PTKj1uaC4L7w"


def _normalize_username(username):
    return username.strip().casefold()


def _username_digest(username):
    return hmac.new(
        PEPPER.encode("utf-8"),
        _normalize_username(username).encode("utf-8"),
        hashlib.sha256
    ).hexdigest()


def _password_hash(password):
    return hashlib.scrypt(
        password.encode("utf-8"),
        salt=PASSWORD_SALT.encode("utf-8"),
        n=32768,
        r=8,
        p=1,
        dklen=64,
        maxmem=128 * 1024 * 1024
    )


def auth_is_configured():
    return True


def verify_credentials(username, password):
    username_ok = hmac.compare_digest(
        _username_digest(username),
        _username_digest(DEFAULT_USERNAME)
    )

    password_ok = hmac.compare_digest(
        _password_hash(password),
        _password_hash(DEFAULT_PASSWORD)
    )

    return username_ok and password_ok