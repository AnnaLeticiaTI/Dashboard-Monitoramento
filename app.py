import os
import secrets
import threading
import time
import webbrowser
from datetime import timedelta
from urllib.parse import urlsplit

from dotenv import load_dotenv
from flask import (
    Flask,
    abort,
    jsonify,
    redirect,
    render_template,
    request,
    send_from_directory,
    session,
    url_for,
)

from config import BASE_DIR
from controllers.dashboard_controller import api
from services.carteira_service import (
    get_carteira_source_summary,
    load_carteira_data,
    load_aguardando_pa_data,
)
from services.data_service import get_information_summary, load_data
from services.paint_service import get_paint_source_summary, load_paint_data
from services.auth_service import verify_credentials


load_dotenv(BASE_DIR / ".env.local")
load_dotenv(BASE_DIR / ".env")

IS_VERCEL = os.environ.get("VERCEL") == "1"

app = Flask(__name__)

app.config.update(
    SECRET_KEY=os.environ.get(
        "DASHBOARD_SECRET_KEY",
        "dashboard-demo-secret-key-2026-anna-leticia-1234567890"
    ),
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=(
        IS_VERCEL
        or os.environ.get("DASHBOARD_SECURE_COOKIE") == "1"
    ),
    SESSION_COOKIE_NAME="dashboard_auditoria_session",
)

app.register_blueprint(api)


if not IS_VERCEL:

    @app.get("/css/<path:filename>")
    def local_css(filename):
        return send_from_directory(
            BASE_DIR / "public" / "css",
            filename
        )

    @app.get("/js/<path:filename>")
    def local_javascript(filename):
        return send_from_directory(
            BASE_DIR / "public" / "js",
            filename
        )


def _csrf_token() -> str:
    token = session.get("csrf_token")

    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token

    return token


def _valid_csrf(token: str) -> bool:
    expected = session.get("csrf_token", "")

    return bool(
        expected
        and token
        and secrets.compare_digest(expected, token)
    )


def _safe_next_url(target: str | None) -> str:
    if not target:
        return url_for("index")

    parsed = urlsplit(target)

    if (
        parsed.scheme
        or parsed.netloc
        or not target.startswith("/")
    ):
        return url_for("index")

    return target


@app.before_request
def require_authentication():

    if request.endpoint in {
        "login",
        "local_css",
        "local_javascript"
    } or request.method == "OPTIONS":
        return None

    if session.get("authenticated"):
        return None

    if request.path.startswith("/api/"):
        return jsonify({
            "error": "Autenticação necessária."
        }), 401

    return redirect(
        url_for(
            "login",
            next=request.full_path.rstrip("?")
        )
    )


@app.after_request
def secure_response(response):

    response.headers.setdefault(
        "X-Content-Type-Options",
        "nosniff"
    )

    response.headers.setdefault(
        "X-Frame-Options",
        "DENY"
    )

    response.headers.setdefault(
        "Referrer-Policy",
        "strict-origin-when-cross-origin"
    )

    response.headers.setdefault(
        "Permissions-Policy",
        "camera=(), microphone=(), geolocation=()"
    )

    if IS_VERCEL or request.is_secure:
        response.headers.setdefault(
            "Strict-Transport-Security",
            "max-age=31536000; includeSubDomains"
        )

    if request.endpoint != "static":
        response.headers.setdefault(
            "Cache-Control",
            "private, no-store, max-age=0"
        )

    return response


@app.context_processor
def inject_security_helpers():

    if not app.secret_key:
        return {
            "csrf_token": ""
        }

    return {
        "csrf_token": _csrf_token()
    }


@app.route("/login", methods=["GET", "POST"])
def login():

    error = None

    if session.get("authenticated"):
        return redirect(url_for("index"))

    if request.method == "POST":

        if not _valid_csrf(
            request.form.get("csrf_token", "")
        ):
            abort(400)

        now = int(time.time())

        locked_until = int(
            session.get("login_locked_until", 0)
        )

        if locked_until > now:

            error = (
                "Muitas tentativas. "
                "Aguarde um minuto e tente novamente."
            )

        elif verify_credentials(
            request.form.get("username", ""),
            request.form.get("password", "")
        ):

            destination = _safe_next_url(
                request.form.get("next")
            )

            session.clear()
            session.permanent = True
            session["authenticated"] = True
            session["csrf_token"] = secrets.token_urlsafe(32)

            return redirect(destination)

        else:

            attempts = int(
                session.get("login_attempts", 0)
            ) + 1

            if attempts >= 5:

                session["login_attempts"] = 0
                session["login_locked_until"] = now + 60

                error = (
                    "Muitas tentativas. "
                    "Aguarde um minuto e tente novamente."
                )

            else:

                session["login_attempts"] = attempts

                error = "Usuário ou senha inválidos."

    return render_template(
        "login.html",
        error=error,
        configuration_error=False,
        next_url=_safe_next_url(
            request.args.get("next")
            or request.form.get("next")
        ),
    )


@app.post("/logout")
def logout():

    if not _valid_csrf(
        request.form.get("csrf_token", "")
    ):
        abort(400)

    session.clear()

    return redirect(url_for("login"))


@app.route("/")
def index():
    return render_template(
        "index.html",
        active_page="inicio"
    )


@app.route("/monitoramento")
def monitoramento():
    return render_template(
        "monitoramento.html",
        active_page="monitoramento"
    )


@app.route("/carteira")
def carteira():
    return render_template(
        "carteira.html",
        active_page="carteira"
    )


@app.route("/semaforo")
def semaforo():
    return render_template(
        "semaforo.html",
        active_page="semaforo"
    )


@app.route("/execucao-paint")
def execucao_paint():
    return render_template(
        "execucao_paint.html",
        active_page="execucao_paint"
    )


@app.route("/informacoes")
def informacoes():

    monitoring_summary = get_information_summary(
        load_data()
    )

    paint_summary = get_paint_source_summary(
        load_paint_data()
    )

    carteira_summary = get_carteira_source_summary(
        load_carteira_data()
    )

    return render_template(
        "informacoes.html",
        active_page="informacoes",
        monitoring=monitoring_summary,
        paint=paint_summary,
        carteira=carteira_summary,
    )


if __name__ == "__main__":

    load_data()
    load_paint_data()
    load_carteira_data()
    load_aguardando_pa_data()

    dashboard_url = "http://127.0.0.1:5000"

    if os.environ.get("DASHBOARD_OPEN_BROWSER") == "1":
        threading.Timer(
            1.0,
            lambda: webbrowser.open(dashboard_url)
        ).start()

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=False,
        use_reloader=False
    )