"""
tracker_auth.py
───────────────────────────────────────────────────────────────
Secure router for the Anderson Lab Report Tracker.

Endpoints mounted at /tracker by backend.py
  GET  /tracker/login      → serve login page HTML
  POST /tracker/login      → validate credentials, set HttpOnly cookie
  GET  /tracker/           → serve tracker dashboard (auth required)
  GET  /tracker/xlsx.min.js → serve bundled SheetJS (auth required)
  POST /tracker/logout     → clear session cookie

Credentials are stored ONLY in .env:
  TRACKER_USER=your_username
  TRACKER_PASS_HASH=<bcrypt hash>   ← run create_credentials.py to generate
  TRACKER_SECRET=<random 64-char hex>   ← used to sign JWTs
"""

import os
import ssl
import time
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

import certifi

import bcrypt
from dotenv import load_dotenv
from fastapi import APIRouter, Cookie, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from jose import JWTError, jwt

load_dotenv()

# ── Config ────────────────────────────────────────────────────
TRACKER_USER      = os.getenv("TRACKER_USER", "")
TRACKER_PASS_HASH = os.getenv("TRACKER_PASS_HASH", "").encode()
TRACKER_SECRET    = os.getenv("TRACKER_SECRET", "")
TOKEN_EXPIRE_H    = int(os.getenv("TRACKER_SESSION_HOURS", "8"))
ALGORITHM         = "HS256"

# Fail fast if critical config is missing
if not TRACKER_USER or not TRACKER_PASS_HASH or not TRACKER_SECRET:
    raise RuntimeError(
        "Missing tracker credentials in .env — run create_credentials.py first. "
        "Required: TRACKER_USER, TRACKER_PASS_HASH, TRACKER_SECRET"
    )

FRONTEND_DIR = Path(__file__).parent.parent / "front end"
COOKIE_NAME  = "tracker_session"

# Google Sheet config
GSHEET_ID = os.getenv(
    "TRACKER_GSHEET_ID",
    "1WToMt7-X3CJ5J2SeBbm9gf9sF6oq6PabECQBhHgS-gY",
)
GSHEET_GID = os.getenv("TRACKER_GSHEET_GID", "1126267100")  # Sheet1 (3rd tab)

router = APIRouter(prefix="/tracker")

# ── Brute-force rate limiter (in-memory, per IP) ──────────────
_fail_counts: dict[str, int]   = defaultdict(int)
_fail_times:  dict[str, float] = defaultdict(float)
MAX_FAILS      = 5          # lock after this many failures
LOCKOUT_SECS   = 300        # 5-minute lockout


def _check_rate_limit(ip: str) -> bool:
    """Return True if IP is currently locked out."""
    if _fail_counts[ip] >= MAX_FAILS:
        elapsed = time.time() - _fail_times[ip]
        if elapsed < LOCKOUT_SECS:
            return True
        # lockout expired — reset
        _fail_counts[ip] = 0
    return False


def _record_failure(ip: str):
    _fail_counts[ip] += 1
    _fail_times[ip] = time.time()


def _reset_failures(ip: str):
    _fail_counts[ip] = 0


# ── Token helpers ─────────────────────────────────────────────

def make_token() -> str:
    expire = datetime.now(timezone.utc) + timedelta(hours=TOKEN_EXPIRE_H)
    return jwt.encode({"sub": TRACKER_USER, "exp": expire}, TRACKER_SECRET, algorithm=ALGORITHM)


def verify_token(token: str | None) -> bool:
    if not token:
        return False
    try:
        payload = jwt.decode(token, TRACKER_SECRET, algorithms=[ALGORITHM])
        return payload.get("sub") == TRACKER_USER
    except JWTError:
        return False


def _secure_cookie(response, token: str):
    """Attach a secure, HttpOnly, SameSite=Strict cookie."""
    response.set_cookie(
        key=COOKIE_NAME,
        value=token,
        httponly=True,        # JS cannot read this cookie
        samesite="strict",    # blocks CSRF
        secure=False,         # set True when served over HTTPS
        max_age=TOKEN_EXPIRE_H * 3600,
    )


# ── Login page ────────────────────────────────────────────────

@router.get("/login", response_class=HTMLResponse)
def get_login(tracker_session: str | None = Cookie(default=None)):
    if verify_token(tracker_session):
        return RedirectResponse("/tracker/")
    login_path = FRONTEND_DIR / "tracker_login.html"
    return HTMLResponse(login_path.read_text(encoding="utf-8"))


@router.post("/login")
def post_login(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
):
    """Validate credentials and issue a session cookie."""
    ip = request.client.host

    # Rate-limit check
    if _check_rate_limit(ip):
        error_html = (FRONTEND_DIR / "tracker_login.html").read_text(encoding="utf-8")
        error_html = error_html.replace(
            "<!--ERROR_PLACEHOLDER-->",
            '<p class="login-error">Too many failed attempts. Please wait 5 minutes.</p>',
        )
        return HTMLResponse(error_html, status_code=429)

    # Constant-time username check + bcrypt password check
    username_ok = username.strip() == TRACKER_USER
    password_ok = bcrypt.checkpw(password.encode(), TRACKER_PASS_HASH)

    if not username_ok or not password_ok:
        _record_failure(ip)
        remaining = MAX_FAILS - _fail_counts[ip]
        error_html = (FRONTEND_DIR / "tracker_login.html").read_text(encoding="utf-8")
        msg = "Invalid username or password."
        if remaining <= 2:
            msg += f" ({remaining} attempt{'s' if remaining != 1 else ''} left before lockout)"
        error_html = error_html.replace(
            "<!--ERROR_PLACEHOLDER-->",
            f'<p class="login-error">{msg}</p>',
        )
        return HTMLResponse(error_html, status_code=401)

    _reset_failures(ip)
    token = make_token()
    response = RedirectResponse("/tracker/", status_code=303)
    _secure_cookie(response, token)
    return response


# ── Dashboard (auth required) ─────────────────────────────────

@router.get("/", response_class=HTMLResponse)
def get_tracker(tracker_session: str | None = Cookie(default=None)):
    if not verify_token(tracker_session):
        return RedirectResponse("/tracker/login")
    tracker_path = FRONTEND_DIR / "report_tracker.html"
    return HTMLResponse(tracker_path.read_text(encoding="utf-8"))


# ── Static assets (auth required) ────────────────────────────

@router.get("/xlsx.min.js")
def get_sheetjs(tracker_session: str | None = Cookie(default=None)):
    if not verify_token(tracker_session):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)
    return FileResponse(FRONTEND_DIR / "xlsx.full.min.js", media_type="application/javascript")


# ── Google Sheet proxy (auth required) ────────────────────────

@router.get("/fetch-sheet")
def fetch_sheet(tracker_session: str | None = Cookie(default=None)):
    """Download the live Google Sheet as XLSX and return it to the frontend."""
    if not verify_token(tracker_session):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)

    url = (
        f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}"
        f"/export?format=xlsx&gid={GSHEET_GID}"
    )
    try:
        ssl_ctx = ssl.create_default_context(cafile=certifi.where())
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30, context=ssl_ctx) as resp:
            data = resp.read()
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Cache-Control": "no-store"},
        )
    except Exception as e:
        return JSONResponse(
            {"error": f"Failed to fetch Google Sheet: {e}"}, status_code=502
        )


# ── Logout ────────────────────────────────────────────────────

@router.post("/logout")
def logout():
    response = RedirectResponse("/tracker/login", status_code=303)
    response.delete_cookie(COOKIE_NAME)
    return response
