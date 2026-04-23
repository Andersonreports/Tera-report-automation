"""
tracker_auth.py
───────────────────────────────────────────────────────────────
Secure router for the Anderson Lab Report Tracker.

Endpoints mounted at /tracker by backend.py
  GET  /tracker/login      → serve login page HTML
  POST /tracker/login      → validate credentials, set HttpOnly cookie
  POST /tracker/auth-sso   → issue tracker cookie via Supabase JWT (SSO)
  GET  /tracker/           → serve tracker dashboard (auth required)
  GET  /tracker/xlsx.min.js → serve bundled SheetJS (auth required)
  POST /tracker/logout     → clear session cookie

Credentials are stored ONLY in .env:
  TRACKER_USER=your_username
  TRACKER_PASS_HASH=<bcrypt hash>   ← run create_credentials.py to generate
  TRACKER_SECRET=<random 64-char hex>   ← used to sign JWTs
"""

import os
import time
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

import bcrypt
from dotenv import load_dotenv
from fastapi import APIRouter, Cookie, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from jose import JWTError, jwt

load_dotenv()

# ── Config ────────────────────────────────────────────────────
TRACKER_USER      = os.getenv("TRACKER_USER", "")
TRACKER_PASS_HASH = os.getenv("TRACKER_PASS_HASH", "").encode()
TRACKER_SECRET    = os.getenv("TRACKER_SECRET", "change-me-set-TRACKER_SECRET-in-env")
TOKEN_EXPIRE_H    = int(os.getenv("TRACKER_SESSION_HOURS", "8"))
ALGORITHM         = "HS256"

_CREDS_CONFIGURED = bool(TRACKER_USER and TRACKER_PASS_HASH and os.getenv("TRACKER_SECRET"))

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


# ── SSO: issue tracker cookie via Supabase JWT ────────────────

@router.post("/auth-sso")
async def sso_auth(request: Request):
    """
    Accept a Supabase access_token (Bearer), verify it against Supabase,
    check that the user's profile section allows tracker access,
    then issue a tracker_session cookie so the user skips the manual login.
    """
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return JSONResponse({"error": "Missing token"}, status_code=401)

    access_token = auth_header[7:]

    try:
        from supabase_client import _get_client
        client = _get_client()

        user_resp = client.auth.get_user(access_token)
        if not user_resp or not user_resp.user:
            return JSONResponse({"error": "Invalid token"}, status_code=401)

        user_id = user_resp.user.id

        profile_resp = (
            client.table("profiles")
            .select("section, role")
            .eq("id", user_id)
            .maybe_single()
            .execute()
        )
        profile  = profile_resp.data or {}
        section  = profile.get("section") or "all"
        role     = profile.get("role")    or "user"

        if role == "admin" or section in ("all", "general"):
            token    = make_token()
            response = JSONResponse({"ok": True})
            _secure_cookie(response, token)
            return response

        return JSONResponse({"error": "Access denied"}, status_code=403)

    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)


# ── Login page ────────────────────────────────────────────────

_NOT_CONFIGURED_HTML = """<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Tracker — Not Configured</title>
<style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;
min-height:100vh;margin:0;background:#f1f5f9;}
.box{background:#fff;border-radius:12px;padding:40px 48px;text-align:center;
box-shadow:0 4px 24px rgba(0,0,0,.1);max-width:480px;}
h2{color:#dc2626;margin-bottom:12px;}p{color:#64748b;line-height:1.6;}
code{background:#f8fafc;padding:2px 6px;border-radius:4px;font-size:13px;color:#0f172a;}
</style></head><body><div class="box">
<h2>Tracker Not Configured</h2>
<p>Set these environment variables on the server, then restart:</p>
<p><code>TRACKER_USER</code> &nbsp; <code>TRACKER_PASS_HASH</code> &nbsp; <code>TRACKER_SECRET</code></p>
<p>Run <code>python create_credentials.py</code> in the backend folder to generate the hash.</p>
</div></body></html>"""


@router.get("/login", response_class=HTMLResponse)
def get_login(tracker_session: str | None = Cookie(default=None)):
    if not _CREDS_CONFIGURED:
        return HTMLResponse(_NOT_CONFIGURED_HTML, status_code=503)
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
    if not _CREDS_CONFIGURED:
        return HTMLResponse(_NOT_CONFIGURED_HTML, status_code=503)
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
    if not _CREDS_CONFIGURED:
        return RedirectResponse("/tracker/login")
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
async def fetch_sheet(tracker_session: str | None = Cookie(default=None)):
    """Download the live Google Sheet as XLSX and return it to the frontend."""
    if not verify_token(tracker_session):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)

    url = (
        f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}"
        f"/export?format=xlsx"
    )
    try:
        import httpx
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
        }
        async with httpx.AsyncClient(follow_redirects=True, timeout=60) as client:
            resp = await client.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.content
        print(f"[tracker] fetch-sheet: {len(data)} bytes, status {resp.status_code}")
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Cache-Control": "no-store"},
        )
    except Exception as e:
        print(f"[tracker] fetch-sheet error: {e}")
        return JSONResponse(
            {"error": f"Failed to fetch Google Sheet: {e}"}, status_code=502
        )


# ── Logout ────────────────────────────────────────────────────

@router.post("/logout")
def logout():
    response = RedirectResponse("/tracker/login", status_code=303)
    response.delete_cookie(COOKIE_NAME)
    return response
