"""
create_credentials.py
─────────────────────────────────────────────────────────────────
Run this once to set your username and password for the Report Tracker.

Usage:
    python create_credentials.py

It will:
  1. Prompt for your chosen username
  2. Prompt for your chosen password (hidden input)
  3. Hash the password with bcrypt
  4. Write TRACKER_USER and TRACKER_PASS_HASH into .env

You do NOT need to store your password anywhere after this — only the
bcrypt hash is saved, which cannot be reversed.
"""

import getpass
import re
from pathlib import Path

try:
    import bcrypt
except ImportError:
    raise SystemExit("Run:  pip install bcrypt  then retry.")

ENV_FILE = Path(__file__).parent / ".env"


def update_env(key: str, value: str):
    """Replace or append KEY=value in .env (in-place)."""
    text = ENV_FILE.read_text(encoding="utf-8") if ENV_FILE.exists() else ""
    pattern = re.compile(rf"^{re.escape(key)}=.*$", re.MULTILINE)
    new_line = f"{key}={value}"
    if pattern.search(text):
        text = pattern.sub(new_line, text)
    else:
        text = text.rstrip("\n") + f"\n{new_line}\n"
    ENV_FILE.write_text(text, encoding="utf-8")


def main():
    print("=" * 55)
    print("  Anderson Lab — Report Tracker Credential Setup")
    print("=" * 55)
    print()

    username = input("Choose a username: ").strip()
    if not username:
        raise SystemExit("Username cannot be empty.")

    while True:
        try:
            password = getpass.getpass("Choose a password (hidden): ")
        except Exception:
            # Fallback for terminals that don't support hidden input (e.g. VS Code)
            password = input("Choose a password (visible — delete history after): ")
        if len(password) < 8:
            print("  Password must be at least 8 characters. Try again.")
            continue
        try:
            confirm = getpass.getpass("Confirm password (hidden): ")
        except Exception:
            confirm = input("Confirm password: ")
        if password != confirm:
            print("  Passwords do not match. Try again.")
            continue
        break

    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt(rounds=12)).decode()

    update_env("TRACKER_USER",      username)
    update_env("TRACKER_PASS_HASH", hashed)

    print()
    print(f"  Username : {username}")
    print(f"  Hash     : {hashed[:20]}…  (stored in .env)")
    print()
    print("Done! Start the server:  uvicorn backend:app --reload")
    print("Then visit:              http://localhost:8000/tracker/login")


if __name__ == "__main__":
    main()
