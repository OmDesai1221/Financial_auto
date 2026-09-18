"""
Simple single-user login for this local tool. The username is fixed
("kishor"); the password starts as the hardcoded default below but can be
changed (only by providing the current password first) — the change is
saved to a local file, hashed, so it survives the app being restarted.
That file is gitignored: it's specific to whichever machine is running
the tool, never something to commit or share.
"""
import os
import json
from werkzeug.security import generate_password_hash, check_password_hash

USERNAME = "kishor"
DEFAULT_PASSWORD = "Shree@1977"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AUTH_FILE = os.path.join(BASE_DIR, "auth_data.json")


def _read_stored_hash():
    if not os.path.isfile(AUTH_FILE):
        return None
    try:
        with open(AUTH_FILE, "r") as f:
            data = json.load(f)
        return data.get("password_hash")
    except (json.JSONDecodeError, OSError):
        return None


def _write_stored_hash(password_hash):
    with open(AUTH_FILE, "w") as f:
        json.dump({"password_hash": password_hash}, f)


def _current_password_hash():
    stored = _read_stored_hash()
    if stored:
        return stored
    # first run ever — seed the file with the default password's hash so
    # behavior is consistent from here on (verify + change both read/write
    # the same file rather than special-casing "no file yet")
    default_hash = generate_password_hash(DEFAULT_PASSWORD)
    _write_stored_hash(default_hash)
    return default_hash


def verify_login(username, password):
    return username == USERNAME and check_password_hash(_current_password_hash(), password)


def change_password(old_password, new_password):
    """Returns (ok, error_message). Requires the current password to
    succeed — this is the only way to change it, there's no recovery/reset
    path by design for a single hardcoded username."""
    if not check_password_hash(_current_password_hash(), old_password):
        return False, "Current password is incorrect."
    if not new_password or len(new_password) < 4:
        return False, "New password must be at least 4 characters."
    _write_stored_hash(generate_password_hash(new_password))
    return True, None
