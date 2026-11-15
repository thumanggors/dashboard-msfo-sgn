"""Ekstrak Excel -> data.json, lalu git push bila Excel berubah. Pakai: py sync_github.py [--once]"""
import base64
import json
import os
import subprocess
import sys
import time
from datetime import datetime

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

import server

OUT = server.BASE_DIR / "data.enc.json"
INTERVAL = 60  # detik
ITER = 200_000  # harus sama dengan index.html


def encrypt(text):
    """AES-GCM, kunci dari password.txt (PBKDF2-SHA256). Repo hanya berisi ciphertext."""
    password = (server.BASE_DIR / "password.txt").read_text(encoding="utf-8").strip()
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(hashes.SHA256(), 32, salt, ITER).derive(password.encode())
    b64 = lambda b: base64.b64encode(b).decode()
    return json.dumps({"salt": b64(salt), "iv": b64(iv), "data": b64(AESGCM(key).encrypt(iv, text.encode(), None))})


def git(*args):
    return subprocess.run(["git", *args], cwd=server.BASE_DIR, capture_output=True, text=True)


def sync(last_mtime):
    src = server.resolve_source()
    mtime = src.stat().st_mtime_ns
    if mtime == last_mtime:
        return mtime
    payload = server.extract_dashboard(src)
    # ponytail: path lokal tidak boleh bocor ke repo publik
    payload["meta"].update(source=src.name, stale=False, warning=None)
    OUT.write_text(encrypt(json.dumps(payload, ensure_ascii=False)), encoding="utf-8")
    git("add", OUT.name)
    if git("diff", "--cached", "--quiet").returncode == 0:
        return mtime
    git("commit", "-m", f"Update data {datetime.now():%Y-%m-%d %H:%M}")
    result = git("push")
    print(datetime.now().strftime("%H:%M:%S"), "push:", "OK" if result.returncode == 0 else result.stderr.strip())
    return mtime


if __name__ == "__main__":
    last = None
    while True:
        try:
            last = sync(last)
        except Exception as exc:
            print(datetime.now().strftime("%H:%M:%S"), "gagal:", exc)
        if "--once" in sys.argv:
            break
        time.sleep(INTERVAL)
