#!/usr/bin/env python3
"""Usuwa wizyty magazynowe na prod (HTTP) lub lokalnie (--db)."""
from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from http.cookiejar import CookieJar
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def purge_db(db_path: str, keep: list[int], restore: bool):
    import db as dbmod

    dbmod.DB_PATH = Path(db_path)
    dbmod.init_db()
    con = dbmod.get_db()
    all_ids = [int(r["id"]) for r in con.execute("SELECT id FROM warehouse_visits ORDER BY id").fetchall()]
    delete_ids = [i for i in all_ids if i not in keep]
    for did in delete_ids:
        con.execute("DELETE FROM warehouse_visits WHERE id=?", (did,))
    if restore:
        for kid in keep:
            con.execute(
                """UPDATE warehouse_visits SET status='planowane',
                   completion_notes=NULL, completed_at=NULL, completed_by=NULL
                   WHERE id=?""",
                (kid,),
            )
    con.commit()
    con.close()
    return {"ok": True, "deleted": delete_ids, "kept": keep, "restored": restore}


def purge_http(base_url: str, login_user: str, password: str, keep: list[int], restore: bool):
    base = base_url.rstrip("/")
    jar = CookieJar()
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
    login_data = urllib.parse.urlencode({"username": login_user, "password": password}).encode()
    op.open(f"{base}/login", login_data, timeout=120)
    payload = [("confirm", "USUN")]
    if restore:
        payload.append(("restore", "1"))
    for kid in keep:
        payload.append(("keep", str(kid)))
    body = urllib.parse.urlencode(payload).encode()
    req = urllib.request.Request(
        f"{base}/sprawdzenia/purge?json=1",
        data=body,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with op.open(req, timeout=120) as resp:
            raw = resp.read().decode()
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode()
        raise SystemExit(f"HTTP {exc.code}: {raw[:500]}") from exc
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"Usunięto (\d+) wizyt", raw)
        if m:
            return {"ok": True, "count": int(m.group(1)), "deleted": [], "kept": keep}
        raise SystemExit(f"Nieoczekiwana odpowiedź: {raw[:500]}")


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--url", help="URL produkcji")
    p.add_argument("--login-user", help="Login HTTP (admin)")
    p.add_argument("--password", help="Hasło HTTP")
    p.add_argument("--db", help="Ścieżka do magazyn.db (lokalnie)")
    p.add_argument("--keep", action="append", type=int, default=[], help="ID wizyt do zostawienia")
    p.add_argument("--restore", action="store_true", help="Przywróć status planowane dla zachowanych")
    args = p.parse_args()

    if args.db:
        data = purge_db(args.db, args.keep, args.restore)
    elif args.url and args.login_user and args.password:
        data = purge_http(args.url, args.login_user, args.password, args.keep, args.restore)
    else:
        p.error("Podaj --db albo --url z --login-user i --password")
    print(json.dumps(data, ensure_ascii=False))


if __name__ == "__main__":
    main()
