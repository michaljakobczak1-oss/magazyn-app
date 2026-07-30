import os
import sqlite3
from datetime import datetime, date
from pathlib import Path
from zoneinfo import ZoneInfo

# DATA_DIR – na Renderze ustaw na ścieżkę dysku trwałego (np. /var/data),
# inaczej po każdym deployu wraca pusta/stara baza z obrazu Dockera.
_DATA_DIR = Path(os.environ.get("DATA_DIR") or (Path(__file__).parent / "data"))
_DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = _DATA_DIR / "magazyn.db"
LOCAL_TZ = ZoneInfo("Europe/Warsaw")


def local_now():
    """Bieżąca data i godzina w Polsce (Europe/Warsaw)."""
    return datetime.now(LOCAL_TZ).replace(tzinfo=None)


def local_today():
    return local_now().date()

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'user',           -- 'admin' | 'user'
    active INTEGER NOT NULL DEFAULT 1,
    first_name TEXT,
    last_name TEXT,
    department TEXT,                             -- dział (np. Logistyka, Zakupy)
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS departments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS warehouses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    address TEXT,
    active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS logistics_partners (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    phone TEXT,
    email TEXT,
    active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS recipients (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,                          -- firma / miejsce / osoba
    contact_person TEXT,
    phone TEXT,
    address TEXT,
    email TEXT,
    last_used TEXT DEFAULT (datetime('now')),
    UNIQUE(name, address)
);

CREATE TABLE IF NOT EXISTS equipment (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,                   -- kod w aplikacji
    project_number TEXT,                         -- numer projektu
    name TEXT NOT NULL,                          -- nazwa sprzętu
    dimensions TEXT,                             -- wymiary
    photo TEXT,                                  -- nazwa pliku zdjęcia
    packaging_photo TEXT,                        -- zdjęcie opakowania / pakowania
    location TEXT,                               -- miejsce wewnątrz magazynu
    warehouse_id INTEGER REFERENCES warehouses(id),
    owner TEXT,                                  -- własność (czyj majątek)
    brand TEXT,                                  -- marka, której dotyczy materiał
    material_type TEXT NOT NULL DEFAULT 'klient',-- 'klient' | 'wlasny'
    condition TEXT NOT NULL DEFAULT 'sprawny',   -- sprawny | uszkodzony | do utylizacji
    condition_notes TEXT,
    damaged_quantity INTEGER NOT NULL DEFAULT 0, -- ile szt. ze stanu jest uszkodzonych
    storage_instructions TEXT,                   -- jak składować / pakować / transportować
    quantity INTEGER NOT NULL DEFAULT 1,
    notes TEXT,
    catalog TEXT NOT NULL DEFAULT 'main',        -- main | tcl
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS reservations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id INTEGER NOT NULL REFERENCES equipment(id),
    user_id INTEGER NOT NULL REFERENCES users(id),
    client TEXT,
    date_from TEXT NOT NULL,                     -- YYYY-MM-DD
    date_to TEXT NOT NULL,
    quantity INTEGER NOT NULL DEFAULT 1,
    status TEXT NOT NULL DEFAULT 'rezerwacja',   -- rezerwacja | wydane | wydane trwale | zwrócone | utylizacja | anulowana
    group_id TEXT,
    receiver TEXT,                               -- podwykonawca logistyczny (nazwa)
    permanent INTEGER NOT NULL DEFAULT 0,        -- 1 = wydanie trwałe (towar nie wraca)
    recipient_name TEXT,                         -- adresat towaru: firma / miejsce / osoba
    recipient_contact TEXT,                      -- osoba kontaktowa
    recipient_phone TEXT,
    recipient_address TEXT,
    recipient_email TEXT,
    damage INTEGER NOT NULL DEFAULT 0,           -- uszkodzenie odnotowane przy zwrocie
    damage_notes TEXT,
    notes TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    issued_at TEXT, issued_by INTEGER REFERENCES users(id),
    returned_at TEXT, returned_by INTEGER REFERENCES users(id)
);

CREATE INDEX IF NOT EXISTS idx_res_equipment ON reservations(equipment_id, status);

CREATE TABLE IF NOT EXISTS equipment_photos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id INTEGER NOT NULL REFERENCES equipment(id) ON DELETE CASCADE,
    filename TEXT NOT NULL,
    sort_order INTEGER NOT NULL DEFAULT 0,
    kind TEXT NOT NULL DEFAULT 'normal'          -- normal | damage
);
CREATE INDEX IF NOT EXISTS idx_eq_photos ON equipment_photos(equipment_id, sort_order);

-- Stan magazynowy per magazyn/miejsce (ten sam kod może leżeć w kilku lokalizacjach)
CREATE TABLE IF NOT EXISTS equipment_stock (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id INTEGER NOT NULL REFERENCES equipment(id) ON DELETE CASCADE,
    warehouse_id INTEGER REFERENCES warehouses(id),
    location TEXT NOT NULL DEFAULT '',
    quantity INTEGER NOT NULL DEFAULT 0,
    UNIQUE(equipment_id, warehouse_id, location)
);
CREATE INDEX IF NOT EXISTS idx_eq_stock ON equipment_stock(equipment_id);
"""

# kolumny dokładane migracją do starszych baz: tabela -> {kolumna: definicja}
MIGRATIONS = {
    "users": {
        "first_name": "TEXT",
        "last_name": "TEXT",
        "department": "TEXT",
    },
    "equipment": {
        "owner": "TEXT",
        "brand": "TEXT",
        "packaging_photo": "TEXT",
        "warehouse_id": "INTEGER REFERENCES warehouses(id)",
        "material_type": "TEXT NOT NULL DEFAULT 'klient'",
        "condition": "TEXT NOT NULL DEFAULT 'sprawny'",
        "condition_notes": "TEXT",
        "damaged_quantity": "INTEGER NOT NULL DEFAULT 0",
        "storage_instructions": "TEXT",
        "catalog": "TEXT NOT NULL DEFAULT 'main'",
    },
    "reservations": {
        "group_id": "TEXT",
        "receiver": "TEXT",
        "recipient_name": "TEXT",
        "recipient_contact": "TEXT",
        "recipient_phone": "TEXT",
        "recipient_address": "TEXT",
        "recipient_email": "TEXT",
        "damage": "INTEGER NOT NULL DEFAULT 0",
        "damage_notes": "TEXT",
        "permanent": "INTEGER NOT NULL DEFAULT 0",
        "return_warehouse_id": "INTEGER REFERENCES warehouses(id)",
        "return_location": "TEXT",
        "project_number": "TEXT",
        "issue_warehouse_id": "INTEGER REFERENCES warehouses(id)",
        "issue_location": "TEXT",
        "returner": "TEXT",
    },
    "equipment_photos": {
        "kind": "TEXT NOT NULL DEFAULT 'normal'",
    },
}


def get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con


def init_db():
    con = get_db()
    con.executescript(SCHEMA)
    for table, cols in MIGRATIONS.items():
        existing = [r["name"] for r in con.execute(f"PRAGMA table_info({table})")]
        for col, ddl in cols.items():
            if col not in existing:
                con.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")
    con.commit()

    # migracja: materiały z kodem 00... oznacz jako własne
    con.execute("UPDATE equipment SET material_type='wlasny' WHERE code LIKE '00%' AND material_type='klient'")
    con.commit()

    # migracja: istniejące photo -> equipment_photos (max 1 na start)
    orphans = con.execute(
        """SELECT id, photo FROM equipment
           WHERE IFNULL(photo,'')!='' AND id NOT IN
           (SELECT DISTINCT equipment_id FROM equipment_photos)""").fetchall()
    for row in orphans:
        con.execute(
            "INSERT INTO equipment_photos (equipment_id, filename, sort_order) VALUES (?,?,0)",
            (row["id"], row["photo"]))
    if orphans:
        con.commit()

    # migracja: dział Warrens (zarządzanie TCL)
    if con.execute("SELECT COUNT(*) c FROM departments WHERE name=?", ("Warrens",)).fetchone()["c"] == 0:
        con.execute("INSERT INTO departments (name, active) VALUES (?,1)", ("Warrens",))
        con.commit()

    # migracja: dotychczasowa sztywna lista odbierających -> słownik podwykonawców
    if con.execute("SELECT COUNT(*) c FROM logistics_partners").fetchone()["c"] == 0:
        for name in ("Markosik", "Stefaniak"):
            con.execute("INSERT OR IGNORE INTO logistics_partners (name) VALUES (?)", (name,))
        con.commit()

    # jeśli oznaczono utylizację części sztuk, a coś zostało – nie trzymaj statusu „do utylizacji” na całym kodzie
    con.execute("""UPDATE equipment SET condition='sprawny'
                   WHERE condition='do utylizacji' AND quantity > 0""")
    con.commit()

    # migracja: stan per magazyn z dotychczasowego warehouse_id / location / quantity
    still_missing = con.execute(
        """SELECT e.id AS eid, e.warehouse_id AS wid, IFNULL(e.location,'') AS loc, e.quantity AS qty
           FROM equipment e
           WHERE e.quantity > 0
             AND e.id NOT IN (SELECT DISTINCT equipment_id FROM equipment_stock WHERE quantity > 0)"""
    ).fetchall()
    for row in still_missing:
        con.execute(
            """INSERT INTO equipment_stock (equipment_id, warehouse_id, location, quantity)
               VALUES (?,?,?,?)""",
            (row["eid"], row["wid"], row["loc"], row["qty"]),
        )
    if still_missing:
        con.commit()

    # pierwszy admin, jeśli brak użytkowników
    if con.execute("SELECT COUNT(*) c FROM users").fetchone()["c"] == 0:
        from werkzeug.security import generate_password_hash
        con.execute(
            "INSERT INTO users (username, password_hash, role, first_name, last_name) VALUES (?,?,?,?,?)",
            ("admin", generate_password_hash("admin123", method="pbkdf2:sha256"), "admin", "Administrator", ""),
        )
        con.commit()
    con.close()


def reserved_qty(con, equipment_id, date_from, date_to, exclude_id=None):
    """Suma sztuk zajętych w zapytanym terminie.

    - rezerwacja / wydane: gdy termin nakłada się na zakres (włącznie z dniem zwrotu)
    - wydane po terminie zwrotu (przetrzymane): zawsze, do czasu przyjęcia
    """
    today = local_today().isoformat()
    q = """SELECT COALESCE(SUM(quantity),0) s FROM reservations
           WHERE equipment_id=? AND (
             (status IN ('rezerwacja','wydane') AND date_from <= ? AND date_to >= ?)
             OR (status='wydane' AND date_to < ?)
           )"""
    params = [equipment_id, date_to, date_from, today]
    if exclude_id:
        q += " AND id != ?"
        params.append(exclude_id)
    return con.execute(q, params).fetchone()["s"]


def handoff_conflict(con, equipment_id, date_from, date_to, exclude_id=None):
    """True, gdy w dniu zwrotu innej rezerwacji próbuje się zacząć nowa (lub odwrotnie).

    Zwrot 21. → kolejne wypożyczenie dopiero od 22. (bez stykania terminów).
    """
    q = """SELECT date_from, date_to, quantity FROM reservations
           WHERE equipment_id=? AND status IN ('rezerwacja','wydane')
           AND (date_to = ? OR date_from = ?)"""
    params = [equipment_id, date_from, date_to]
    if exclude_id:
        q += " AND id != ?"
        params.append(exclude_id)
    return con.execute(q, params).fetchone() is not None


def next_free_after_return(date_to):
    """Pierwszy dzień, w którym wolno zacząć wypożyczenie po zwrocie w date_to."""
    from datetime import timedelta
    return (date.fromisoformat(str(date_to)[:10]) + timedelta(days=1)).isoformat()


def display_name(row):
    """Imię i nazwisko użytkownika; fallback na login."""
    fn = (row["first_name"] or "").strip() if "first_name" in row.keys() else ""
    ln = (row["last_name"] or "").strip() if "last_name" in row.keys() else ""
    full = f"{fn} {ln}".strip()
    return full or row["username"]


def upsert_recipient(con, name, contact, phone, address, email):
    """Słownik adresatów: aktualizuje wpis (name+address) albo tworzy nowy."""
    name = (name or "").strip()
    if not name:
        return
    row = con.execute("SELECT id FROM recipients WHERE name=? AND IFNULL(address,'')=IFNULL(?,'')",
                      (name, (address or "").strip())).fetchone()
    if row:
        con.execute("""UPDATE recipients SET contact_person=?, phone=?, email=?,
                       last_used=datetime('now') WHERE id=?""",
                    ((contact or "").strip(), (phone or "").strip(),
                     (email or "").strip(), row["id"]))
    else:
        con.execute("""INSERT INTO recipients (name, contact_person, phone, address, email)
                       VALUES (?,?,?,?,?)""",
                    (name, (contact or "").strip(), (phone or "").strip(),
                     (address or "").strip(), (email or "").strip()))


def equipment_photo_rows(con, equipment_id):
    """Wiersze galerii: filename + kind (normal|damage)."""
    return con.execute(
        """SELECT filename, IFNULL(kind, 'normal') AS kind
           FROM equipment_photos WHERE equipment_id=?
           ORDER BY sort_order, id""",
        (equipment_id,)).fetchall()


def equipment_photo_list(con, equipment_id):
    """Lista nazw plików zdjęć sprzętu (kolejność sort_order)."""
    return [r["filename"] for r in equipment_photo_rows(con, equipment_id)]


def equipment_photo_kind_map(con, equipment_id):
    """filename -> kind, do zachowania typu przy edycji karty."""
    return {r["filename"]: r["kind"] for r in equipment_photo_rows(con, equipment_id)}


def sync_equipment_primary_photo(con, equipment_id):
    """Ustaw equipment.photo na pierwsze zdjęcie z galerii (miniatury / kompatybilność)."""
    photos = equipment_photo_list(con, equipment_id)
    con.execute("UPDATE equipment SET photo=? WHERE id=?",
                (photos[0] if photos else None, equipment_id))


def ensure_equipment_stock(con, equipment_id):
    """Gwarantuje wiersze stock; przy braku tworzy jeden z equipment.warehouse/location/qty."""
    rows = con.execute(
        "SELECT * FROM equipment_stock WHERE equipment_id=? AND quantity > 0 ORDER BY quantity DESC, id",
        (equipment_id,)).fetchall()
    if rows:
        return rows
    eq = con.execute("SELECT warehouse_id, IFNULL(location,'') loc, quantity FROM equipment WHERE id=?",
                     (equipment_id,)).fetchone()
    if not eq or not eq["quantity"] or eq["quantity"] <= 0:
        return []
    con.execute(
        """INSERT INTO equipment_stock (equipment_id, warehouse_id, location, quantity)
           VALUES (?,?,?,?)""",
        (equipment_id, eq["warehouse_id"], eq["loc"], eq["quantity"]))
    return con.execute(
        "SELECT * FROM equipment_stock WHERE equipment_id=? AND quantity > 0 ORDER BY quantity DESC, id",
        (equipment_id,)).fetchall()


def replace_equipment_stock(con, equipment_id, warehouse_id, location, quantity):
    """Ustawia stan jako jeden magazyn/miejsce (edycja karty / import)."""
    con.execute("DELETE FROM equipment_stock WHERE equipment_id=?", (equipment_id,))
    qty = int(quantity or 0)
    if qty > 0:
        con.execute(
            """INSERT INTO equipment_stock (equipment_id, warehouse_id, location, quantity)
               VALUES (?,?,?,?)""",
            (equipment_id, warehouse_id, (location or "").strip(), qty))
    sync_equipment_from_stock(con, equipment_id)


def _stock_key_clause(warehouse_id, location):
    loc = (location or "").strip()
    if warehouse_id is None:
        return "warehouse_id IS NULL AND location=?", [loc]
    return "warehouse_id=? AND location=?", [warehouse_id, loc]


def add_equipment_stock(con, equipment_id, warehouse_id, location, qty):
    """Dodaje sztuki do wskazanego magazynu/miejsca."""
    qty = int(qty)
    if qty <= 0:
        return
    loc = (location or "").strip()
    where, params = _stock_key_clause(warehouse_id, loc)
    row = con.execute(
        f"SELECT id, quantity FROM equipment_stock WHERE equipment_id=? AND {where}",
        [equipment_id, *params]).fetchone()
    if row:
        con.execute("UPDATE equipment_stock SET quantity=quantity+? WHERE id=?",
                    (qty, row["id"]))
    else:
        con.execute(
            """INSERT INTO equipment_stock (equipment_id, warehouse_id, location, quantity)
               VALUES (?,?,?,?)""",
            (equipment_id, warehouse_id, loc, qty))


def take_equipment_stock(con, equipment_id, qty, prefer_warehouse_id=None):
    """Zdejmuje sztuki ze stanu (najpierw preferowany magazyn, potem największe stany)."""
    qty = int(qty)
    if qty <= 0:
        return True
    ensure_equipment_stock(con, equipment_id)
    rows = list(con.execute(
        "SELECT * FROM equipment_stock WHERE equipment_id=? AND quantity > 0 ORDER BY quantity DESC, id",
        (equipment_id,)).fetchall())
    if prefer_warehouse_id is not None:
        rows.sort(key=lambda r: (
            0 if r["warehouse_id"] == prefer_warehouse_id else 1,
            -int(r["quantity"]),
            r["id"],
        ))
    left = qty
    for r in rows:
        if left <= 0:
            break
        take = min(int(r["quantity"]), left)
        new_q = int(r["quantity"]) - take
        if new_q <= 0:
            con.execute("DELETE FROM equipment_stock WHERE id=?", (r["id"],))
        else:
            con.execute("UPDATE equipment_stock SET quantity=? WHERE id=?", (new_q, r["id"]))
        left -= take
    return left == 0


def move_equipment_stock_on_return(con, equipment_id, qty, to_warehouse_id, to_location):
    """Przy zwrocie: przenosi qty szt. ze starego stanu na magazyn/miejsce przyjęcia.

    Nie zmienia łącznej quantity sprzętu – tylko rozkład po magazynach.
    """
    qty = int(qty)
    if qty <= 0:
        return False
    eq = con.execute("SELECT warehouse_id FROM equipment WHERE id=?", (equipment_id,)).fetchone()
    prefer = eq["warehouse_id"] if eq else None
    ensure_equipment_stock(con, equipment_id)
    take_equipment_stock(con, equipment_id, qty, prefer_warehouse_id=prefer)
    add_equipment_stock(con, equipment_id, to_warehouse_id, to_location, qty)
    sync_equipment_from_stock(con, equipment_id, keep_total=True)
    return True


def sync_equipment_from_stock(con, equipment_id, keep_total=False):
    """Ustawia equipment.warehouse_id / location / (opcjonalnie quantity) wg stock.

    keep_total=True: nie rusza equipment.quantity (zwrot tylko zmienia rozkład).
    """
    rows = con.execute(
        """SELECT es.*, w.name AS warehouse_name
           FROM equipment_stock es
           LEFT JOIN warehouses w ON w.id=es.warehouse_id
           WHERE es.equipment_id=? AND es.quantity > 0
           ORDER BY es.quantity DESC, es.id""",
        (equipment_id,)).fetchall()
    if not rows:
        if not keep_total:
            con.execute(
                "UPDATE equipment SET quantity=0, warehouse_id=NULL, location='' WHERE id=?",
                (equipment_id,))
        return

    primary = rows[0]
    # Główny magazyn = ten z największą ilością; lokalizacja bez sklejania multi-stock
    loc = primary["location"] or ""

    if keep_total:
        con.execute(
            "UPDATE equipment SET warehouse_id=?, location=? WHERE id=?",
            (primary["warehouse_id"], loc, equipment_id))
    else:
        stock_sum = sum(int(r["quantity"]) for r in rows)
        con.execute(
            "UPDATE equipment SET warehouse_id=?, location=?, quantity=? WHERE id=?",
            (primary["warehouse_id"], loc, stock_sum, equipment_id))


def stock_summary(con, equipment_id):
    """Krótki opis rozkładu stanu, np. 'Łowicz: 4, Stalowa: 1'."""
    rows = con.execute(
        """SELECT es.quantity, IFNULL(es.location,'') loc, w.name AS warehouse_name
           FROM equipment_stock es
           LEFT JOIN warehouses w ON w.id=es.warehouse_id
           WHERE es.equipment_id=? AND es.quantity > 0
           ORDER BY es.quantity DESC, es.id""",
        (equipment_id,)).fetchall()
    if not rows:
        return ""
    parts = []
    for r in rows:
        wh = r["warehouse_name"] or "—"
        loc = (r["loc"] or "").strip()
        label = wh + (f"/{loc}" if loc else "")
        parts.append(f"{label}: {int(r['quantity'])}")
    return ", ".join(parts)
