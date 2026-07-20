import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "magazyn.db"

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
    storage_instructions TEXT,                   -- jak składować / pakować / transportować
    quantity INTEGER NOT NULL DEFAULT 1,
    notes TEXT,
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
    status TEXT NOT NULL DEFAULT 'rezerwacja',   -- rezerwacja | wydane | wydane trwale | zwrócone | anulowana
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
    sort_order INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_eq_photos ON equipment_photos(equipment_id, sort_order);
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
        "storage_instructions": "TEXT",
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

    # migracja: dotychczasowa sztywna lista odbierających -> słownik podwykonawców
    if con.execute("SELECT COUNT(*) c FROM logistics_partners").fetchone()["c"] == 0:
        for name in ("Markosik", "Stefaniak"):
            con.execute("INSERT OR IGNORE INTO logistics_partners (name) VALUES (?)", (name,))
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
    """Suma sztuk zajętych w terminie.

    - rezerwacja: tylko gdy termin nakłada się na zapytany zakres
    - wydane: zawsze (towar jest poza magazynem do zwrotu)
    """
    q = """SELECT COALESCE(SUM(quantity),0) s FROM reservations
           WHERE equipment_id=? AND (
             (status='rezerwacja' AND date_from <= ? AND date_to >= ?)
             OR status='wydane'
           )"""
    params = [equipment_id, date_to, date_from]
    if exclude_id:
        q += " AND id != ?"
        params.append(exclude_id)
    return con.execute(q, params).fetchone()["s"]


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


def equipment_photo_list(con, equipment_id):
    """Lista nazw plików zdjęć sprzętu (kolejność sort_order)."""
    return [r["filename"] for r in con.execute(
        "SELECT filename FROM equipment_photos WHERE equipment_id=? ORDER BY sort_order, id",
        (equipment_id,)).fetchall()]


def sync_equipment_primary_photo(con, equipment_id):
    """Ustaw equipment.photo na pierwsze zdjęcie z galerii (miniatury / kompatybilność)."""
    photos = equipment_photo_list(con, equipment_id)
    con.execute("UPDATE equipment SET photo=? WHERE id=?",
                (photos[0] if photos else None, equipment_id))
