"""
Jednorazowy import rejestru sprzętu z istniejącego arkusza Excel (pkt 2.6).

Użycie:
    python import_excel.py plik.xlsx [--photos katalog_ze_zdjeciami] [--warehouse "Nazwa magazynu"] [--dry-run]

Oczekiwane kolumny w pierwszym wierszu arkusza (nagłówki, wielkość liter bez znaczenia,
kolejność dowolna – dopasowanie po nazwie):
    kod | zdjęcie | nazwa | ilość | rozmiar/waga/opis (lub: wymiary / opis) | miejsce | własność | projekt | brand (opcjonalnie)

Zdjęcia – dwa wspierane warianty:
    1) katalog --photos z plikami nazwanymi kodem sprzętu (np. CC0043.png / cc0043.jpg),
    2) zdjęcia osadzone w arkuszu (xlsx) – skrypt wyciąga je automatycznie i dopasowuje
       po wierszu zakotwiczenia.

Wymaga: pip install openpyxl
"""
import argparse
import re
import shutil
import sys
import uuid
from pathlib import Path

from openpyxl import load_workbook

from db import get_db, init_db

BASE = Path(__file__).parent
UPLOAD_DIR = BASE / "static" / "uploads"
ALLOWED_EXT = {".png", ".jpg", ".jpeg", ".gif", ".webp"}

# nagłówek w arkuszu -> pole w bazie
HEADER_MAP = {
    "kod": "code",
    "zdjęcie": "photo", "zdjecie": "photo", "foto": "photo",
    "nazwa": "name",
    "ilość": "quantity", "ilosc": "quantity", "szt": "quantity",
    "rozmiar/waga/opis": "dimensions", "rozmiar": "dimensions",
    "wymiary": "dimensions", "opis": "dimensions",
    "miejsce": "location", "miejsce w magazynie": "location",
    "własność": "owner", "wlasnosc": "owner",
    "projekt": "project_number", "numer projektu": "project_number", "nr projektu": "project_number",
    "brand": "brand", "marka": "brand",
}


def norm(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def map_headers(ws):
    headers = {}
    for idx, cell in enumerate(ws[1], start=1):
        key = HEADER_MAP.get(norm(cell.value))
        if key:
            headers[key] = idx
    missing = {"code", "name"} - set(headers)
    if missing:
        sys.exit(f"Brak wymaganych kolumn w arkuszu: {', '.join(missing)}. "
                 f"Znalezione nagłówki: {[c.value for c in ws[1]]}")
    return headers


def extract_embedded_images(ws):
    """Zwraca {nr_wiersza: bytes} dla zdjęć osadzonych w arkuszu."""
    out = {}
    for img in getattr(ws, "_images", []):
        try:
            row = img.anchor._from.row + 1  # 0-indexed
            out[row] = img._data()
        except Exception:
            pass
    return out


def find_photo_file(photos_dir, code):
    if not photos_dir:
        return None
    for ext in ALLOWED_EXT:
        for candidate in (code, code.lower(), code.upper()):
            p = photos_dir / f"{candidate}{ext}"
            if p.exists():
                return p
    return None


def save_photo_bytes(data, ext=".png"):
    fname = f"{uuid.uuid4().hex}{ext}"
    (UPLOAD_DIR / fname).write_bytes(data)
    return fname


def save_photo_file(path):
    fname = f"{uuid.uuid4().hex}{path.suffix.lower()}"
    shutil.copy(path, UPLOAD_DIR / fname)
    return fname


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="plik .xlsx z rejestrem")
    ap.add_argument("--photos", help="katalog ze zdjęciami nazwanymi kodem sprzętu")
    ap.add_argument("--warehouse", help="nazwa magazynu przypisywanego importowanym pozycjom")
    ap.add_argument("--dry-run", action="store_true", help="tylko pokaż, nic nie zapisuj")
    args = ap.parse_args()

    init_db()
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    photos_dir = Path(args.photos) if args.photos else None

    wb = load_workbook(args.xlsx)
    ws = wb.active
    headers = map_headers(ws)
    embedded = extract_embedded_images(ws)

    con = get_db()

    warehouse_id = None
    if args.warehouse:
        row = con.execute("SELECT id FROM warehouses WHERE name=?", (args.warehouse,)).fetchone()
        if row:
            warehouse_id = row["id"]
        elif not args.dry_run:
            cur = con.execute("INSERT INTO warehouses (name) VALUES (?)", (args.warehouse,))
            warehouse_id = cur.lastrowid
            print(f"Utworzono magazyn: {args.warehouse}")

    added, skipped, no_photo = 0, 0, 0
    for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        def val(field):
            col = headers.get(field)
            return row[col - 1].value if col else None

        code = str(val("code") or "").strip()
        name = str(val("name") or "").strip()
        if not code or not name:
            continue

        if con.execute("SELECT 1 FROM equipment WHERE code=?", (code,)).fetchone():
            print(f"POMIJAM (kod istnieje): {code}")
            skipped += 1
            continue

        qty_raw = val("quantity")
        try:
            qty = max(1, int(float(qty_raw)))
        except (TypeError, ValueError):
            qty = 1

        # zdjęcie: najpierw katalog, potem osadzone w arkuszu
        photo = None
        pfile = find_photo_file(photos_dir, code)
        if pfile:
            photo = None if args.dry_run else save_photo_file(pfile)
        elif r_idx in embedded:
            photo = None if args.dry_run else save_photo_bytes(embedded[r_idx])
        else:
            no_photo += 1

        material_type = "wlasny" if code.startswith("00") else "klient"

        if args.dry_run:
            print(f"DODAŁBYM: {code} | {name} | {qty} szt. | mat.: {material_type} | "
                  f"zdjęcie: {'tak' if (pfile or r_idx in embedded) else 'BRAK'}")
        else:
            con.execute(
                """INSERT INTO equipment (code, project_number, name, dimensions, photo,
                   location, warehouse_id, owner, brand, material_type, quantity)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (code, str(val("project_number") or "").strip(),
                 name, str(val("dimensions") or "").strip(), photo,
                 str(val("location") or "").strip(), warehouse_id,
                 str(val("owner") or "").strip(), str(val("brand") or "").strip(),
                 material_type, qty))
        added += 1

    if not args.dry_run:
        con.commit()
    con.close()
    print(f"\nGotowe. Dodano: {added}, pominięto (duplikaty kodu): {skipped}, bez zdjęcia: {no_photo}.")
    if args.dry_run:
        print("(dry-run – nic nie zapisano)")


if __name__ == "__main__":
    main()
