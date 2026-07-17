"""Dodaje przykładowe produkty (uruchom raz: python seed.py)."""
from db import get_db, init_db
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

UPLOADS = Path(__file__).parent / "static" / "uploads"

PRODUCTS = [
    # code, name, qty, dimensions, location, owner, project, kolor placeholdera
    ("CC0039", "CZERWONA BECZKA CC", 8, "wysokość ok. 85 cm, średnica ok. 60 cm", "A1", "COCA-COLA", "4997", (196, 30, 30)),
    ("CC0041", "USZKODZONA SKRZYNIA NA LISTY", 1, "", "A2", "COCA-COLA", "4997", (170, 40, 40)),
    ("CC0043", "KOSTKA KINLEY", 5, "50X50X50", "AR1P2", "KINLEY", "4997", (230, 190, 60)),
]


def placeholder(code, color):
    """Prosty placeholder do podmiany na prawdziwe zdjęcie (Edytuj -> Zdjęcie)."""
    UPLOADS.mkdir(parents=True, exist_ok=True)
    fname = f"seed_{code}.png"
    img = Image.new("RGB", (400, 300), color)
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(str(Path(__file__).parent / "static" / "fonts" / "DejaVuSans-Bold.ttf"), 36)
    except Exception:
        font = ImageFont.load_default()
    d.text((200, 150), code, fill="white", font=font, anchor="mm")
    img.save(UPLOADS / fname)
    return fname


def main():
    init_db()
    con = get_db()
    added = 0
    for code, name, qty, dims, loc, owner, project, color in PRODUCTS:
        if con.execute("SELECT 1 FROM equipment WHERE code=?", (code,)).fetchone():
            print(f"{code}: już istnieje, pomijam")
            continue
        con.execute(
            """INSERT INTO equipment (code, project_number, name, dimensions,
               photo, location, owner, quantity, notes) VALUES (?,?,?,?,?,?,?,?,?)""",
            (code, project, name, dims, placeholder(code, color), loc, owner, qty, ""))
        added += 1
        print(f"{code}: dodano")
    con.commit()
    con.close()
    print(f"Gotowe – dodano {added} produkt(y).")


if __name__ == "__main__":
    main()
