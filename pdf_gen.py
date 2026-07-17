"""Generowanie protokołów wydania / przyjęcia (PDF)."""
from io import BytesIO
from pathlib import Path
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

FONT = "Helvetica"
FONT_B = "Helvetica-Bold"
_font_dir = Path(__file__).parent / "static" / "fonts"
try:
    pdfmetrics.registerFont(TTFont("DejaVu", str(_font_dir / "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(_font_dir / "DejaVuSans-Bold.ttf")))
    FONT, FONT_B = "DejaVu", "DejaVu-Bold"
except Exception:
    pass  # brak fontu -> Helvetica (bez polskich znaków)

UPLOADS = Path(__file__).parent / "static" / "uploads"


def _wrap(text, max_chars):
    """Proste zawijanie tekstu do listy linii."""
    lines = []
    for raw in str(text).splitlines():
        raw = raw.strip()
        while len(raw) > max_chars:
            cut = raw.rfind(" ", 0, max_chars)
            cut = cut if cut > 0 else max_chars
            lines.append(raw[:cut])
            raw = raw[cut:].strip()
        lines.append(raw)
    return lines or [""]


def _get(row, key):
    try:
        return row[key]
    except (KeyError, IndexError):
        return None


def _recipient_lines(res):
    """Sekcja adresata towaru z pełnymi danymi kontaktowymi."""
    out = []
    if _get(res, "recipient_name"):
        out.append(("Adresat", res["recipient_name"]))
    if _get(res, "recipient_address"):
        out.append(("Adres dostawy", res["recipient_address"]))
    if _get(res, "recipient_contact"):
        out.append(("Osoba kontaktowa", res["recipient_contact"]))
    if _get(res, "recipient_phone"):
        out.append(("Telefon", res["recipient_phone"]))
    if _get(res, "recipient_email"):
        out.append(("E-mail", res["recipient_email"]))
    return out


def _draw_label_value(c, x, y, label, value, font_size=9, value_max_chars=42):
    """Rysuje 'Etykieta: wartość' ze spacją po dwukropku (bez nakładania)."""
    c.setFont(FONT_B, font_size)
    prefix = f"{label}: "
    c.drawString(x, y, prefix)
    vx = x + c.stringWidth(prefix, FONT_B, font_size)
    c.setFont(FONT, font_size)
    wrapped = _wrap(value, value_max_chars)
    for i, wl in enumerate(wrapped):
        c.drawString(vx if i == 0 else x + 2 * mm, y - i * 4.2 * mm, wl)
    return len(wrapped)


def _boxed_section(c, x, y, w, title, rows, font_size=9):
    """Wyraźnie wydzielona ramka z tytułem i wierszami etykieta: wartość."""
    pad = 2.5 * mm
    line_h = 4.8 * mm
    lines = []
    for label, value in rows:
        wrapped = _wrap(value, 55)
        lines.append((label, wrapped))
    n_lines = sum(len(wr) for _, wr in lines)
    box_h = pad * 2 + 5 * mm + n_lines * line_h
    top = y
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.rect(x, top - box_h, w, box_h)
    ty = top - pad - 3.5 * mm
    c.setFont(FONT_B, font_size + 1)
    c.drawString(x + pad, ty, title)
    ty -= 5.5 * mm
    for label, wrapped in lines:
        c.setFont(FONT_B, font_size)
        prefix = f"{label}: "
        c.drawString(x + pad, ty, prefix)
        vx = x + pad + c.stringWidth(prefix, FONT_B, font_size)
        c.setFont(FONT, font_size)
        for i, wl in enumerate(wrapped):
            c.drawString(vx if i == 0 else x + pad + 2 * mm, ty, wl)
            if i < len(wrapped) - 1:
                ty -= line_h
        ty -= line_h
    return top - box_h - 4 * mm


def _draw_photos(c, m, y, w, photo_names, max_h=32 * mm):
    """Rysuje do 5 zdjęć w jednym rzędzie. Zwraca nowe y."""
    names = [p for p in (photo_names or []) if p][:5]
    if not names:
        return y
    gap = 3 * mm
    n = len(names)
    cell_w = (w - 2 * m - gap * (n - 1)) / n
    max_ph = max_h
    drawn_h = 0
    images = []
    for fn in names:
        path = UPLOADS / fn
        if not path.exists():
            images.append(None)
            continue
        try:
            img = ImageReader(str(path))
            iw, ih = img.getSize()
            scale = min(cell_w / iw, max_ph / ih)
            dw, dh = iw * scale, ih * scale
            images.append((img, dw, dh))
            drawn_h = max(drawn_h, dh)
        except Exception:
            images.append(None)
    if drawn_h <= 0:
        return y
    y -= drawn_h + 2 * mm
    x = m
    for item in images:
        if item:
            img, dw, dh = item
            c.drawImage(img, x + (cell_w - dw) / 2, y, dw, dh,
                        preserveAspectRatio=True, anchor="sw")
        x += cell_w + gap
    c.setFont(FONT, 7)
    c.drawString(m, y - 3.5 * mm, "Zdjęcie sprzętu" if n == 1 else f"Zdjęcia sprzętu ({n})")
    return y - 6 * mm


def protocol_pdf(kind, res, eq, user_name, operator_name=None, photos=None):
    """kind: 'wydanie' | 'przyjecie'. Kompaktowy układ na 1 stronę A4."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    m = 14 * mm
    y = h - m

    title = "PROTOKÓŁ WYDANIA SPRZĘTU" if kind == "wydanie" else "PROTOKÓŁ PRZYJĘCIA SPRZĘTU"
    doc_no = f"{'WZ' if kind == 'wydanie' else 'PZ'}/{res['id']}/{datetime.now():%Y}"

    c.setFont(FONT_B, 13)
    c.drawString(m, y, title)
    c.setFont(FONT, 8)
    c.drawRightString(w - m, y, f"Nr dokumentu: {doc_no}")
    y -= 4.5 * mm
    c.drawRightString(w - m, y, f"Data wygenerowania: {datetime.now():%Y-%m-%d %H:%M}")
    y -= 3 * mm
    c.setStrokeColor(colors.black)
    c.line(m, y, w - m, y)
    y -= 6 * mm

    wh = _get(eq, "warehouse_name")
    wh_addr = _get(eq, "warehouse_address")
    wh_txt = "-"
    if wh:
        wh_txt = wh + (f" ({wh_addr})" if wh_addr else "")

    left = [
        ("Kod sprzętu", eq["code"]),
        ("Nazwa", eq["name"]),
        ("Numer projektu", eq["project_number"] or "-"),
        ("Wymiary", eq["dimensions"] or "-"),
        ("Magazyn" if kind == "wydanie" else "Magazyn przyjęcia", wh_txt),
        ("Miejsce w magazynie", eq["location"] or "-"),
        ("Własność", eq["owner"] or "-"),
        ("Brand", _get(eq, "brand") or "-"),
    ]
    right = [
        ("Ilość sztuk", str(res["quantity"])),
        ("Termin", f"{res['date_from']} – {res['date_to']}"),
        ("Klient / cel", res["client"] or "-"),
        ("Odbiera towar", res["receiver"] or "-"),
        ("Rezerwujący", user_name),
    ]
    if operator_name:
        right.append(("Obsługa magazynu", operator_name))
    if kind == "przyjecie":
        damaged = bool(_get(res, "damage"))
        right.append(("Stan przy zwrocie", "uszkodzony" if damaged else "sprawny"))
        if damaged and _get(res, "damage_notes"):
            right.append(("Opis uszkodzenia", res["damage_notes"]))
        if _get(res, "returned_at"):
            right.append(("Data zwrotu", str(res["returned_at"])[:16].replace("T", " ")))

    col_gap = 6 * mm
    col_w = (w - 2 * m - col_gap) / 2
    fs = 9
    row_h = 5 * mm
    y0 = y
    yl, yr = y0, y0
    for label, value in left:
        n = _draw_label_value(c, m, yl, label, value, fs, 38)
        yl -= max(1, n) * row_h
    for label, value in right:
        n = _draw_label_value(c, m + col_w + col_gap, yr, label, value, fs, 38)
        yr -= max(1, n) * row_h
    y = min(yl, yr) - 2 * mm

    if res["notes"]:
        n = _draw_label_value(c, m, y, "Uwagi", res["notes"], fs, 90)
        y -= max(1, n) * row_h + 1 * mm

    if _get(eq, "storage_instructions"):
        n = _draw_label_value(c, m, y, "Pakowanie / transport",
                              eq["storage_instructions"], fs, 80)
        y -= max(1, n) * row_h + 1 * mm

    rec = _recipient_lines(res)
    if rec:
        y = _boxed_section(c, m, y, w - 2 * m, "ADRESAT TOWARU (dostawa)", rec, font_size=8.5)

    photo_list = list(photos or [])
    if not photo_list and eq["photo"]:
        photo_list = [eq["photo"]]
    y = _draw_photos(c, m, y, w, photo_list, max_h=28 * mm)

    # podpisy – zawsze na dole strony
    sig_y = 28 * mm
    if y < sig_y + 8 * mm:
        # brak miejsca – zmniejsz zdjęcia nie da się już, podpisy i tak na dole
        pass
    c.setFont(FONT, 9)
    c.line(m, sig_y, m + 55 * mm, sig_y)
    c.line(w - m - 55 * mm, sig_y, w - m, sig_y)
    c.drawString(m + 6 * mm, sig_y - 4.5 * mm, "Wydający / Przyjmujący")
    c.drawString(w - m - 48 * mm, sig_y - 4.5 * mm, "Odbierający / Zwracający")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


def group_pdf(kind, rows):
    """Zbiorczy protokół dla wielu rezerwacji. rows: sqlite3.Row z JOIN equipment+users."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    m = 18 * mm

    title = ("ZBIORCZY PROTOKÓŁ WYDANIA SPRZĘTU" if kind == "wydanie"
             else "ZBIORCZY PROTOKÓŁ PRZYJĘCIA SPRZĘTU")
    doc_no = f"{'WZ' if kind == 'wydanie' else 'PZ'}-ZB/{datetime.now():%Y%m%d%H%M}"

    wh_set = {( _get(r, "warehouse_name"), _get(r, "warehouse_address"))
              for r in rows if _get(r, "warehouse_name")}
    single_wh = list(wh_set)[0] if len(wh_set) == 1 else None
    multi_wh = len(wh_set) > 1

    def header():
        y = h - m
        c.setFont(FONT_B, 14)
        c.drawString(m, y, title)
        c.setFont(FONT, 9)
        c.drawRightString(w - m, y + 1, f"Nr: {doc_no}")
        y -= 5 * mm
        c.drawRightString(w - m, y, f"Data: {datetime.now():%Y-%m-%d %H:%M}")
        if single_wh:
            c.setFont(FONT_B, 9)
            label = "Magazyn (przyjęcie)" if kind == "przyjecie" else "Magazyn (odbiór)"
            wh_line = f"{label}: {single_wh[0]}"
            if single_wh[1]:
                wh_line += f", {single_wh[1]}"
            c.drawString(m, y, wh_line)
            c.setFont(FONT, 9)
        elif multi_wh:
            c.setFont(FONT_B, 9)
            c.setFillColor(colors.red)
            msg = ("UWAGA: zwrot do więcej niż jednego magazynu – szczegóły przy pozycjach"
                   if kind == "przyjecie"
                   else "UWAGA: pozycje z więcej niż jednego magazynu – szczegóły przy pozycjach")
            c.drawString(m, y, msg)
            c.setFillColor(colors.black)
            c.setFont(FONT, 9)
        y -= 3 * mm
        c.line(m, y, w - m, y)
        return y - 8 * mm

    y = header()

    terms = sorted({(r["date_from"], r["date_to"]) for r in rows})
    clients = sorted({r["client"] for r in rows if r["client"]})
    users = sorted({f"{(_get(r,'first_name') or '').strip()} {(_get(r,'last_name') or '').strip()}".strip()
                    or r["username"] for r in rows})
    receivers = sorted({r["receiver"] for r in rows if r["receiver"]})
    common = [
        ("Termin", ", ".join(f"{a} – {b}" for a, b in terms)),
        ("Klient / cel", ", ".join(clients) or "-"),
        ("Odbiera towar", ", ".join(receivers) or "-"),
        ("Rezerwujący", ", ".join(users)),
        ("Liczba pozycji", str(len(rows))),
    ]
    c.setFont(FONT, 10)
    for label, value in common:
        c.setFont(FONT_B, 10)
        prefix = f"{label}: "
        c.drawString(m, y, prefix)
        c.setFont(FONT, 10)
        c.drawString(m + c.stringWidth(prefix, FONT_B, 10), y, value)
        y -= 6 * mm

    notes = sorted({r["notes"].strip() for r in rows if (r["notes"] or "").strip()})
    if notes:
        c.setFont(FONT_B, 10)
        prefix = "Uwagi: "
        c.drawString(m, y, prefix)
        c.setFont(FONT, 10)
        vx = m + c.stringWidth(prefix, FONT_B, 10)
        for n in notes:
            for wl in _wrap(n, 75):
                c.drawString(vx, y, wl)
                y -= 5 * mm
                vx = m + 2 * mm
        y -= 1 * mm
    y -= 2 * mm

    rec_rows = []
    for r in rows:
        rec_rows = _recipient_lines(r)
        if rec_rows:
            break
    if rec_rows:
        y = _boxed_section(c, m, y, w - 2 * m, "ADRESAT TOWARU (dostawa)", rec_rows, font_size=9)

    row_h = 26 * mm
    col_x = [m, m + 24 * mm, m + 46 * mm, m + 94 * mm, m + 128 * mm, m + 139 * mm, w - m]

    def table_head(y):
        c.setFont(FONT_B, 9)
        for x, t in zip(col_x, ["Zdjęcie", "Kod", "Nazwa", "Magazyn / miejsce", "Szt.", "Własność / brand"]):
            c.drawString(x + 1 * mm, y, t)
        y -= 2 * mm
        c.line(m, y, w - m, y)
        return y

    y = table_head(y)
    for r in rows:
        if y - row_h < 35 * mm:
            c.showPage()
            y = header()
            y = table_head(y)
        yr = y - row_h
        if r["photo"]:
            p = UPLOADS / r["photo"]
            if p.exists():
                try:
                    img = ImageReader(str(p))
                    iw, ih = img.getSize()
                    s = min(22 * mm / iw, (row_h - 4 * mm) / ih)
                    c.drawImage(img, col_x[0], yr + 2 * mm, iw * s, ih * s,
                                preserveAspectRatio=True, anchor="sw")
                except Exception:
                    pass
        c.setFont(FONT, 9)
        ty = y - 6 * mm
        c.drawString(col_x[1] + 1 * mm, ty, r["code"])
        name = r["name"]
        max_chars = 22
        c.drawString(col_x[2] + 1 * mm, ty, name[:max_chars])
        if len(name) > max_chars:
            c.drawString(col_x[2] + 1 * mm, ty - 4 * mm, name[max_chars:max_chars * 2])
        wh_name = _get(r, "warehouse_name") or "-"
        c.drawString(col_x[3] + 1 * mm, ty, wh_name[:18])
        c.drawString(col_x[3] + 1 * mm, ty - 4 * mm, (r["location"] or "-")[:18])
        c.drawString(col_x[4] + 1 * mm, ty, str(r["quantity"]))
        c.drawString(col_x[5] + 1 * mm, ty, (r["owner"] or "-")[:16])
        if _get(r, "brand"):
            c.drawString(col_x[5] + 1 * mm, ty - 4 * mm, r["brand"][:16])
        extra_y = ty - 9 * mm
        if kind == "przyjecie" and _get(r, "damage"):
            c.setFont(FONT_B, 7.5)
            c.setFillColor(colors.HexColor("#b42318"))
            note = "USZKODZONY"
            if _get(r, "damage_notes"):
                note += ": " + str(r["damage_notes"])
            for wl in _wrap(note, 90)[:2]:
                c.drawString(col_x[1] + 1 * mm, extra_y, wl)
                extra_y -= 3.5 * mm
            c.setFillColor(colors.black)
            c.setFont(FONT, 9)
        if _get(r, "storage_instructions"):
            c.setFont(FONT, 7.5)
            c.setFillColor(colors.HexColor("#444444"))
            si = _wrap("Pakowanie: " + r["storage_instructions"], 90)[:2]
            for wl in si:
                c.drawString(col_x[1] + 1 * mm, extra_y, wl)
                extra_y -= 3.5 * mm
            c.setFillColor(colors.black)
            c.setFont(FONT, 9)
        y = yr
        c.setStrokeColor(colors.grey)
        c.line(m, y, w - m, y)
        c.setStrokeColor(colors.black)

    y = max(y - 18 * mm, 25 * mm)
    c.setFont(FONT, 10)
    c.line(m, y, m + 60 * mm, y)
    c.line(w - m - 60 * mm, y, w - m, y)
    c.drawString(m + 10 * mm, y - 5 * mm, "Wydający / Przyjmujący")
    c.drawString(w - m - 50 * mm, y - 5 * mm, "Odbierający / Zwracający")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf
