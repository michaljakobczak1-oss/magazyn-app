"""Generowanie protokołów wydania / przyjęcia (PDF)."""
import os
from io import BytesIO
from pathlib import Path
from datetime import datetime

from db import local_now

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

_BASE = Path(__file__).parent


def _upload_dirs():
    """Katalogi ze zdjęciami: dysk trwały (DATA_DIR) + static/uploads."""
    dirs = []
    data = os.environ.get("DATA_DIR")
    if data:
        dirs.append(Path(data) / "uploads")
    dirs.append(_BASE / "data" / "uploads")
    dirs.append(_BASE / "static" / "uploads")
    return dirs


def _resolve_photo(fn):
    if not fn:
        return None
    for d in _upload_dirs():
        path = d / fn
        if path.exists():
            return path
    return None


# kompatybilność wsteczna
UPLOADS = _BASE / "static" / "uploads"


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


def _wrap_width(c, text, font, size, max_w):
    """Zawijanie tekstu wg szerokości kolumny (mm → pt w ReportLab)."""
    words = str(text or "").split()
    if not words:
        return [""]
    lines, cur = [], words[0]
    for w in words[1:]:
        trial = cur + " " + w
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    lines.append(cur)
    return lines


def _draw_col_lines(c, x, y, lines, font, size, line_h=3.5 * mm):
    """Rysuje linie tekstu w kolumnie, zwraca nowe y."""
    c.setFont(font, size)
    for i, line in enumerate(lines):
        c.drawString(x, y - i * line_h, line)
    return y - max(1, len(lines)) * line_h


def _get(row, key):
    try:
        return row[key]
    except (KeyError, IndexError):
        return None


def _fmt_ts(val):
    """2026-07-20T16:45:00 → '2026-07-20 16:45' (jak w tabeli)."""
    if not val:
        return None
    return str(val)[:16].replace("T", " ")


def _fmt_day(val):
    if not val:
        return None
    return str(val)[:10]


def _actual_period(r, kind):
    """Rzeczywisty okres jak w statusach: wydanie → zwrot/utylizacja."""
    start = _fmt_day(_get(r, "issued_at")) or r["date_from"]
    if kind == "przyjecie":
        end = _fmt_day(_get(r, "returned_at")) or r["date_to"]
    else:
        end = r["date_to"]
    return start, end


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
        path = _resolve_photo(fn)
        if not path:
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


def _draw_protocol_page(c, kind, res, eq, user_name, operator_name=None, photos=None):
    """Rysuje jedną stronę protokołu (WZ/PZ) na istniejącym canvasie."""
    w, h = A4
    m = 14 * mm
    y = h - m

    title = "PROTOKÓŁ WYDANIA SPRZĘTU" if kind == "wydanie" else "PROTOKÓŁ PRZYJĘCIA SPRZĘTU"
    if kind == "przyjecie" and res["status"] == "utylizacja":
        title = "PROTOKÓŁ UTYLIZACJI SPRZĘTU"
    doc_no = f"{'WZ' if kind == 'wydanie' else 'PZ'}/{res['id']}/{local_now():%Y}"

    c.setFont(FONT_B, 13)
    c.drawString(m, y, title)
    c.setFont(FONT, 8)
    c.drawRightString(w - m, y, f"Nr dokumentu: {doc_no}")
    y -= 4.5 * mm
    c.drawRightString(w - m, y, f"Data wygenerowania: {local_now():%Y-%m-%d %H:%M}")
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
        ("Termin planowany", f"{res['date_from']} – {res['date_to']}"),
        ("Klient / cel", res["client"] or "-"),
        ("Odbiera towar", res["receiver"] or "-"),
        ("Rezerwujący", user_name),
    ]
    if operator_name:
        right.append(("Obsługa magazynu", operator_name))
    issued = _fmt_ts(_get(res, "issued_at"))
    if issued:
        right.append(("Data wydania", issued))
    if kind == "wydanie":
        start, end = _actual_period(res, kind)
        right[1] = ("Termin", f"{start} – {end}")
    if kind == "przyjecie":
        start, end = _actual_period(res, kind)
        right[1] = ("Termin (wydanie – zwrot)", f"{start} – {end}")
        if res["status"] == "utylizacja":
            right.append(("Rozstrzygnięcie", "UTYLIZACJA – towar nie wraca"))
            if _get(res, "damage_notes"):
                right.append(("Powód", res["damage_notes"]))
            ret = _fmt_ts(_get(res, "returned_at"))
            if ret:
                right.append(("Data utylizacji", ret))
        else:
            damaged = bool(_get(res, "damage"))
            right.append(("Stan przy zwrocie", "uszkodzony" if damaged else "sprawny"))
            if damaged and _get(res, "damage_notes"):
                right.append(("Opis uszkodzenia", res["damage_notes"]))
            ret = _fmt_ts(_get(res, "returned_at"))
            if ret:
                right.append(("Data zwrotu", ret))

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

    if kind == "wydanie" and (_get(res, "permanent") or res["status"] == "wydane trwale"):
        n = _draw_label_value(c, m, y, "Typ wydania",
                              "Wydanie trwałe – towar nie wraca do magazynu", fs, 90)
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
    c.setFont(FONT, 9)
    c.line(m, sig_y, m + 55 * mm, sig_y)
    c.line(w - m - 55 * mm, sig_y, w - m, sig_y)
    c.drawString(m + 6 * mm, sig_y - 4.5 * mm, "Wydający / Przyjmujący")
    c.drawString(w - m - 48 * mm, sig_y - 4.5 * mm, "Odbierający / Zwracający")


def protocol_pdf(kind, res, eq, user_name, operator_name=None, photos=None):
    """kind: 'wydanie' | 'przyjecie'. Kompaktowy układ na 1 stronę A4."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    _draw_protocol_page(c, kind, res, eq, user_name, operator_name, photos)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf


def protocols_pdf(kind, pages):
    """Wiele pozycji – każda na osobnej stronie w tym samym układzie co pojedynczy WZ/PZ.

    pages: lista dictów {res, eq, user_name, operator_name, photos}
    """
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    for p in pages:
        _draw_protocol_page(
            c, kind, p["res"], p["eq"], p["user_name"],
            p.get("operator_name"), p.get("photos"))
        c.showPage()
    c.save()
    buf.seek(0)
    return buf


def group_pdf(kind, rows):
    """Zbiorczy protokół tabelaryczny (pozostawiony dla kompatybilności / jawnego eksportu)."""
    # Zdjęcia z dysku trwałego – _resolve_photo
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    m = 18 * mm

    title = ("ZBIORCZY PROTOKÓŁ WYDANIA SPRZĘTU" if kind == "wydanie"
             else "ZBIORCZY PROTOKÓŁ PRZYJĘCIA / UTYLIZACJI")
    if kind == "przyjecie":
        n_disp = sum(1 for r in rows if r["status"] == "utylizacja")
        n_ret = len(rows) - n_disp
        if n_disp and not n_ret:
            title = "ZBIORCZY PROTOKÓŁ UTYLIZACJI SPRZĘTU"
        elif n_disp and n_ret:
            title = "ZBIORCZY PROTOKÓŁ PRZYJĘCIA I UTYLIZACJI"
    doc_no = f"{'WZ' if kind == 'wydanie' else 'PZ'}-ZB/{local_now():%Y%m%d%H%M}"

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
        c.drawRightString(w - m, y, f"Data: {local_now():%Y-%m-%d %H:%M}")
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

    terms = sorted({_actual_period(r, kind) for r in rows})
    clients = sorted({r["client"] for r in rows if r["client"]})
    users = sorted({f"{(_get(r,'first_name') or '').strip()} {(_get(r,'last_name') or '').strip()}".strip()
                    or r["username"] for r in rows})
    receivers = sorted({r["receiver"] for r in rows if r["receiver"]})
    if len(terms) == 1:
        termin_txt = f"{terms[0][0]} – {terms[0][1]}"
    else:
        termin_txt = "różne – szczegóły przy pozycjach"
    termin_label = "Termin (wydanie – zwrot)" if kind == "przyjecie" else "Termin"
    common = [
        (termin_label, termin_txt),
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

    row_h = 28 * mm
    # Zdjęcie | Kod | Nazwa | Termin | Magazyn | Szt. | Własność  (szerokości dopasowane do długich kodów)
    tbl_left = m
    tbl_right = w - m
    col_w = [22 * mm, 28 * mm, 38 * mm, 34 * mm, 30 * mm, 10 * mm,
             tbl_right - tbl_left - (22 + 28 + 38 + 34 + 30 + 10) * mm]
    col_x = [tbl_left]
    for cw in col_w[:-1]:
        col_x.append(col_x[-1] + cw)

    def table_head(y):
        c.setFont(FONT_B, 8)
        headers = ["Zdjęcie", "Kod", "Nazwa", "Termin", "Magazyn / miejsce", "Szt.", "Własność / brand"]
        for x, cw, t in zip(col_x, col_w, headers):
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
            p = _resolve_photo(r["photo"])
            if p:
                try:
                    img = ImageReader(str(p))
                    iw, ih = img.getSize()
                    s = min((col_w[0] - 2 * mm) / iw, (row_h - 4 * mm) / ih)
                    c.drawImage(img, col_x[0] + 1 * mm, yr + 2 * mm, iw * s, ih * s,
                                preserveAspectRatio=True, anchor="sw")
                except Exception:
                    pass
        ty = y - 5 * mm
        pad = 1 * mm
        # kod – dopasuj czcionkę, żeby nie wchodził w nazwę
        code = str(r["code"])
        code_size = 8
        while code_size >= 6.5 and c.stringWidth(code, FONT, code_size) > col_w[1] - 2 * pad:
            code_size -= 0.5
        c.setFont(FONT, code_size)
        c.drawString(col_x[1] + pad, ty, code)
        # nazwa – zawijanie po słowach w ramach kolumny
        name_lines = _wrap_width(c, r["name"], FONT, 8, col_w[2] - 2 * pad)[:2]
        _draw_col_lines(c, col_x[2] + pad, ty, name_lines, FONT, 8)
        # termin
        c.setFont(FONT_B, 7.5)
        start, end = _actual_period(r, kind)
        issued_ts = _fmt_ts(_get(r, "issued_at"))
        returned_ts = _fmt_ts(_get(r, "returned_at"))
        term_x = col_x[3] + pad
        term_w = col_w[3] - 2 * pad
        if kind == "przyjecie" and (issued_ts or returned_ts):
            line1 = ("wyd. " + issued_ts) if issued_ts else start
            if r["status"] == "utylizacja" and returned_ts:
                line2 = "utyl. " + returned_ts
            elif returned_ts:
                line2 = "zwr. " + returned_ts
            else:
                line2 = "– " + end
            c.drawString(term_x, ty, line1[:28])
            c.setFont(FONT, 7.5)
            c.drawString(term_x, ty - 3.5 * mm, line2[:28])
        elif kind == "wydanie" and issued_ts:
            c.drawString(term_x, ty, start[:28])
            c.setFont(FONT, 7.5)
            c.drawString(term_x, ty - 3.5 * mm, "– " + str(end)[:26])
        else:
            c.drawString(term_x, ty, str(start)[:28])
            c.setFont(FONT, 7.5)
            c.drawString(term_x, ty - 3.5 * mm, "– " + str(end)[:26])
        if _get(r, "client"):
            c.setFillColor(colors.HexColor("#444444"))
            for i, wl in enumerate(_wrap_width(c, r["client"], FONT, 7, term_w)[:1]):
                c.drawString(term_x, ty - 7 * mm - i * 3.2 * mm, wl)
            c.setFillColor(colors.black)
        c.setFont(FONT, 8)
        wh_name = _get(r, "warehouse_name") or "-"
        wh_lines = _wrap_width(c, wh_name, FONT, 8, col_w[4] - 2 * pad)[:1]
        loc_lines = _wrap_width(c, r["location"] or "-", FONT, 8, col_w[4] - 2 * pad)[:1]
        _draw_col_lines(c, col_x[4] + pad, ty, wh_lines, FONT, 8)
        if loc_lines:
            c.setFont(FONT, 8)
            c.drawString(col_x[4] + pad, ty - 3.5 * mm, loc_lines[0])
        c.drawString(col_x[5] + pad, ty, str(r["quantity"]))
        owner_lines = _wrap_width(c, r["owner"] or "-", FONT, 8, col_w[6] - 2 * pad)[:1]
        _draw_col_lines(c, col_x[6] + pad, ty, owner_lines, FONT, 8)
        if _get(r, "brand"):
            brand_lines = _wrap_width(c, r["brand"], FONT, 8, col_w[6] - 2 * pad)[:1]
            c.setFont(FONT, 8)
            c.drawString(col_x[6] + pad, ty - 3.5 * mm, brand_lines[0])
        extra_y = ty - 11 * mm
        if kind == "przyjecie":
            c.setFont(FONT_B, 7)
            if r["status"] == "utylizacja":
                c.setFillColor(colors.HexColor("#6b21a8"))
                note = "UTYLIZACJA / NIE WRACA"
                if _get(r, "damage_notes"):
                    note += ": " + str(r["damage_notes"])
            elif _get(r, "damage"):
                c.setFillColor(colors.HexColor("#b42318"))
                note = "USZKODZONY (zwrot)"
                if _get(r, "damage_notes"):
                    note += ": " + str(r["damage_notes"])
            else:
                c.setFillColor(colors.HexColor("#166534"))
                note = "ZWROT NA MAGAZYN"
            for wl in _wrap(note, 85)[:3]:
                c.drawString(col_x[1] + 0.5 * mm, extra_y, wl)
                extra_y -= 3.2 * mm
            c.setFillColor(colors.black)
            c.setFont(FONT, 8)
        if _get(r, "storage_instructions"):
            c.setFont(FONT, 7)
            c.setFillColor(colors.HexColor("#444444"))
            si = _wrap("Pakowanie: " + r["storage_instructions"], 85)[:2]
            for wl in si:
                c.drawString(col_x[1] + 0.5 * mm, extra_y, wl)
                extra_y -= 3.2 * mm
            c.setFillColor(colors.black)
            c.setFont(FONT, 8)
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
