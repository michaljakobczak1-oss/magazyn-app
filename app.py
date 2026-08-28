import os
import re
import uuid
from datetime import datetime, date, timedelta
from functools import wraps
from pathlib import Path
from urllib.parse import urlencode

from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, send_file, abort, jsonify, after_this_request)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

from db import (get_db, init_db, reserved_qty, display_name, upsert_recipient,
                equipment_photo_list, equipment_photo_rows, equipment_photo_kind_map,
                local_now, local_today, handoff_conflict,
                next_free_after_return, replace_equipment_stock,
                move_equipment_stock_on_return, take_equipment_stock,
                add_equipment_stock, sync_equipment_from_stock,
                sync_equipment_primary_photo, DB_PATH)
from pdf_gen import protocol_pdf, group_pdf
from import_excel import run_import
from export_excel import (
    build_catalog_miniatures_xlsx,
    build_catalog_import_xlsx,
    export_photo_zip_names,
)
from xbs_awizacja import (
    build_xbs_awizacja_xlsx, xbs_filename, MATERIAL_OPTIONS,
)
import tempfile
import shutil
import zipfile
import json
import threading
from io import BytesIO

BASE = Path(__file__).parent
# Zdjęcia: przy DATA_DIR (dysk trwały) trzymaj w DATA_DIR/uploads
_DATA_DIR = Path(os.environ.get("DATA_DIR") or (BASE / "data"))
_DATA_DIR.mkdir(parents=True, exist_ok=True)
_UPLOAD_PERSIST = _DATA_DIR / "uploads"
_UPLOAD_PERSIST.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR = BASE / "static" / "uploads"
if os.environ.get("DATA_DIR"):
    UPLOAD_DIR = _UPLOAD_PERSIST
    # serwuj /static/uploads/* z dysku trwałego
ALLOWED_EXT = {"png", "jpg", "jpeg", "gif", "webp"}
MAX_PHOTOS = 5

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "zmien-mnie-w-produkcji")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # import katalogu + zdjęcia (zip)
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

init_db()
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

if os.environ.get("DATA_DIR"):
    @app.route("/static/uploads/<path:filename>")
    def persistent_upload(filename):
        from flask import send_from_directory
        return send_from_directory(_UPLOAD_PERSIST, filename)


# ---------- pomocnicze ----------

def login_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if "user_id" not in session:
            return redirect(url_for("login", next=request.path))
        return f(*a, **kw)
    return wrapper


def admin_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if session.get("role") != "admin":
            flash("Wymagane uprawnienia administratora.", "error")
            return redirect(url_for("index"))
        return f(*a, **kw)
    return wrapper


TCL_DEPT_NAME = "Warrens"


def equipment_catalog(eq):
    """main | tcl"""
    try:
        return ((eq["catalog"] or "main").strip().lower() or "main")
    except (KeyError, IndexError, TypeError):
        return "main"


def can_manage_tcl():
    """Admin albo dział Warrens – zarządzanie katalogiem TCL."""
    if session.get("role") == "admin":
        return True
    dept = (session.get("department") or "").strip()
    return dept.lower() == TCL_DEPT_NAME.lower()


def tcl_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if not can_manage_tcl():
            flash("Brak dostępu do modułu TCL (wymagany dział Warrens lub admin).", "error")
            return redirect(url_for("dashboard"))
        return f(*a, **kw)
    return wrapper


@app.context_processor
def inject_acl_flags():
    return {"can_manage_tcl": can_manage_tcl()}


def _normalize_brand_key(brand):
    """Ujednolica markę do porównań (COCA -COLA → COCA-COLA)."""
    s = (brand or "").strip().upper()
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\s+", " ", s)
    return s


def _canonical_brand_label(brand):
    """Etykieta w filtrze bez spacji wokół myślnika."""
    return re.sub(r"\s*-\s*", "-", (brand or "").strip())


def _dedupe_brand_labels(raw_brands):
    """Jedna pozycja na wariant marki – preferuj zapis bez spacji przy myślniku."""
    by_key = {}
    for b in raw_brands:
        if not b:
            continue
        key = _normalize_brand_key(b)
        label = _canonical_brand_label(b)
        prev = by_key.get(key)
        if prev is None:
            by_key[key] = label
        elif (" -" in prev or "- " in prev) and (" -" not in b and "- " not in b):
            by_key[key] = label
    return sorted(by_key.values(), key=lambda x: x.upper())


def _photos_for_pdf(con, equipment_id, fallback_photo=None):
    """Zdjęcia do WZ/PZ – bez naprawionych uszkodzeń."""
    photos = [
        p for p in equipment_photo_rows(con, equipment_id)
        if (p["kind"] or "normal") != "repaired"
    ]
    if photos:
        return photos
    if fallback_photo:
        kinds = equipment_photo_kind_map(con, equipment_id)
        if kinds.get(fallback_photo) != "repaired":
            return [{"filename": fallback_photo, "kind": "normal"}]
    return []


def can_manage_reservation(r):
    """Admin, własne rezerwacje, albo ten sam dział co rezerwujący."""
    if session.get("role") == "admin":
        return True
    if r["user_id"] == session.get("user_id"):
        return True
    my_dept = (session.get("department") or "").strip()
    if not my_dept:
        return False
    owner_dept = ""
    if "owner_department" in r.keys():
        owner_dept = (r["owner_department"] or "").strip()
    return bool(owner_dept) and owner_dept == my_dept


def active_departments(con):
    return con.execute(
        "SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall()


def save_photo(file):
    if not file or not file.filename:
        return None
    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ALLOWED_EXT:
        flash("Niedozwolony format zdjęcia (png/jpg/jpeg/gif/webp).", "error")
        return None
    fname = f"{uuid.uuid4().hex}.{ext}"
    file.save(UPLOAD_DIR / secure_filename(fname))
    return fname


def valid_dates(d_from, d_to):
    try:
        f, t = date.fromisoformat(d_from), date.fromisoformat(d_to)
        return f <= t
    except (ValueError, TypeError):
        return False


def valid_pickup_date(d_from):
    try:
        date.fromisoformat(d_from)
        return True
    except (ValueError, TypeError):
        return False


SELF_PICKUP_VALUE = "__self__"


def _session_person_name():
    return (session.get("full_name") or session.get("username") or "").strip()


def resolve_receiver(raw):
    """Odbiór własny → zapis z imieniem zalogowanego użytkownika."""
    v = (raw or "").strip()
    if v == SELF_PICKUP_VALUE or v.lower() in ("odbiór własny", "odbior wlasny"):
        name = _session_person_name()
        return f"Odbiór własny ({name})" if name else "Odbiór własny"
    return v


def resolve_returner(raw):
    """Zwrot własny → zapis z imieniem zalogowanego użytkownika."""
    v = (raw or "").strip()
    low = v.lower()
    if (v == SELF_PICKUP_VALUE
            or low in ("zwrot własny", "zwrot wlasny")
            or low.startswith("zwrot własny")
            or low.startswith("zwrot wlasny")
            or low.startswith("odbiór własny")
            or low.startswith("odbior wlasny")):
        name = _session_person_name()
        return f"Zwrot własny ({name})" if name else "Zwrot własny"
    return v


def reservation_dates_from_form(form):
    """Przy wydaniu trwałym wymagana jest data Od (bez terminu zwrotu)."""
    permanent = 1 if form.get("permanent") else 0
    d_from = (form.get("date_from") or "").strip()
    d_to = (form.get("date_to") or "").strip()
    if permanent:
        return d_from, d_from, permanent
    return d_from, d_to, permanent


def active_partners(con):
    return con.execute(
        "SELECT * FROM logistics_partners WHERE active=1 ORDER BY name").fetchall()


def active_warehouses(con):
    return con.execute(
        "SELECT * FROM warehouses WHERE active=1 ORDER BY name").fetchall()


def recent_recipients(con, limit=30):
    return con.execute(
        "SELECT * FROM recipients ORDER BY last_used DESC LIMIT ?", (limit,)).fetchall()


def project_suggestions(con):
    """Numery projektu do wyboru przy rezerwacji – tylko wartości z cyframi
    (bez nazw typu brand/opis). Filtry na liście sprzętu bez zmian.
    """
    rows = con.execute(
        """SELECT DISTINCT project_number AS p FROM equipment
           WHERE IFNULL(project_number,'')!=''
           UNION
           SELECT DISTINCT project_number AS p FROM reservations
           WHERE IFNULL(project_number,'')!=''
           ORDER BY 1"""
    ).fetchall()
    return [r["p"] for r in rows if any(ch.isdigit() for ch in (r["p"] or ""))]


def recipient_form_fields(form):
    return dict(
        recipient_name=form.get("recipient_name", "").strip(),
        recipient_contact=form.get("recipient_contact", "").strip(),
        recipient_phone=form.get("recipient_phone", "").strip(),
        recipient_address=form.get("recipient_address", "").strip(),
        recipient_city=form.get("recipient_city", "").strip(),
        recipient_email=form.get("recipient_email", "").strip(),
    )


def recipient_required_error(rec):
    """None jeśli wymagane pola adresata są uzupełnione; inaczej komunikat błędu."""
    missing = []
    if not rec.get("recipient_name"):
        missing.append("firma / miejsce / osoba")
    if not rec.get("recipient_contact"):
        missing.append("osoba kontaktowa")
    if not rec.get("recipient_phone"):
        missing.append("telefon")
    if not rec.get("recipient_address"):
        missing.append("adres dostawy")
    if not rec.get("recipient_city"):
        missing.append("miasto")
    if not missing:
        return None
    return "Uzupełnij wymagane pola adresata: " + ", ".join(missing) + "."


def password_policy_error(pw):
    """None jeśli hasło OK; inaczej komunikat błędu.
    Wymagania: min. 10 znaków, co najmniej 1 litera i 1 cyfra."""
    pw = pw or ""
    if len(pw) < 10:
        return "Hasło musi mieć min. 10 znaków."
    if not any(c.isalpha() for c in pw):
        return "Hasło musi zawierać co najmniej jedną literę."
    if not any(c.isdigit() for c in pw):
        return "Hasło musi zawierać co najmniej jedną cyfrę."
    weak = {"admin123", "password", "haslo12345", "1234567890", "qwerty1234"}
    if pw.lower() in weak:
        return "Hasło jest zbyt oczywiste – wybierz inne."
    return None


# ---------- logowanie / konto ----------

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        con = get_db()
        u = con.execute("SELECT * FROM users WHERE username=? AND active=1",
                        (request.form["username"].strip(),)).fetchone()
        con.close()
        if u and check_password_hash(u["password_hash"], request.form["password"]):
            session.update(user_id=u["id"], username=u["username"], role=u["role"],
                           full_name=display_name(u),
                           department=(u["department"] or "") if "department" in u.keys() else "")
            return redirect(request.args.get("next") or url_for("dashboard"))
        flash("Błędny login lub hasło.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/account/password", methods=["POST"])
@login_required
def account_password():
    """Zmiana własnego hasła przez użytkownika."""
    con = get_db()
    u = con.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
    new_pw = request.form.get("new_password", "")
    pw_err = password_policy_error(new_pw)
    if not check_password_hash(u["password_hash"], request.form.get("current_password", "")):
        flash("Obecne hasło jest nieprawidłowe.", "error")
    elif pw_err:
        flash(pw_err, "error")
    else:
        con.execute("UPDATE users SET password_hash=? WHERE id=?",
                    (generate_password_hash(new_pw, method="pbkdf2:sha256"), u["id"]))
        con.commit()
        flash("Hasło zmienione.", "ok")
    con.close()
    default = "dashboard"
    return redirect(request.referrer or url_for(default))


# ---------- dashboard (tydzień w magazynie) ----------

@app.route("/dashboard")
@login_required
def dashboard():
    today = local_today()
    week_end = today + timedelta(days=6)
    today_s, week_end_s = today.isoformat(), week_end.isoformat()
    con = get_db()
    base_sql = """SELECT r.*, u.username, u.first_name, u.last_name,
                  e.code, e.name, e.location, IFNULL(e.catalog,'main') AS catalog,
                  w.name AS warehouse_name
                  FROM reservations r
                  JOIN users u ON u.id=r.user_id
                  JOIN equipment e ON e.id=r.equipment_id
                  LEFT JOIN warehouses w ON w.id=e.warehouse_id"""
    out_week = con.execute(
        base_sql + """ WHERE r.status='rezerwacja'
                       AND r.date_from>=? AND r.date_from<=?
                       ORDER BY r.date_from, e.code""",
        (today_s, week_end_s)).fetchall()
    back_week = con.execute(
        base_sql + """ WHERE r.status='wydane'
                       AND r.date_to>=? AND r.date_to<=?
                       ORDER BY r.date_to, e.code""",
        (today_s, week_end_s)).fetchall()
    overdue = con.execute(
        base_sql + " WHERE r.status='wydane' AND r.date_to<? ORDER BY r.date_to",
        (today_s,)).fetchall()
    con.close()
    days_overdue = {r["id"]: (today - date.fromisoformat(r["date_to"])).days
                    for r in overdue}
    return render_template("dashboard.html", out_week=out_week,
                           back_week=back_week, overdue=overdue,
                           today=today_s, week_end=week_end_s,
                           days_overdue=days_overdue, dn=display_name)


# ---------- rejestr sprzętu ----------

def _equipment_list_filters(catalog="main"):
    """Parametry filtrów listy sprzętu (z request.args)."""
    return {
        "q": request.args.get("q", "").strip(),
        "f_project": request.args.get("project", "").strip(),
        "f_owner": request.args.get("owner", "").strip(),
        "f_brand": request.args.get("brand", "").strip(),
        "f_warehouse": request.args.get("warehouse", "").strip(),
        "f_own": request.args.get("own", "").strip(),
        "f_condition": request.args.get("condition", "").strip(),
        "catalog": catalog,
    }


def _equipment_where(filters):
    """WHERE + params dla listy sprzętu (bez archiwum)."""
    catalog = filters.get("catalog") or "main"
    where, params = ["IFNULL(e.catalog,'main')=?", "IFNULL(e.archived,0)=0"], [catalog]
    q = filters.get("q") or ""
    if q:
        where.append("(e.code LIKE ? OR e.name LIKE ?)")
        params += [f"%{q}%"] * 2
    if filters.get("f_project"):
        where.append("e.project_number = ?"); params.append(filters["f_project"])
    if filters.get("f_owner"):
        where.append("e.owner = ?"); params.append(filters["f_owner"])
    if filters.get("f_brand"):
        where.append(
            "REPLACE(REPLACE(UPPER(TRIM(e.brand)), ' -', '-'), '- ', '-') = ?")
        params.append(_normalize_brand_key(filters["f_brand"]))
    f_warehouse = filters.get("f_warehouse") or ""
    if f_warehouse.isdigit():
        wid = int(f_warehouse)
        where.append(
            """(e.warehouse_id = ? OR e.id IN (
                 SELECT equipment_id FROM equipment_stock
                 WHERE warehouse_id=? AND quantity > 0))""")
        params.extend([wid, wid])
    if filters.get("f_own") == "1":
        where.append("e.material_type = 'wlasny'")
    f_condition = filters.get("f_condition") or ""
    if f_condition:
        if f_condition == "uszkodzony":
            where.append("(e.condition = ? OR IFNULL(e.damaged_quantity,0) > 0)")
            params.append(f_condition)
        else:
            where.append("e.condition = ?"); params.append(f_condition)
    return where, params


def _fetch_equipment_rows(con, filters):
    where, params = _equipment_where(filters)
    sql = """SELECT e.*, w.name AS warehouse_name FROM equipment e
             LEFT JOIN warehouses w ON w.id=e.warehouse_id
             WHERE """ + " AND ".join(where) + " ORDER BY e.code"
    return con.execute(sql, params).fetchall()


def _render_equipment_index(catalog="main"):
    """Lista sprzętu dla katalogu main lub tcl."""
    filters = _equipment_list_filters(catalog)
    q = filters["q"]
    f_project = filters["f_project"]
    f_owner = filters["f_owner"]
    f_brand = filters["f_brand"]
    f_warehouse = filters["f_warehouse"]
    f_own = filters["f_own"]
    f_condition = filters["f_condition"]
    per_page_raw = (request.args.get("per_page") or "all").strip().lower()
    page_raw = (request.args.get("page") or "1").strip()

    con = get_db()
    all_items = _fetch_equipment_rows(con, filters)

    total = len(all_items)
    per_page_choices = ["all", 25, 50, 100]
    if per_page_raw in ("25", "50", "100"):
        per_page = int(per_page_raw)
        per_page_value = per_page_raw
        try:
            page = max(1, int(page_raw))
        except ValueError:
            page = 1
        total_pages = max(1, (total + per_page - 1) // per_page) if total else 1
        if page > total_pages:
            page = total_pages
        start = (page - 1) * per_page
        items = all_items[start:start + per_page]
    else:
        per_page = total or 1
        per_page_value = "all"
        page = 1
        total_pages = 1
        items = all_items

    projects = [r[0] for r in con.execute(
        """SELECT DISTINCT project_number FROM equipment
           WHERE IFNULL(project_number,'')!='' AND IFNULL(catalog,'main')=?
             AND IFNULL(archived,0)=0 ORDER BY 1""",
        (catalog,))]
    owners = [r[0] for r in con.execute(
        """SELECT DISTINCT owner FROM equipment
           WHERE IFNULL(owner,'')!='' AND IFNULL(catalog,'main')=?
             AND IFNULL(archived,0)=0 ORDER BY 1""",
        (catalog,))]
    brands_raw = [r[0] for r in con.execute(
        """SELECT DISTINCT brand FROM equipment
           WHERE IFNULL(brand,'')!='' AND IFNULL(catalog,'main')=?
             AND IFNULL(archived,0)=0 ORDER BY 1""",
        (catalog,))]
    brands = _dedupe_brand_labels(brands_raw)
    warehouses = active_warehouses(con)

    today = local_today().isoformat()
    availability = {
        it["id"]: _usable_qty(it) - reserved_qty(con, it["id"], today, today)
        for it in items
    }

    eids = [it["id"] for it in items]
    stock_map = {}
    if eids:
        stock_rows = con.execute(
            f"""SELECT es.equipment_id, es.quantity, IFNULL(es.location,'') loc,
                       w.name wh
                FROM equipment_stock es
                LEFT JOIN warehouses w ON w.id=es.warehouse_id
                WHERE es.equipment_id IN ({','.join('?'*len(eids))}) AND es.quantity > 0
                ORDER BY es.quantity DESC""",
            eids).fetchall()
        for sr in stock_rows:
            stock_map.setdefault(sr["equipment_id"], []).append(sr)

    con.close()
    return render_template(
        "index.html", items=items, availability=availability,
        stock_map=stock_map,
        q=q, f_project=f_project, f_owner=f_owner, f_brand=f_brand,
        f_warehouse=f_warehouse, f_own=f_own, f_condition=f_condition,
        projects=projects, owners=owners, brands=brands, warehouses=warehouses,
        total=total, page=page, total_pages=total_pages,
        per_page=per_page_value, per_page_choices=per_page_choices,
        catalog=catalog,
    )


@app.route("/")
@login_required
def index():
    return _render_equipment_index("main")


@app.route("/tcl")
@login_required
@tcl_required
def tcl_index():
    return _render_equipment_index("tcl")



def _equipment_form_values(form, files, current=None, primary_photo=None):
    """Wartości do INSERT/UPDATE equipment. primary_photo – pierwsze z galerii."""
    photo = primary_photo if primary_photo is not None else (
        current["photo"] if current else None)
    pack_photo = save_photo(files.get("packaging_photo")) or \
        (current["packaging_photo"] if current else None)
    wid = form.get("warehouse_id", "")
    return (
        form["code"].strip(), form["project_number"].strip(), form["name"].strip(),
        form["dimensions"].strip(), photo, pack_photo, form["location"].strip(),
        int(wid) if wid.isdigit() else None,
        form["owner"].strip(), form.get("brand", "").strip(),
        form.get("material_type", "klient"),
        form.get("condition", "sprawny"), form.get("condition_notes", "").strip(),
        form.get("storage_instructions", "").strip(),
        max(0, int(form.get("quantity") or 0)), form["notes"].strip(),
    )


def _save_new_photos(files):
    """Zapisuje nowe pliki z pola photos (multiple). Zwraca listę nazw plików."""
    out = []
    for f in files.getlist("photos"):
        fname = save_photo(f)
        if fname:
            out.append(fname)
    return out


EQ_COLS = """code, project_number, name, dimensions, photo, packaging_photo,
             location, warehouse_id, owner, brand, material_type,
             condition, condition_notes, storage_instructions, quantity, notes"""


@app.route("/equipment/new", methods=["GET", "POST"])
@login_required
def equipment_new():
    catalog = (request.values.get("catalog") or "main").strip().lower()
    if catalog not in ("main", "tcl"):
        catalog = "main"
    if catalog == "tcl":
        if not can_manage_tcl():
            flash("Brak dostępu do katalogu TCL.", "error")
            return redirect(url_for("dashboard"))
    elif session.get("role") != "admin":
        flash("Wymagane uprawnienia administratora.", "error")
        return redirect(url_for("index"))

    con = get_db()
    if request.method == "POST":
        try:
            new_photos = _save_new_photos(request.files)[:MAX_PHOTOS]
            primary = new_photos[0] if new_photos else None
            vals = _equipment_form_values(request.form, request.files, primary_photo=primary)
            cur = con.execute(
                f"INSERT INTO equipment ({EQ_COLS}) VALUES ({','.join('?'*16)})", vals)
            eid = cur.lastrowid
            con.execute("UPDATE equipment SET catalog=? WHERE id=?", (catalog, eid))
            for i, fn in enumerate(new_photos):
                con.execute(
                    """INSERT INTO equipment_photos (equipment_id, filename, sort_order, kind)
                       VALUES (?,?,?,'normal')""",
                    (eid, fn, i))
            replace_equipment_stock(con, eid, vals[7], vals[6], vals[14])
            con.commit()
            flash("Sprzęt dodany.", "ok")
            return redirect(url_for("tcl_index" if catalog == "tcl" else "index"))
        except Exception as e:
            msg = str(e)
            if "UNIQUE" in msg:
                code = (request.form.get("code") or "").strip()
                arch = con.execute(
                    "SELECT id FROM equipment WHERE code=? AND IFNULL(archived,0)=1",
                    (code,)).fetchone() if code else None
                if arch:
                    flash(f"Kod {code} jest w archiwum – przywróć go zamiast dodawać nowy "
                          f"(Archiwum → Przywróć).", "error")
                else:
                    flash("Błąd: kod już istnieje", "error")
            else:
                flash(f"Błąd: {e}", "error")
        finally:
            con.close()
        return redirect(url_for("equipment_new", catalog=catalog))
    warehouses = active_warehouses(con)
    con.close()
    return render_template("equipment_form.html", eq=None, warehouses=warehouses,
                           photos=[], catalog=catalog)


@app.route("/equipment/<int:eid>/edit", methods=["GET", "POST"])
@login_required
def equipment_edit(eid):
    con = get_db()
    eq = con.execute("SELECT * FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        abort(404)
    cat = equipment_catalog(eq)
    if cat == "tcl":
        if not can_manage_tcl():
            con.close()
            flash("Brak dostępu do edycji sprzętu TCL.", "error")
            return redirect(url_for("equipment_detail", eid=eid))
    elif session.get("role") != "admin":
        con.close()
        flash("Wymagane uprawnienia administratora.", "error")
        return redirect(url_for("equipment_detail", eid=eid))
    photos = equipment_photo_list(con, eid)
    photo_kinds = equipment_photo_kind_map(con, eid)
    if request.method == "POST":
        try:
            # które istniejące zostawić
            keep = request.form.getlist("keep_photo")
            kept = [fn for fn in photos if fn in keep]
            added = _save_new_photos(request.files)
            final = (kept + added)[:MAX_PHOTOS]
            primary = final[0] if final else None
            vals = _equipment_form_values(request.form, request.files, current=eq,
                                          primary_photo=primary)
            sets = ", ".join(c.strip() + "=?" for c in EQ_COLS.split(","))
            con.execute(f"UPDATE equipment SET {sets} WHERE id=?", vals + (eid,))
            con.execute("DELETE FROM equipment_photos WHERE equipment_id=?", (eid,))
            for i, fn in enumerate(final):
                kind = photo_kinds.get(fn, "normal") if fn in kept else "normal"
                con.execute(
                    """INSERT INTO equipment_photos (equipment_id, filename, sort_order, kind)
                       VALUES (?,?,?,?)""",
                    (eid, fn, i, kind))
            # ręczna edycja karty: jeden magazyn/miejsce zgodne z formularzem
            replace_equipment_stock(con, eid, vals[7], vals[6], vals[14])
            con.commit()
            flash("Zapisano zmiany.", "ok")
            return redirect(url_for("equipment_detail", eid=eid))
        except Exception as e:
            flash(f"Błąd: {'kod już istnieje' if 'UNIQUE' in str(e) else e}", "error")
        finally:
            con.close()
        return redirect(url_for("equipment_edit", eid=eid))
    warehouses = active_warehouses(con)
    con.close()
    return render_template("equipment_form.html", eq=eq, warehouses=warehouses,
                           photos=photos, photo_kinds=photo_kinds)


@app.route("/equipment/<int:eid>")
@login_required
def equipment_detail(eid):
    con = get_db()
    eq = con.execute("""SELECT e.*, w.name AS warehouse_name, w.address AS warehouse_address
                        FROM equipment e LEFT JOIN warehouses w ON w.id=e.warehouse_id
                        WHERE e.id=?""", (eid,)).fetchone()
    if not eq:
        abort(404)
    photo_rows = list(equipment_photo_rows(con, eid))
    if not photo_rows and eq["photo"]:
        photo_rows = [{"filename": eq["photo"], "kind": "normal"}]
    res = con.execute(
        """SELECT r.*, u.username, u.first_name, u.last_name, u.department AS owner_department
           FROM reservations r
           JOIN users u ON u.id=r.user_id
           WHERE r.equipment_id=? AND r.status != 'anulowana'
           ORDER BY r.date_from DESC""", (eid,)).fetchall()
    today = local_today().isoformat()
    avail_today = _usable_qty(eq) - reserved_qty(con, eid, today, today)
    manage_ids = {r["id"] for r in res if can_manage_reservation(r)}
    is_archived = bool(eq["archived"]) if "archived" in eq.keys() else False
    con.close()
    return render_template("equipment_detail.html", eq=eq, reservations=res,
                           avail_today=avail_today, today=today, dn=display_name,
                           photo_rows=photo_rows, manage_ids=manage_ids,
                           is_archived=is_archived)


@app.route("/equipment/<int:eid>/repair", methods=["POST"])
@login_required
def equipment_repair(eid):
    """Oznacz część (lub wszystkie) uszkodzone sztuki jako naprawione."""
    con = get_db()
    eq = con.execute("SELECT * FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        con.close()
        abort(404)
    dmg = int(eq["damaged_quantity"] or 0)
    if dmg <= 0:
        con.close()
        flash("Brak uszkodzonych sztuk do oznaczenia.", "error")
        return redirect(url_for("equipment_detail", eid=eid))
    try:
        qty = int(request.form.get("qty") or 0)
    except (TypeError, ValueError):
        qty = 0
    note = (request.form.get("note") or "").strip()
    if qty < 1 or qty > dmg:
        con.close()
        flash(f"Podaj liczbę sztuk od 1 do {dmg}.", "error")
        return redirect(url_for("equipment_detail", eid=eid))
    who = (session.get("full_name") or session.get("username") or "?").strip()
    stamp = (f"[{local_today().isoformat()}] naprawione ({qty} szt.) – {who}"
             + (f": {note}" if note else ""))
    new_dmg = dmg - qty
    if new_dmg == 0:
        con.execute(
            """UPDATE equipment SET damaged_quantity=0, condition='sprawny',
               condition_notes=IFNULL(condition_notes,'') ||
               CASE WHEN IFNULL(condition_notes,'')='' THEN '' ELSE char(10) END || ?
               WHERE id=?""",
            (stamp, eid))
        # brak uszkodzonych → zostaw zdjęcia z etykietą „naprawione” (bez WZ/PZ)
        con.execute(
            """UPDATE equipment_photos SET kind='repaired'
               WHERE equipment_id=? AND kind='damage'""",
            (eid,))
        sync_equipment_primary_photo(con, eid)
    else:
        con.execute(
            """UPDATE equipment SET damaged_quantity=?,
               condition_notes=IFNULL(condition_notes,'') ||
               CASE WHEN IFNULL(condition_notes,'')='' THEN '' ELSE char(10) END || ?
               WHERE id=?""",
            (new_dmg, stamp, eid))
    con.commit()
    con.close()
    left = f" Pozostało uszkodzonych: {new_dmg} szt." if new_dmg else " Wszystkie uszkodzone oznaczono jako sprawne."
    flash(f"Oznaczono jako naprawione: {qty} szt.{left}", "ok")
    return redirect(url_for("equipment_detail", eid=eid))


@app.route("/equipment/<int:eid>/delete", methods=["POST"])
@login_required
def equipment_delete(eid):
    con = get_db()
    eq = con.execute("SELECT * FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        con.close()
        abort(404)
    cat = equipment_catalog(eq)
    if cat == "tcl":
        if not can_manage_tcl():
            con.close()
            flash("Brak dostępu do usuwania sprzętu TCL.", "error")
            return redirect(url_for("equipment_detail", eid=eid))
    elif session.get("role") != "admin":
        con.close()
        flash("Wymagane uprawnienia administratora.", "error")
        return redirect(url_for("equipment_detail", eid=eid))
    active = con.execute(
        """SELECT COUNT(*) c FROM reservations WHERE equipment_id=?
           AND status IN ('rezerwacja','wydane')""", (eid,)).fetchone()["c"]
    if active:
        flash("Nie można usunąć – sprzęt ma aktywne rezerwacje.", "error")
    else:
        con.execute("DELETE FROM equipment_photos WHERE equipment_id=?", (eid,))
        con.execute("DELETE FROM reservations WHERE equipment_id=?", (eid,))
        con.execute("DELETE FROM equipment WHERE id=?", (eid,))
        con.commit()
        flash("Sprzęt usunięty.", "ok")
    con.close()
    return redirect(url_for("tcl_index" if cat == "tcl" else "index"))


# ---------- rezerwacje ----------

@app.route("/reservations")
@login_required
def reservations():
    f = request.args.get("status", "")
    mine = request.args.get("mine", "")
    overdue = request.args.get("overdue", "")
    today = local_today().isoformat()
    con = get_db()
    sql = """SELECT r.*, u.username, u.first_name, u.last_name, u.department AS owner_department,
                    e.code, e.name, e.photo, e.location, IFNULL(e.catalog,'main') AS catalog,
                    w.name AS warehouse_name,
                    COALESCE(iw.name, w.name) AS issue_warehouse_name
             FROM reservations r
             JOIN users u ON u.id=r.user_id JOIN equipment e ON e.id=r.equipment_id
             LEFT JOIN warehouses w ON w.id=e.warehouse_id
             LEFT JOIN warehouses iw ON iw.id=r.issue_warehouse_id"""
    where, params = [], []
    if f:
        where.append("r.status=?"); params.append(f)
    if mine == "1":
        where.append("r.user_id=?"); params.append(session["user_id"])
    if overdue == "1":
        where.append("r.status='wydane' AND r.date_to<?"); params.append(today)
    if where:
        sql += " WHERE " + " AND ".join(where)
    if mine == "1":
        order = " ORDER BY IFNULL(r.created_at, r.date_from) DESC, r.id DESC"
    else:
        order = " ORDER BY r.date_from DESC, r.id DESC"
    rows = con.execute(sql + order, params).fetchall()
    warehouses = active_warehouses(con)
    receivers = active_partners(con)
    manage_ids = {r["id"] for r in rows if can_manage_reservation(r)}
    con.close()
    return render_template("reservations.html", rows=rows, f=f, mine=mine,
                           overdue=overdue, today=today, dn=display_name,
                           warehouses=warehouses, receivers=receivers,
                           manage_ids=manage_ids,
                           self_pickup_value=SELF_PICKUP_VALUE)


@app.route("/equipment/<int:eid>/reserve", methods=["GET", "POST"])
@login_required
def reserve(eid):
    con = get_db()
    eq = con.execute("SELECT * FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        abort(404)
    if "archived" in eq.keys() and eq["archived"]:
        con.close()
        flash("Ten sprzęt jest w archiwum – nie można tworzyć rezerwacji. Admin może go przywrócić.", "error")
        return redirect(url_for("equipment_detail", eid=eid))
    if equipment_catalog(eq) == "tcl" and not can_manage_tcl():
        con.close()
        flash("Rezerwacje sprzętu TCL może tworzyć tylko dział Warrens lub admin.", "error")
        return redirect(url_for("equipment_detail", eid=eid))
    form_data = None
    if request.method == "POST":
        form_data = request.form
        try:
            d_from, d_to, permanent = reservation_dates_from_form(request.form)
            qty = max(1, int(request.form.get("quantity") or 1))
        except Exception as exc:
            flash(f"Nieprawidłowe dane formularza: {exc}", "error")
            d_from = d_to = None
            permanent = 0
            qty = 1
        proj = (request.form.get("project_number") or "").strip()
        receiver = resolve_receiver(request.form.get("receiver", ""))
        if not proj:
            flash("Podaj numer projektu.", "error")
        elif not receiver:
            flash("Wybierz, kto odbiera towar.", "error")
        elif permanent and not valid_pickup_date(d_from):
            flash("Podaj datę Od.", "error")
        elif not permanent and not valid_dates(d_from, d_to):
            flash("Nieprawidłowy zakres dat.", "error")
        elif not permanent and handoff_conflict(con, eid, d_from, d_to):
            # znajdź konflikt, żeby podać konkretną datę
            row = con.execute(
                """SELECT date_to FROM reservations
                   WHERE equipment_id=? AND status IN ('rezerwacja','wydane')
                   AND date_to=? LIMIT 1""", (eid, d_from)).fetchone()
            if row:
                flash(f"W dniu zwrotu ({row['date_to']}) sprzęt jest jeszcze zajęty – "
                      f"wypożyczenie możliwe od {next_free_after_return(row['date_to'])}.", "error")
            else:
                flash(f"Termin styka się z inną rezerwacją (zwrot i wydanie tego samego dnia). "
                      f"Ustaw start najwcześniej na dzień po zwrocie.", "error")
        else:
            taken = reserved_qty(con, eid, d_from, d_to)
            usable = _usable_qty(eq)
            free = usable - taken
            if qty > free:
                flash(f"Brak dostępności w tym terminie. Wolne sztuki: {free} z {usable} sprawnych "
                      f"({eq['quantity']} łącznie).", "error")
            else:
                rec = recipient_form_fields(request.form)
                rec_err = recipient_required_error(rec)
                if rec_err:
                    flash(rec_err, "error")
                else:
                    try:
                        con.execute(
                            """INSERT INTO reservations (equipment_id, user_id, client,
                               date_from, date_to, quantity, notes, receiver, permanent,
                               project_number,
                               recipient_name, recipient_contact, recipient_phone,
                               recipient_address, recipient_city, recipient_email)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (eid, session["user_id"], request.form["client"].strip(),
                             d_from, d_to, qty, request.form["notes"].strip(),
                             receiver, permanent, proj,
                             rec["recipient_name"], rec["recipient_contact"],
                             rec["recipient_phone"], rec["recipient_address"],
                             rec["recipient_city"], rec["recipient_email"]))
                        upsert_recipient(con, rec["recipient_name"], rec["recipient_contact"],
                                         rec["recipient_phone"], rec["recipient_address"],
                                         rec["recipient_email"], rec["recipient_city"])
                        con.commit()
                        con.close()
                        flash("Rezerwacja utworzona.", "ok")
                        return redirect(url_for("equipment_detail", eid=eid))
                    except Exception as exc:
                        flash(f"Nie udało się zapisać rezerwacji: {exc}. Spróbuj ponownie.", "error")
    receivers = active_partners(con)
    recipients = recent_recipients(con)
    projects = project_suggestions(con)
    con.close()
    return render_template("reserve.html", eq=eq, receivers=receivers,
                           recipients=recipients, projects=projects, form=form_data,
                           self_pickup_value=SELF_PICKUP_VALUE)


@app.route("/reserve-multi", methods=["GET", "POST"])
@login_required
def reserve_multi():
    con = get_db()
    if request.method == "POST":
        ids = request.form.getlist("eid")
    else:
        ids = request.args.getlist("eid")
    ids = [int(i) for i in ids if str(i).isdigit()]
    if not ids:
        flash("Zaznacz przynajmniej jeden sprzęt.", "error")
        con.close()
        return redirect(url_for("index"))
    items = con.execute(
        f"""SELECT e.*, w.name AS warehouse_name FROM equipment e
            LEFT JOIN warehouses w ON w.id=e.warehouse_id
            WHERE e.id IN ({','.join('?'*len(ids))}) ORDER BY e.code""",
        ids).fetchall()
    if any((it["archived"] if "archived" in it.keys() else 0) for it in items):
        flash("Nie można rezerwować sprzętu z archiwum.", "error")
        con.close()
        return redirect(url_for("index"))
    catalogs = {equipment_catalog(it) for it in items}
    if "tcl" in catalogs and not can_manage_tcl():
        flash("Rezerwacje sprzętu TCL może tworzyć tylko dział Warrens lub admin.", "error")
        con.close()
        return redirect(url_for("index"))
    if len(catalogs) > 1:
        flash("Nie łącz w jednej rezerwacji pozycji z katalogu głównego i TCL.", "error")
        con.close()
        return redirect(url_for("tcl_index" if "tcl" in catalogs else "index"))
    list_endpoint = "tcl_index" if catalogs == {"tcl"} else "index"

    wh_names = {it["warehouse_name"] for it in items if it["warehouse_name"]}
    multi_warehouse = len(wh_names) > 1
    form_data = None

    if request.method == "POST" and (
            "receiver" in request.form
            or "date_from" in request.form
            or request.form.get("permanent")):
        # POST z formularza wspólnej rezerwacji (daty opcjonalne przy trwałym)
        form_data = request.form
        d_from, d_to, permanent = reservation_dates_from_form(request.form)
        proj = (request.form.get("project_number") or "").strip()
        receiver = resolve_receiver(request.form.get("receiver", ""))
        if not proj:
            flash("Podaj numer projektu.", "error")
        elif not receiver:
            flash("Wybierz, kto odbiera towar.", "error")
        elif permanent and not valid_pickup_date(d_from):
            flash("Podaj datę Od.", "error")
        elif not permanent and not valid_dates(d_from, d_to):
            flash("Nieprawidłowy zakres dat.", "error")
        else:
            errors = []
            wanted = {}
            for it in items:
                raw = (request.form.get(f"qty_{it['id']}") or "0").strip()
                try:
                    qty = int(raw)
                except ValueError:
                    qty = 0
                if qty <= 0:
                    continue  # pomiń pozycję (usunięta / ilość 0)
                if not permanent and handoff_conflict(con, it["id"], d_from, d_to):
                    row = con.execute(
                        """SELECT date_to FROM reservations
                           WHERE equipment_id=? AND status IN ('rezerwacja','wydane')
                           AND date_to=? LIMIT 1""", (it["id"], d_from)).fetchone()
                    if row:
                        errors.append(
                            f"{it['code']}: zwrot {row['date_to']} – wolne od "
                            f"{next_free_after_return(row['date_to'])}")
                    else:
                        errors.append(f"{it['code']}: termin styka się z inną rezerwacją "
                                      f"(zwrot i wydanie tego samego dnia)")
                    continue
                usable = _usable_qty(it)
                free = usable - reserved_qty(con, it["id"], d_from, d_to)
                if qty > free:
                    errors.append(f"{it['code']}: wolne {free} z {usable} sprawnych szt.")
                wanted[it["id"]] = qty
            if not wanted:
                flash("Zostaw przynajmniej jedną pozycję z ilością > 0."
                      + ((" " + "; ".join(errors)) if errors else ""), "error")
            elif errors:
                flash("Brak dostępności – " + "; ".join(errors), "error")
            else:
                rec = recipient_form_fields(request.form)
                rec_err = recipient_required_error(rec)
                if rec_err:
                    flash(rec_err, "error")
                else:
                    gid = uuid.uuid4().hex[:8]
                    for it in items:
                        if it["id"] not in wanted:
                            continue
                        con.execute(
                            """INSERT INTO reservations (equipment_id, user_id, client,
                               date_from, date_to, quantity, notes, group_id, receiver, permanent,
                               project_number,
                               recipient_name, recipient_contact, recipient_phone,
                               recipient_address, recipient_city, recipient_email)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (it["id"], session["user_id"], request.form["client"].strip(),
                             d_from, d_to, wanted[it["id"]],
                             request.form["notes"].strip(), gid,
                             receiver, permanent, proj,
                             rec["recipient_name"], rec["recipient_contact"],
                             rec["recipient_phone"], rec["recipient_address"],
                             rec["recipient_city"], rec["recipient_email"]))
                    upsert_recipient(con, rec["recipient_name"], rec["recipient_contact"],
                                     rec["recipient_phone"], rec["recipient_address"],
                                     rec["recipient_email"], rec["recipient_city"])
                    con.commit()
                    con.close()
                    flash(f"Utworzono wspólną rezerwację ({len(wanted)} pozycji).", "ok")
                    return redirect(url_for("reservations"))
    receivers = active_partners(con)
    recipients = recent_recipients(con)
    projects = project_suggestions(con)
    # podpowiedź: wspólny numer projektu ze sprzętu, jeśli wszystkie mają ten sam
    eq_projects = {(it["project_number"] or "").strip() for it in items}
    eq_projects.discard("")
    default_proj = next(iter(eq_projects)) if len(eq_projects) == 1 else ""
    con.close()
    return render_template("reserve_multi.html", items=items, receivers=receivers,
                           recipients=recipients, projects=projects,
                           multi_warehouse=multi_warehouse,
                           wh_names=sorted(wh_names), form=form_data,
                           default_project=default_proj,
                           self_pickup_value=SELF_PICKUP_VALUE)


def _selected_reservations(con, rids):
    rids = [int(r) for r in rids if str(r).isdigit()]
    if not rids:
        return []
    return con.execute(
        f"""SELECT r.*, u.username, u.first_name, u.last_name, u.department AS owner_department,
            e.code, e.name, e.location, e.owner, e.brand,
            e.photo, e.dimensions, e.project_number AS equipment_project_number,
            e.storage_instructions, IFNULL(e.catalog,'main') AS catalog,
            w.name AS warehouse_name, w.address AS warehouse_address
            FROM reservations r JOIN users u ON u.id=r.user_id
            JOIN equipment e ON e.id=r.equipment_id
            LEFT JOIN warehouses w ON w.id=e.warehouse_id
            WHERE r.id IN ({','.join('?'*len(rids))}) ORDER BY e.code""",
        rids).fetchall()


def _return_location_for(form, r):
    """Miejsce przyjęcia – per pozycja (zbiorczo) albo wspólne pole."""
    loc = (form.get(f"item_return_location_{r['id']}") or "").strip()
    if not loc:
        loc = (form.get("return_location") or "").strip()
    return loc


def _return_form_valid(form, r=None):
    """Magazyn, miejsce i oddający wymagani przy zwrocie na magazyn; przy samej utylizacji – nie."""
    if form.get("dispose"):
        return True
    wid = (form.get("return_warehouse_id") or "").strip()
    returner = (form.get("returner") or "").strip()
    loc = _return_location_for(form, r) if r is not None else (form.get("return_location") or "").strip()
    return wid.isdigit() and bool(returner) and bool(loc)


def _eq_damaged_qty(eq):
    try:
        return max(0, int(eq["damaged_quantity"] or 0))
    except (KeyError, IndexError, TypeError, ValueError):
        return 0


def _usable_qty(eq):
    """Sztuki sprawne (do rezerwacji) = stan − uszkodzone."""
    return max(0, int(eq["quantity"]) - _eq_damaged_qty(eq))


def _damage_notes_for(form, r):
    notes = (form.get(f"item_damage_notes_{r['id']}") or "").strip()
    if not notes:
        notes = (form.get("damage_notes") or "").strip()
    return notes


def _attach_damage_photo(con, eid, files, rid=None):
    """Dodaje zdjęcie uszkodzenia do galerii karty produktu. Zwraca filename lub None."""
    if not files:
        return None
    f = None
    if rid is not None:
        f = files.get(f"item_damage_photo_{rid}")
    if (not f or not getattr(f, "filename", None)) and files.get("damage_photo"):
        f = files.get("damage_photo")
    fname = save_photo(f)
    if not fname:
        return None
    order = con.execute(
        "SELECT COALESCE(MAX(sort_order), -1) + 1 AS n FROM equipment_photos WHERE equipment_id=?",
        (eid,)).fetchone()["n"]
    con.execute(
        """INSERT INTO equipment_photos (equipment_id, filename, sort_order, kind)
           VALUES (?,?,?,'damage')""",
        (eid, fname, order))
    # uzupełnij miniaturę tylko gdy karta nie miała jeszcze zdjęcia
    eq = con.execute("SELECT photo FROM equipment WHERE id=?", (eid,)).fetchone()
    if eq and not eq["photo"]:
        sync_equipment_primary_photo(con, eid)
    return fname


def _process_qty_from_form(form, r):
    """Ile sztuk z rezerwacji obejmuje ta operacja (zwrot / utylizacja)."""
    rid = r["id"]
    raw = form.get(f"item_qty_{rid}")
    if raw in (None, ""):
        raw = form.get("return_qty") or form.get("qty")
    try:
        qty = int(raw) if raw not in (None, "") else int(r["quantity"])
    except (TypeError, ValueError):
        return None
    total = int(r["quantity"])
    if qty < 1 or qty > total:
        return None
    return qty


def _split_partial(con, r, process_qty):
    """Częściowy zwrot/utylizacja: reszta sztuk zostaje osobną rezerwacją „wydane”.

    Zwraca (processed_row, remaining_row_or_None).
    """
    total = int(r["quantity"])
    if process_qty >= total:
        return r, None
    remaining = total - process_qty

    def g(key, default=None):
        try:
            return r[key]
        except (KeyError, IndexError):
            return default

    cur = con.execute(
        """INSERT INTO reservations (
            equipment_id, user_id, client, date_from, date_to, quantity, status,
            group_id, receiver, permanent, project_number,
            issue_warehouse_id, issue_location,
            recipient_name, recipient_contact,
            recipient_phone, recipient_address, recipient_city, recipient_email, notes,
            issued_at, issued_by
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            r["equipment_id"], r["user_id"], g("client"), r["date_from"], r["date_to"],
            remaining, "wydane", g("group_id"), g("receiver"),
            1 if g("permanent") else 0, g("project_number"),
            g("issue_warehouse_id"), g("issue_location"),
            g("recipient_name"), g("recipient_contact"), g("recipient_phone"),
            g("recipient_address"), g("recipient_city"), g("recipient_email"), g("notes"),
            g("issued_at"), g("issued_by"),
        ),
    )
    rem_id = cur.lastrowid
    con.execute("UPDATE reservations SET quantity=? WHERE id=?", (process_qty, r["id"]))
    processed = con.execute("SELECT * FROM reservations WHERE id=?", (r["id"],)).fetchone()
    rem = con.execute("SELECT * FROM reservations WHERE id=?", (rem_id,)).fetchone()
    return processed, rem


def _dispose_notes_for(form, r):
    notes = (form.get(f"item_dispose_notes_{r['id']}") or "").strip()
    if not notes:
        notes = (form.get("dispose_notes") or "").strip()
    if not notes:
        # legacy: wyłącznie utylizacja (flaga dispose / item_action) w damage_notes
        if form.get("dispose") or (form.get(f"item_action_{r['id']}") or "") == "dispose":
            notes = (form.get("damage_notes") or "").strip()
    return notes


def _parse_return_ok_damage_qty(form, r, force_damage=False):
    """Zwraca (ok_qty, damage_qty, dispose_qty) albo None.

    Pola: item_qty_ok/damage/dispose_<id> albo return_qty_ok/damage/dispose.
    Legacy: dispose / item_action=dispose + item_qty; damage + qty; zwykły return_qty.
    """
    total = int(r["quantity"])
    rid = r["id"]
    ok_raw = form.get(f"item_qty_ok_{rid}")
    dmg_raw = form.get(f"item_qty_damage_{rid}")
    disp_raw = form.get(f"item_qty_dispose_{rid}")
    if ok_raw is None and dmg_raw is None and disp_raw is None:
        ok_raw = form.get("return_qty_ok")
        dmg_raw = form.get("return_qty_damage")
        disp_raw = form.get("return_qty_dispose")
    if ok_raw is not None or dmg_raw is not None or disp_raw is not None:
        try:
            ok = int(ok_raw) if ok_raw not in (None, "") else 0
            dmg = int(dmg_raw) if dmg_raw not in (None, "") else 0
            disp = int(disp_raw) if disp_raw not in (None, "") else 0
        except (TypeError, ValueError):
            return None
        if ok < 0 or dmg < 0 or disp < 0:
            return None
        if ok + dmg + disp < 1 or ok + dmg + disp > total:
            return None
        return ok, dmg, disp

    if form.get("dispose") or (form.get(f"item_action_{rid}") or "") == "dispose":
        qty = _process_qty_from_form(form, r)
        if qty is None:
            return None
        return 0, 0, qty

    qty = _process_qty_from_form(form, r)
    if qty is None:
        return None
    if force_damage or form.get("damage"):
        return 0, qty, 0
    return qty, 0, 0


def _parse_xbs_form(form):
    """Zwraca dict meta awizacji XBS albo None; False gdy włączona bez daty dostawy."""
    if (form.get("xbs_awizacja") or "").strip() not in ("1", "on", "true", "yes"):
        return None
    delivery_date = (form.get("xbs_delivery_date") or "").strip()
    if not delivery_date:
        return False
    material = (form.get("xbs_material") or "").strip()
    if material and material not in MATERIAL_OPTIONS:
        material = ""
    return {
        "supplier": (form.get("xbs_supplier") or "").strip(),
        "supplier_person": (form.get("xbs_supplier_person") or "").strip(),
        "supplier_phone": (form.get("xbs_supplier_phone") or "").strip(),
        "supplier_address": (form.get("xbs_supplier_address") or "").strip(),
        "order_no": (form.get("xbs_order_no") or "").strip(),
        "delivery_date": delivery_date,
        "delivery_time": (form.get("xbs_delivery_time") or "").strip(),
        "notes": (form.get("xbs_notes") or "").strip(),
        "carrier": (form.get("xbs_carrier") or "").strip(),
        "plate": (form.get("xbs_plate") or "").strip(),
        "qty_per_pallet": (form.get("xbs_qty_per_pallet") or "").strip(),
        "weight": (form.get("xbs_weight") or "").strip(),
        "material": material,
        "pallets": (form.get("xbs_pallets") or "").strip(),
    }


def _store_xbs_session(rids, meta):
    session["xbs_awizacja"] = {
        "rids": [int(x) for x in rids],
        "meta": meta,
    }


def _save_xbs_on_reservations(con, rids, meta):
    """Zapisuje dane awizacji XBS przy rezerwacjach (do ponownego pobrania jak WZ/PZ)."""
    if not rids or not meta:
        return None
    batch_id = uuid.uuid4().hex
    payload = json.dumps(meta, ensure_ascii=False)
    for rid in rids:
        con.execute(
            """UPDATE reservations SET xbs_awizacja_json=?, xbs_batch_id=? WHERE id=?""",
            (payload, batch_id, int(rid)),
        )
    return batch_id


def _xbs_meta_from_row(r):
    raw = None
    try:
        raw = r["xbs_awizacja_json"]
    except (KeyError, IndexError, TypeError):
        raw = None
    if not raw:
        return None
    try:
        return json.loads(raw)
    except (TypeError, ValueError, json.JSONDecodeError):
        return None


def _xbs_items_for_rids(con, rids):
    items = []
    for rid in rids:
        row = con.execute(
            """SELECT e.code, e.name, r.quantity
               FROM reservations r JOIN equipment e ON e.id=r.equipment_id
               WHERE r.id=?""",
            (rid,),
        ).fetchone()
        if row:
            items.append({
                "code": row["code"],
                "name": row["name"],
                "quantity": row["quantity"],
            })
    return items


def _xbs_file_response(con, rids, meta):
    items = _xbs_items_for_rids(con, rids)
    if not items or not meta:
        return None
    buf = build_xbs_awizacja_xlsx(items, meta)
    return buf, xbs_filename(local_now().strftime("%Y%m%d_%H%M"))


def _xbs_payload_for_download(rids_arg=None):
    """Meta + lista rid z sesji (po wydaniu) albo z query (Safari / odświeżenie)."""
    payload = session.get("xbs_awizacja")
    if payload and payload.get("rids"):
        return payload["rids"], payload.get("meta") or {}
    rids = [int(x) for x in (rids_arg or []) if str(x).isdigit()]
    if not rids:
        return None, None
    con = get_db()
    try:
        r = con.execute("SELECT * FROM reservations WHERE id=?", (rids[0],)).fetchone()
        if not r:
            return None, None
        meta = _xbs_meta_from_row(r)
        if not meta:
            return None, None
        batch = None
        try:
            batch = r["xbs_batch_id"]
        except (KeyError, IndexError, TypeError):
            batch = None
        if batch:
            rows = con.execute(
                "SELECT id FROM reservations WHERE xbs_batch_id=? ORDER BY id",
                (batch,),
            ).fetchall()
            rids = [row["id"] for row in rows]
        return rids, meta
    finally:
        con.close()


def _redirect_after_issue(done_ids, xbs_meta):
    """Po wydaniu: PDF i opcjonalnie XBS (osobne żądania GET, z opóźnieniem w UI)."""
    flash(f"Wydano: {len(done_ids)} pozycji. Dokumenty pobiorą się za chwilę (lub użyj linku na liście).", "ok")
    if xbs_meta:
        _store_xbs_session(done_ids, xbs_meta)
        flash("Awizacja XBS: Excel pobierze się zaraz po PDF (lub link w kolumnie Dokumenty).", "ok")
    kw = {"auto_pdf": "wydanie", "rid": done_ids}
    if xbs_meta:
        kw["xbs_hint"] = "1"
    return redirect(url_for("reservations", **kw))


def _finalize_one_return(con, r, user_id, form, files=None, damage=False, damage_notes=""):
    """Zamyka jedną rezerwację jako zwrot (cała quantity wiersza)."""
    if r["status"] != "wydane":
        return False
    wid = int((form.get("return_warehouse_id") or "").strip())
    loc = _return_location_for(form, r)
    raw_returner = (form.get("returner") or "").strip() or (r["receiver"] or "")
    returner = resolve_returner(raw_returner)
    try:
        issue_wid = r["issue_warehouse_id"]
    except (KeyError, IndexError, TypeError):
        issue_wid = None
    if not issue_wid:
        eq = con.execute("SELECT warehouse_id, IFNULL(location,'') loc FROM equipment WHERE id=?",
                         (r["equipment_id"],)).fetchone()
        if eq and eq["warehouse_id"]:
            issue_loc = _location_from_stock(con, r["equipment_id"], eq["warehouse_id"])
            con.execute("""UPDATE reservations SET issue_warehouse_id=?, issue_location=?
                           WHERE id=?""", (eq["warehouse_id"], issue_loc, r["id"]))
    now = local_now().isoformat(timespec="seconds")
    dmg_flag = 1 if damage else 0
    notes = damage_notes if damage else None
    con.execute("""UPDATE reservations SET status='zwrócone', returned_at=?,
                   returned_by=?, damage=?, damage_notes=?,
                   return_warehouse_id=?, return_location=?, returner=? WHERE id=?""",
                (now, user_id, dmg_flag, notes, wid, loc, returner, r["id"]))
    eid = r["equipment_id"]
    move_equipment_stock_on_return(con, eid, r["quantity"], wid, loc)
    if damage:
        stamp = (f"[{local_today().isoformat()}] zwrot uszkodzony rez. #{r['id']} "
                 f"({r['quantity']} szt.): {damage_notes}")
        con.execute("""UPDATE equipment SET damaged_quantity=IFNULL(damaged_quantity,0)+?,
                       condition='uszkodzony',
                       condition_notes=IFNULL(condition_notes,'') ||
                       CASE WHEN IFNULL(condition_notes,'')='' THEN '' ELSE char(10) END || ?
                       WHERE id=?""", (r["quantity"], stamp, eid))
        _attach_damage_photo(con, eid, files, r["id"])
    return True


def _finalize_one_dispose(con, r, user_id, notes):
    """Zamyka jedną rezerwację jako utylizacja (cała quantity wiersza)."""
    if r["status"] != "wydane" or not notes:
        return False
    eq = con.execute("SELECT quantity, warehouse_id FROM equipment WHERE id=?",
                     (r["equipment_id"],)).fetchone()
    if not eq or eq["quantity"] < r["quantity"]:
        return False
    now = local_now().isoformat(timespec="seconds")
    con.execute("""UPDATE reservations SET status='utylizacja', returned_at=?,
                   returned_by=?, damage=1, damage_notes=?, permanent=1 WHERE id=?""",
                (now, user_id, notes, r["id"]))
    eid = r["equipment_id"]
    new_qty = eq["quantity"] - r["quantity"]
    stamp = f"[{local_today().isoformat()}] utylizacja rez. #{r['id']} ({r['quantity']} szt.): {notes}"
    take_equipment_stock(con, eid, r["quantity"], prefer_warehouse_id=eq["warehouse_id"])
    if new_qty <= 0:
        con.execute("""UPDATE equipment SET quantity=0, condition='do utylizacji',
                       condition_notes=IFNULL(condition_notes,'') ||
                       CASE WHEN IFNULL(condition_notes,'')='' THEN '' ELSE char(10) END || ?
                       WHERE id=?""", (stamp, eid))
        con.execute("DELETE FROM equipment_stock WHERE equipment_id=?", (eid,))
    else:
        con.execute("""UPDATE equipment SET quantity=?,
                       condition_notes=IFNULL(condition_notes,'') ||
                       CASE WHEN IFNULL(condition_notes,'')='' THEN '' ELSE char(10) END || ?
                       WHERE id=?""", (new_qty, stamp, eid))
        sync_equipment_from_stock(con, eid, keep_total=True)
    return True


def _apply_return(con, r, user_id, form, files=None, force_damage=False):
    """Przyjmuje zwrot. W jednym kroku: N sprawnych + M uszkodzonych + K utylizacja.

    force_damage=True – całość jako uszkodzona (legacy zbiorczy item_action=damage).
    Zwraca listę id zamkniętych rezerwacji albo False.
    """
    if r["status"] != "wydane":
        return False
    parsed = _parse_return_ok_damage_qty(form, r, force_damage=force_damage)
    if parsed is None:
        return False
    ok_qty, dmg_qty, disp_qty = parsed
    damage_notes = _damage_notes_for(form, r) if dmg_qty else ""
    dispose_notes = _dispose_notes_for(form, r) if disp_qty else ""
    if dmg_qty and not damage_notes:
        return False
    if disp_qty and not dispose_notes:
        return False
    if (ok_qty or dmg_qty) and not _return_form_valid(form, r):
        return False
    process = ok_qty + dmg_qty + disp_qty
    r, _stays = _split_partial(con, r, process)
    done_ids = []
    rest = r

    if disp_qty > 0:
        if ok_qty + dmg_qty > 0:
            r_disp, rest = _split_partial(con, rest, disp_qty)
        else:
            r_disp, rest = rest, None
        if not _finalize_one_dispose(con, r_disp, user_id, dispose_notes):
            return False
        done_ids.append(r_disp["id"])

    if dmg_qty > 0 and rest is not None:
        if ok_qty > 0:
            r_dmg, rest = _split_partial(con, rest, dmg_qty)
        else:
            r_dmg, rest = rest, None
        if not _finalize_one_return(con, r_dmg, user_id, form, files=files,
                                    damage=True, damage_notes=damage_notes):
            return False
        done_ids.append(r_dmg["id"])

    if ok_qty > 0 and rest is not None:
        if not _finalize_one_return(con, rest, user_id, form, files=None, damage=False):
            return False
        done_ids.append(rest["id"])

    return done_ids


def _is_aggregate_location(loc):
    """True dla sklejonych opisów multi-magazyn (nie pojedynczego miejsca, np. „RP 5, P 1”)."""
    loc = (loc or "").strip()
    if not loc:
        return False
    if " szt" in loc.lower():
        return True
    # stock_summary: „Perła/lok: 2, Stalowa: 1”
    return bool(re.search(r":\s*\d+", loc) and "," in loc)


def _location_from_stock(con, equipment_id, warehouse_id):
    """Miejsce w magazynie przy wydaniu – z wiersza stock, nie ze sklejonego opisu."""
    if not warehouse_id:
        return ""
    row = con.execute(
        """SELECT IFNULL(location,'') loc FROM equipment_stock
           WHERE equipment_id=? AND warehouse_id=? AND quantity > 0
           ORDER BY quantity DESC, id LIMIT 1""",
        (equipment_id, warehouse_id)).fetchone()
    if row:
        return (row["loc"] or "").strip()
    eq = con.execute(
        "SELECT IFNULL(location,'') loc FROM equipment WHERE id=?",
        (equipment_id,)).fetchone()
    loc = (eq["loc"] or "").strip() if eq else ""
    return "" if _is_aggregate_location(loc) else loc


def _pdf_location(eq, preferred=""):
    """Miejsce w magazynie do PDF – preferowane pole rezerwacji, potem stan sprzętu."""
    loc = (preferred or "").strip()
    if loc and not _is_aggregate_location(loc):
        return loc
    fallback = (eq.get("location") or "").strip()
    if fallback and not _is_aggregate_location(fallback):
        return fallback
    return loc or fallback or ""


def _effective_issue_date_from(r):
    """Start rezerwacji przy wydaniu (przyszły start → dziś)."""
    today = local_today().isoformat()
    date_from = r["date_from"]
    if date_from > today:
        date_from = today
    return date_from


def _validate_return_date(con, r, new_to, exclude_id=None):
    """Sprawdza nowy termin zwrotu (kolizje z innymi rezerwacjami)."""
    date_from = _effective_issue_date_from(r) if r["status"] == "rezerwacja" else r["date_from"]
    if not valid_dates(date_from, new_to):
        return False, "Nieprawidłowa data zwrotu (musi być ≥ daty rozpoczęcia)."
    eid = r["equipment_id"]
    eq = con.execute("SELECT quantity, code FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        return False, "Brak sprzętu."
    rid = r["id"] if exclude_id is None else exclude_id
    taken = reserved_qty(con, eid, date_from, new_to, exclude_id=rid)
    free = _usable_qty(eq) - taken
    if r["quantity"] > free or handoff_conflict(con, eid, date_from, new_to, exclude_id=rid):
        return False, (
            f"Niedostępny w tych dniach – jest już rezerwacja ({eq['code']}, zwrot {new_to})."
        )
    return True, None


def _apply_issue(con, r, user_id, permanent=None, date_to=None):
    """Oznacza rezerwację jako wydaną.

    permanent=None → bierze flagę z rezerwacji (checkbox przy tworzeniu).
    Przy permanent=True towar schodzi ze stanu i nie wraca.
    date_to: opcjonalny nowy termin zwrotu (walidacja kolizji przed wywołaniem).
    """
    if r["status"] != "rezerwacja":
        return False
    if permanent is None:
        permanent = bool(r["permanent"]) if "permanent" in r.keys() else False
    else:
        permanent = bool(permanent)
    now = local_now()
    date_from = _effective_issue_date_from(r)
    if date_to is None:
        date_to = r["date_to"]
    if not valid_dates(date_from, date_to):
        return False
    status = "wydane trwale" if permanent else "wydane"
    if permanent:
        eq = con.execute("SELECT quantity, warehouse_id, IFNULL(location,'') loc FROM equipment WHERE id=?",
                         (r["equipment_id"],)).fetchone()
        if not eq or eq["quantity"] < r["quantity"]:
            return False
        if not take_equipment_stock(con, r["equipment_id"], r["quantity"],
                                    prefer_warehouse_id=eq["warehouse_id"]):
            return False
        con.execute("UPDATE equipment SET quantity = quantity - ? WHERE id=?",
                    (r["quantity"], r["equipment_id"]))
        sync_equipment_from_stock(con, r["equipment_id"], keep_total=True)
        left = con.execute("SELECT quantity FROM equipment WHERE id=?",
                           (r["equipment_id"],)).fetchone()
        if left and int(left["quantity"] or 0) <= 0:
            con.execute(
                """UPDATE equipment SET archived=1, archived_at=?, quantity=0,
                   damaged_quantity=0 WHERE id=?""",
                (now.isoformat(timespec="seconds"), r["equipment_id"]))
    else:
        eq = con.execute("SELECT warehouse_id, IFNULL(location,'') loc FROM equipment WHERE id=?",
                         (r["equipment_id"],)).fetchone()
    issue_wid = eq["warehouse_id"] if eq else None
    issue_loc = _location_from_stock(con, r["equipment_id"], issue_wid)
    con.execute("""UPDATE reservations SET status=?, issued_at=?, issued_by=?,
                   date_from=?, date_to=?, permanent=?,
                   issue_warehouse_id=?, issue_location=? WHERE id=?""",
                (status, now.isoformat(timespec="seconds"), user_id, date_from, date_to,
                 1 if permanent else 0, issue_wid, issue_loc, r["id"]))
    return True


def _apply_dispose(con, r, user_id, form):
    """Towar wydany nie wraca (utylizacja / zniszczenie). Schodzi ze stanu magazynowego.

    Możliwa częściowa liczba sztuk – reszta zostaje jako „wydane”.
    """
    if r["status"] != "wydane":
        return False
    notes = _dispose_notes_for(form, r)
    if not notes:
        return False
    qty = _process_qty_from_form(form, r)
    if qty is None:
        return False
    r, _stays = _split_partial(con, r, qty)
    return _finalize_one_dispose(con, r, user_id, notes)

def _eq_for_pdf(con, equipment_id, kind=None, reservation=None, ret_wid=None, ret_loc=None):
    """Dane sprzętu do PDF.

    WZ: magazyn z momentu wydania (issue_*), jeśli zapisany.
    PZ: Magazyn przyjęcia ze zwrotu; Magazyn wydania z issue_*.
    """
    eq = con.execute(
        """SELECT e.*, w.name AS warehouse_name, w.address AS warehouse_address
           FROM equipment e LEFT JOIN warehouses w ON w.id=e.warehouse_id
           WHERE e.id=?""",
        (equipment_id,)).fetchone()
    if not eq:
        return None
    eq = dict(eq)
    eq["equipment_location"] = (eq.get("location") or "").strip()

    issue_wid = None
    issue_loc = ""
    if reservation is not None:
        try:
            issue_wid = reservation["issue_warehouse_id"]
            issue_loc = (reservation["issue_location"] or "").strip()
        except (KeyError, IndexError, TypeError):
            pass
    if issue_wid:
        iss_wh = con.execute("SELECT name, address FROM warehouses WHERE id=?",
                             (issue_wid,)).fetchone()
        if iss_wh:
            eq["issue_warehouse_name"] = iss_wh["name"]
            eq["issue_warehouse_address"] = iss_wh["address"] or ""
            eq["issue_location"] = issue_loc

    if kind == "wydanie" and eq.get("issue_warehouse_name"):
        eq["warehouse_name"] = eq["issue_warehouse_name"]
        eq["warehouse_address"] = eq.get("issue_warehouse_address") or ""
        eq["location"] = _pdf_location(eq, issue_loc)
        return eq

    if kind != "przyjecie":
        return eq

    wid = ret_wid
    loc = ret_loc
    if wid is None and reservation is not None:
        try:
            wid = reservation["return_warehouse_id"]
            loc = reservation["return_location"]
        except (KeyError, IndexError, TypeError):
            wid = None
            loc = None
    if not wid:
        eq["location"] = _pdf_location(eq, issue_loc)
        return eq

    ret_wh = con.execute("SELECT name, address FROM warehouses WHERE id=?",
                         (wid,)).fetchone()
    if not ret_wh:
        eq["location"] = _pdf_location(eq, loc or issue_loc)
        return eq
    eq["warehouse_name"] = ret_wh["name"]
    eq["warehouse_address"] = ret_wh["address"] or ""
    eq["location"] = _pdf_location(eq, loc)
    return eq


def _pdf_for_rids(con, kind, rids, ret_wid=None, ret_loc=None):
    rows = _selected_reservations(con, rids)
    if not rows:
        return None
    prefix = "WZ" if kind == "wydanie" else "PZ"
    if len(rows) == 1:
        r = rows[0]
        eq = _eq_for_pdf(con, r["equipment_id"], kind=kind, reservation=r,
                         ret_wid=ret_wid, ret_loc=ret_loc)
        op = None
        if kind == "wydanie" and r["issued_by"]:
            op = con.execute("SELECT * FROM users WHERE id=?", (r["issued_by"],)).fetchone()
        elif kind == "przyjecie" and r["returned_by"]:
            op = con.execute("SELECT * FROM users WHERE id=?", (r["returned_by"],)).fetchone()
        photos = _photos_for_pdf(con, r["equipment_id"],
                                 eq["photo"] if eq else None)
        buf = protocol_pdf(kind, r, eq, display_name(r),
                           display_name(op) if op else None, photos=photos)
        name = f"{prefix}_{r['code']}_{local_now():%Y%m%d_%H%M}.pdf"
    else:
        # Zbiorcze wydanie/przyjęcie – jeden PDF tabelaryczny (jak wcześniej)
        enriched = []
        for r in rows:
            row = dict(r)
            if not row.get("photo"):
                photos = equipment_photo_list(con, r["equipment_id"])
                if photos:
                    row["photo"] = photos[0]
            enriched.append(row)
        buf = group_pdf(kind, enriched)
        name = f"{prefix}_zbiorczy_{local_now():%Y%m%d_%H%M}.pdf"
    return buf, name


@app.route("/reservations/bulk/<action>", methods=["GET", "POST"])
@login_required
def bulk_action(action):
    if request.method == "GET":
        flash("Operacja zbiorcza wymaga ponownego zatwierdzenia z listy.", "error")
        return redirect(url_for("reservations"))
    if action not in ("issue", "return"):
        abort(404)
    try:
        return _bulk_action_post(action)
    except Exception as exc:
        flash(f"Operacja zbiorcza nie powiodła się: {exc}. Spróbuj ponownie (ew. po jednej pozycji).", "error")
        return redirect(url_for("reservations"))


def _bulk_action_post(action):
    con = get_db()
    rows = _selected_reservations(con, request.form.getlist("rid"))
    # return: item_qty_ok/damage/dispose_<id> (mieszany zwrot + utylizacja w jednym kroku)
    if action == "return":
        need_wh = False
        missing_dmg = []
        missing_disp = []
        bad_qty = []
        missing_loc = []
        for r in rows:
            parsed = _parse_return_ok_damage_qty(request.form, r)
            if parsed is None:
                bad_qty.append(r["code"])
                continue
            ok, dmg, disp = parsed
            if ok or dmg:
                need_wh = True
                if not _return_location_for(request.form, r):
                    missing_loc.append(r["code"])
            if dmg and not ((request.form.get(f"item_damage_notes_{r['id']}") or "").strip()
                            or (request.form.get("damage_notes") or "").strip()):
                missing_dmg.append(r["code"])
            if disp and not _dispose_notes_for(request.form, r):
                missing_disp.append(r["code"])
        if bad_qty:
            con.close()
            flash("Podaj poprawne liczby (sprawne + uszkodzone + utylizacja) dla: "
                  + ", ".join(bad_qty) + ".", "error")
            return redirect(url_for("reservations"))
        if missing_dmg:
            con.close()
            flash("Podaj opis uszkodzenia przy: " + ", ".join(missing_dmg) + ".", "error")
            return redirect(url_for("reservations"))
        if missing_disp:
            con.close()
            flash("Podaj powód utylizacji przy: " + ", ".join(missing_disp) + ".", "error")
            return redirect(url_for("reservations"))
        if need_wh and not (request.form.get("return_warehouse_id") or "").strip().isdigit():
            con.close()
            flash("Wybierz magazyn przyjęcia dla pozycji wracających.", "error")
            return redirect(url_for("reservations"))
        if need_wh and not (request.form.get("returner") or "").strip():
            con.close()
            flash("Wybierz, kto oddaje towar.", "error")
            return redirect(url_for("reservations"))
        if missing_loc:
            con.close()
            flash("Potwierdź miejsce w magazynie dla: " + ", ".join(missing_loc) + ".", "error")
            return redirect(url_for("reservations"))
    xbs_meta = None
    if action == "issue":
        xbs_meta = _parse_xbs_form(request.form)
        if xbs_meta is False:
            con.close()
            flash("Awizacja XBS: podaj datę dostawy.", "error")
            return redirect(url_for("reservations"))
    done_ids = []
    returned_ids = []
    disposed_ids = []
    denied = 0
    stock_err = []
    for r in rows:
        if not can_manage_reservation(r):
            denied += 1
            continue
        if action == "issue":
            is_perm = bool(r["permanent"]) if "permanent" in r.keys() else False
            if is_perm and r["status"] == "rezerwacja":
                eq = con.execute("SELECT quantity FROM equipment WHERE id=?",
                                 (r["equipment_id"],)).fetchone()
                if not eq or eq["quantity"] < r["quantity"]:
                    stock_err.append(r["code"])
                    continue
            if _apply_issue(con, r, session["user_id"]):
                done_ids.append(r["id"])
        elif action == "return":
            legacy_dmg = (request.form.get(f"item_action_{r['id']}") or "") == "damage"
            has_split = (
                request.form.get(f"item_qty_ok_{r['id']}") is not None
                or request.form.get(f"item_qty_damage_{r['id']}") is not None
                or request.form.get(f"item_qty_dispose_{r['id']}") is not None
            )
            done = _apply_return(con, r, session["user_id"], request.form,
                                 files=request.files,
                                 force_damage=(legacy_dmg and not has_split))
            if done:
                for did in (done if isinstance(done, list) else [r["id"]]):
                    st = con.execute("SELECT status FROM reservations WHERE id=?",
                                     (did,)).fetchone()
                    if st and st["status"] == "utylizacja":
                        disposed_ids.append(did)
                    else:
                        returned_ids.append(did)
                    done_ids.append(did)
            elif r["status"] == "wydane":
                stock_err.append(r["code"])
    if action == "issue" and done_ids and xbs_meta:
        _save_xbs_on_reservations(con, done_ids, xbs_meta)
    con.commit()
    if stock_err:
        flash(f"Brak stanu magazynowego / nie udało się dla: {', '.join(stock_err)}.", "error")
    if denied and not done_ids:
        con.close()
        flash("Brak uprawnień do rezerwacji spoza Twojego działu.", "error")
        return redirect(url_for("reservations"))
    if not done_ids:
        con.close()
        if not stock_err:
            flash("Brak pozycji o odpowiednim statusie.", "error")
        return redirect(url_for("reservations"))
    con.close()
    if action == "return":
        parts = []
        if returned_ids:
            parts.append(f"zwrot: {len(returned_ids)}")
        if disposed_ids:
            parts.append(f"utylizacja: {len(disposed_ids)}")
        msg = "Zapisano (" + ", ".join(parts) + ")."
        # PDF obejmuje zwroty i utylizacje z tej operacji
        pdf_ids = returned_ids + disposed_ids
        flash(msg + (" Dokumenty pobiorą się za chwilę (lub użyj linku na liście)." if pdf_ids else ""), "ok")
        if pdf_ids:
            ret_wid = (request.form.get("return_warehouse_id") or "").strip()
            ret_loc = (request.form.get("return_location") or "").strip()
            return redirect(url_for("reservations", auto_pdf="przyjecie", rid=pdf_ids,
                                    ret_wid=ret_wid, ret_loc=ret_loc))
        return redirect(url_for("reservations"))
    return _redirect_after_issue(done_ids, xbs_meta)


@app.route("/reservations/pdf-group/<kind>")
@login_required
def pdf_group(kind):
    """Tylko ponowne pobranie PDF – bez zmiany statusu (ten sam układ co auto WZ/PZ)."""
    if kind not in ("wydanie", "przyjecie"):
        abort(404)
    con = get_db()
    result = _pdf_for_rids(con, kind, request.args.getlist("rid"))
    con.close()
    if not result:
        flash("Zaznacz przynajmniej jedną rezerwację.", "error")
        return redirect(url_for("reservations"))
    buf, name = result
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                     download_name=name)


def _get_reservation(con, rid):
    r = con.execute(
        """SELECT r.*, u.username, u.first_name, u.last_name, u.department AS owner_department,
                  e.code, e.name
           FROM reservations r
           JOIN users u ON u.id=r.user_id JOIN equipment e ON e.id=r.equipment_id
           WHERE r.id=?""", (rid,)).fetchone()
    if not r:
        abort(404)
    return r


@app.route("/reservations/purge-all", methods=["POST"])
@login_required
@admin_required
def purge_all_reservations():
    """Usuwa wszystkie rezerwacje. Przy wydaniach trwałych / utylizacji oddaje sztuki na stan."""
    if request.form.get("confirm") != "USUN":
        flash("Potwierdzenie nieprawidłowe.", "error")
        return redirect(url_for("reservations"))
    con = get_db()
    rows = con.execute(
        """SELECT id, equipment_id, quantity, status FROM reservations
           WHERE status IN ('wydane trwale', 'utylizacja')"""
    ).fetchall()
    for r in rows:
        con.execute(
            "UPDATE equipment SET quantity = quantity + ? WHERE id=?",
            (r["quantity"], r["equipment_id"]),
        )
        eq = con.execute("SELECT warehouse_id, IFNULL(location,'') loc FROM equipment WHERE id=?",
                         (r["equipment_id"],)).fetchone()
        if eq:
            add_equipment_stock(con, r["equipment_id"], eq["warehouse_id"], eq["loc"], r["quantity"])
            sync_equipment_from_stock(con, r["equipment_id"], keep_total=True)
    n = con.execute("SELECT COUNT(*) c FROM reservations").fetchone()["c"]
    con.execute("DELETE FROM reservations")
    con.commit()
    flash(f"Usunięto wszystkie rezerwacje ({n}).", "ok")
    if request.accept_mimetypes.best == "application/json" or request.args.get("json"):
        return jsonify(ok=True, deleted=n)
    return redirect(url_for("reservations"))


@app.route("/reservations/<int:rid>/issue", methods=["GET", "POST"])
@login_required
def issue(rid):
    if request.method == "GET":
        # Odświeżenie / wejście GET po POST nie może kończyć się 405 Method Not Allowed
        flash("Aby wydać sprzęt, użyj przycisku „Wydaj” na liście rezerwacji.", "error")
        return redirect(url_for("reservations"))
    try:
        return _issue_post(rid)
    except Exception as exc:
        flash(f"Wydanie nie powiodło się: {exc}. Spróbuj ponownie.", "error")
        return redirect(url_for("reservations"))


def _issue_post(rid):
    con = get_db()
    r = _get_reservation(con, rid)
    if not can_manage_reservation(r):
        con.close()
        flash("Możesz zarządzać tylko rezerwacjami swojego działu.", "error")
        return redirect(request.referrer or url_for("reservations"))
    if r["status"] != "rezerwacja":
        con.close()
        flash("Można wydać tylko aktywną rezerwację.", "error")
        return redirect(request.referrer or url_for("reservations"))
    new_to = (request.form.get("date_to") or "").strip() or r["date_to"]
    if new_to != r["date_to"]:
        ok_to, to_err = _validate_return_date(con, r, new_to)
        if not ok_to:
            con.close()
            flash(to_err, "error")
            return redirect(request.referrer or url_for("reservations"))
    is_perm = bool(r["permanent"]) if "permanent" in r.keys() else False
    xbs_meta = _parse_xbs_form(request.form)
    if xbs_meta is False:
        con.close()
        flash("Awizacja XBS: podaj datę dostawy.", "error")
        return redirect(request.referrer or url_for("reservations"))
    if is_perm:
        eq = con.execute("SELECT quantity FROM equipment WHERE id=?",
                         (r["equipment_id"],)).fetchone()
        if not eq or eq["quantity"] < r["quantity"]:
            con.close()
            flash(f"Brak stanu magazynowego ({r['quantity']} szt. wymagane).", "error")
            return redirect(request.referrer or url_for("reservations"))
    if not _apply_issue(con, r, session["user_id"], date_to=new_to):
        con.close()
        flash("Nie udało się wydać rezerwacji.", "error")
        return redirect(request.referrer or url_for("reservations"))
    if xbs_meta:
        _save_xbs_on_reservations(con, [rid], xbs_meta)
    con.commit()
    archived = False
    if is_perm:
        eq_left = con.execute(
            "SELECT IFNULL(archived,0) a FROM equipment WHERE id=?",
            (r["equipment_id"],)).fetchone()
        archived = bool(eq_left and eq_left["a"])
    con.close()
    if is_perm and archived:
        msg = ("Oznaczono jako wydane trwale – sprzęt przeniesiony do archiwum "
               "(stan 0). Przywrócenie: tylko admin. Dokumenty pobiorą się za chwilę.")
    elif is_perm:
        msg = "Oznaczono jako wydane trwale (towar nie wraca). Dokumenty pobiorą się za chwilę."
    else:
        msg = "Oznaczono jako wydane. Dokumenty pobiorą się za chwilę."
    if new_to != r["date_to"]:
        msg = f"Termin zwrotu {new_to}. " + msg
    flash(msg, "ok")
    if xbs_meta:
        _store_xbs_session([rid], xbs_meta)
        flash("Awizacja XBS: Excel pobierze się zaraz po PDF (lub link w kolumnie Dokumenty).", "ok")
        return redirect(url_for("reservations", auto_pdf="wydanie", rid=rid, xbs_hint="1"))
    return redirect(url_for("reservations", auto_pdf="wydanie", rid=rid))


@app.route("/reservations/<int:rid>/return", methods=["POST"])
@login_required
def return_item(rid):
    con = get_db()
    r = _get_reservation(con, rid)
    if not can_manage_reservation(r):
        con.close()
        flash("Możesz zarządzać tylko rezerwacjami swojego działu.", "error")
        return redirect(request.referrer or url_for("reservations"))
    if r["status"] != "wydane":
        flash("Można zwrócić tylko wydany sprzęt.", "error")
        con.close()
        return redirect(request.referrer or url_for("reservations"))
    parsed = _parse_return_ok_damage_qty(request.form, r)
    if parsed is None:
        flash(f"Podaj poprawne liczby sztuk (sprawne + uszkodzone + utylizacja = 1–{r['quantity']}).",
              "error")
        con.close()
        return redirect(request.referrer or url_for("reservations"))
    ok_qty, dmg_qty, disp_qty = parsed
    if dmg_qty and not (request.form.get("damage_notes") or "").strip():
        flash("Podaj opis uszkodzenia.", "error")
        con.close()
        return redirect(request.referrer or url_for("reservations"))
    if disp_qty and not _dispose_notes_for(request.form, r):
        flash("Podaj powód utylizacji.", "error")
        con.close()
        return redirect(request.referrer or url_for("reservations"))
    if (ok_qty or dmg_qty) and not _return_form_valid(request.form):
        flash("Wypełnij magazyn przyjęcia, miejsce w magazynie i kto oddaje towar.", "error")
        con.close()
        return redirect(request.referrer or url_for("reservations"))
    ret_wid = (request.form.get("return_warehouse_id") or "").strip()
    ret_loc = (request.form.get("return_location") or "").strip()
    done = _apply_return(con, r, session["user_id"], request.form, files=request.files)
    if done:
        con.commit()
        con.close()
        parts = []
        if ok_qty:
            parts.append(f"{ok_qty} sprawne")
        if dmg_qty:
            parts.append(f"{dmg_qty} uszkodzone")
        if disp_qty:
            parts.append(f"{disp_qty} utylizacja")
        flash("Zapisano (" + ", ".join(parts) + "). Pobieranie PDF…", "ok")
        q = [("auto_pdf", "przyjecie"), ("ret_wid", ret_wid), ("ret_loc", ret_loc)]
        for did in done:
            q.append(("rid", str(did)))
        return redirect(url_for("reservations") + "?" + urlencode(q))
    con.close()
    flash("Nie udało się przyjąć zwrotu / utylizacji (sprawdź stan magazynowy).", "error")
    return redirect(request.referrer or url_for("reservations"))


@app.route("/reservations/auto-pdf/<kind>")
@login_required
def auto_pdf(kind):
    """Pobiera WZ/PZ po wydaniu/zwrocie (wywoływane automatycznie z listy)."""
    if kind not in ("wydanie", "przyjecie"):
        abort(404)
    try:
        con = get_db()
        ret_wid = request.args.get("ret_wid", "").strip()
        ret_loc = request.args.get("ret_loc", "").strip()
        result = _pdf_for_rids(con, kind, request.args.getlist("rid"),
                               ret_wid=ret_wid or None, ret_loc=ret_loc or None)
        con.close()
    except Exception as exc:
        flash(f"Nie udało się wygenerować PDF: {exc}. Spróbuj ponownie z listy rezerwacji.", "error")
        return redirect(url_for("reservations"))
    if not result:
        flash("Brak danych do PDF.", "error")
        return redirect(url_for("reservations"))
    buf, name = result
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                     download_name=name)


@app.route("/reservations/auto-xbs")
@login_required
def auto_xbs():
    """Pobiera awizację XBS (osobne żądanie – nie blokuje wydania)."""
    rids, meta = _xbs_payload_for_download(request.args.getlist("rid"))
    if not rids or not meta:
        flash("Brak danych awizacji XBS.", "error")
        return redirect(url_for("reservations"))
    session.pop("xbs_awizacja", None)
    con = get_db()
    try:
        result = _xbs_file_response(con, rids, meta)
    except Exception as exc:
        con.close()
        flash(f"Nie udało się wygenerować awizacji XBS: {exc}", "error")
        return redirect(url_for("reservations"))
    con.close()
    if not result:
        flash("Brak pozycji do awizacji XBS.", "error")
        return redirect(url_for("reservations"))
    buf, name = result
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=name,
    )


@app.route("/reservations/<int:rid>/xbs")
@login_required
def reservation_xbs(rid):
    """Ponowne pobranie awizacji XBS (jak PDF WZ/PZ) – pełna paczka batcha jeśli była zbiorcza."""
    con = get_db()
    r = _get_reservation(con, rid)
    meta = _xbs_meta_from_row(r)
    if not meta:
        con.close()
        flash("Brak zapisanej awizacji XBS dla tej rezerwacji.", "error")
        return redirect(url_for("reservations"))
    batch = None
    try:
        batch = r["xbs_batch_id"]
    except (KeyError, IndexError, TypeError):
        batch = None
    if batch:
        rows = con.execute(
            "SELECT id FROM reservations WHERE xbs_batch_id=? ORDER BY id",
            (batch,),
        ).fetchall()
        rids = [row["id"] for row in rows]
    else:
        rids = [rid]
    try:
        result = _xbs_file_response(con, rids, meta)
    except Exception as exc:
        con.close()
        flash(f"Nie udało się wygenerować awizacji XBS: {exc}", "error")
        return redirect(url_for("reservations"))
    con.close()
    if not result:
        flash("Brak pozycji do awizacji XBS.", "error")
        return redirect(url_for("reservations"))
    buf, name = result
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=name,
    )


@app.route("/reservations/<int:rid>/cancel", methods=["POST"])
@login_required
def cancel(rid):
    con = get_db()
    r = _get_reservation(con, rid)
    if r["status"] != "rezerwacja":
        flash("Można anulować tylko aktywną rezerwację.", "error")
    elif r["user_id"] != session["user_id"] and session["role"] != "admin":
        flash("Możesz anulować tylko własne rezerwacje.", "error")
    else:
        con.execute("UPDATE reservations SET status='anulowana' WHERE id=?", (rid,))
        con.commit()
        flash("Rezerwacja anulowana.", "ok")
    con.close()
    return redirect(request.referrer or url_for("reservations"))


@app.route("/reservations/<int:rid>/change-return", methods=["POST"])
@login_required
def change_return_date(rid):
    """Zmiana terminu zwrotu po wydaniu – z kontrolą kolizji z kolejnymi rezerwacjami."""
    wants_json = (
        request.headers.get("X-Requested-With") == "XMLHttpRequest"
        or "application/json" in (request.headers.get("Accept") or "")
    )

    def respond(ok, message, status=200):
        if wants_json:
            return jsonify(ok=ok, message=message), (status if not ok else 200)
        flash(message, "ok" if ok else "error")
        return redirect(request.referrer or url_for("reservations"))

    con = get_db()
    r = _get_reservation(con, rid)
    if not can_manage_reservation(r):
        con.close()
        return respond(False, "Możesz zarządzać tylko rezerwacjami swojego działu.", 403)
    if r["status"] != "wydane":
        con.close()
        return respond(False, "Termin zwrotu można zmienić tylko dla wydanego sprzętu.", 400)

    new_to = (request.form.get("date_to") or "").strip()
    if new_to == r["date_to"]:
        con.close()
        return respond(True, "Termin zwrotu bez zmian.")
    ok_to, to_err = _validate_return_date(con, r, new_to)
    if not ok_to:
        con.close()
        return respond(False, to_err, 409)

    eq = con.execute("SELECT code FROM equipment WHERE id=?", (r["equipment_id"],)).fetchone()
    old_to = r["date_to"]
    con.execute("UPDATE reservations SET date_to=? WHERE id=?", (new_to, rid))
    con.commit()
    con.close()
    return respond(True, f"Zmieniono termin zwrotu {eq['code']}: {old_to} → {new_to}.")


@app.route("/reservations/<int:rid>/pdf/<kind>")
@login_required
def reservation_pdf(rid, kind):
    if kind not in ("wydanie", "przyjecie"):
        abort(404)
    con = get_db()
    r = _get_reservation(con, rid)
    eq = _eq_for_pdf(con, r["equipment_id"], kind=kind, reservation=r)
    photos = _photos_for_pdf(con, r["equipment_id"], eq["photo"] if eq else None)
    op_id = r["issued_by"] if kind == "wydanie" else r["returned_by"]
    op = None
    if op_id:
        row = con.execute("SELECT * FROM users WHERE id=?", (op_id,)).fetchone()
        op = display_name(row) if row else None
    con.close()
    buf = protocol_pdf(kind, r, eq, display_name(r), op, photos=photos)
    prefix = "WZ" if kind == "wydanie" else "PZ"
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                     download_name=f"{prefix}_{eq['code']}_{rid}.pdf")


# ---------- archiwum sprzętu (admin) ----------

@app.route("/archive")
@login_required
@admin_required
def archive_index():
    q = request.args.get("q", "").strip()
    con = get_db()
    sql = """SELECT e.*, w.name AS warehouse_name FROM equipment e
             LEFT JOIN warehouses w ON w.id=e.warehouse_id
             WHERE IFNULL(e.archived,0)=1"""
    params = []
    if q:
        sql += " AND (e.code LIKE ? OR e.name LIKE ?)"
        params += [f"%{q}%", f"%{q}%"]
    sql += " ORDER BY e.archived_at DESC, e.code"
    items = con.execute(sql, params).fetchall()
    warehouses = active_warehouses(con)
    con.close()
    return render_template("archive.html", items=items, q=q, warehouses=warehouses)


@app.route("/equipment/<int:eid>/restore", methods=["POST"])
@login_required
@admin_required
def equipment_restore(eid):
    con = get_db()
    eq = con.execute("SELECT * FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        con.close()
        abort(404)
    if not (eq["archived"] if "archived" in eq.keys() else 0):
        con.close()
        flash("Ten sprzęt nie jest w archiwum.", "error")
        return redirect(url_for("archive_index"))
    try:
        qty = max(1, int(request.form.get("quantity") or 1))
    except (TypeError, ValueError):
        qty = 1
    wid_raw = request.form.get("warehouse_id", "")
    wid = int(wid_raw) if str(wid_raw).isdigit() else None
    location = (request.form.get("location") or "").strip()
    con.execute(
        """UPDATE equipment SET archived=0, archived_at=NULL, quantity=?,
           damaged_quantity=0, condition='sprawny',
           warehouse_id=?, location=? WHERE id=?""",
        (qty, wid, location, eid))
    replace_equipment_stock(con, eid, wid, location, qty)
    con.commit()
    con.close()
    flash(f"Przywrócono {eq['code']} na stan ({qty} szt.).", "ok")
    return redirect(url_for("equipment_detail", eid=eid))


# ---------- API: autouzupełnianie adresatów ----------

@app.route("/api/recipients")
@login_required
def api_recipients():
    con = get_db()
    rows = recent_recipients(con, limit=100)
    con.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/recipients/<int:rid>/delete", methods=["POST"])
@login_required
@admin_required
def api_recipient_delete(rid):
    """Usuwa adresata ze słownika ostatnich podpowiedzi."""
    con = get_db()
    con.execute("DELETE FROM recipients WHERE id=?", (rid,))
    con.commit()
    con.close()
    return jsonify({"ok": True, "id": rid})


# ---------- słowniki (admin): magazyny + podwykonawcy ----------

@app.route("/dictionaries")
@login_required
@admin_required
def dictionaries():
    con = get_db()
    warehouses = con.execute("SELECT * FROM warehouses ORDER BY name").fetchall()
    partners = con.execute("SELECT * FROM logistics_partners ORDER BY name").fetchall()
    departments = con.execute("SELECT * FROM departments ORDER BY name").fetchall()
    con.close()
    return render_template("dictionaries.html", warehouses=warehouses, partners=partners,
                           departments=departments)


@app.route("/warehouses/add", methods=["POST"])
@login_required
@admin_required
def warehouse_add():
    con = get_db()
    try:
        con.execute("INSERT INTO warehouses (name, address) VALUES (?,?)",
                    (request.form["name"].strip(), request.form.get("address", "").strip()))
        con.commit()
        flash("Magazyn dodany.", "ok")
    except Exception:
        flash("Magazyn o tej nazwie już istnieje.", "error")
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/warehouses/<int:wid>/toggle", methods=["POST"])
@login_required
@admin_required
def warehouse_toggle(wid):
    con = get_db()
    con.execute("UPDATE warehouses SET active = 1 - active WHERE id=?", (wid,))
    con.commit()
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/warehouses/<int:wid>/edit", methods=["POST"])
@login_required
@admin_required
def warehouse_edit(wid):
    con = get_db()
    try:
        con.execute("UPDATE warehouses SET name=?, address=? WHERE id=?",
                    (request.form["name"].strip(), request.form.get("address", "").strip(), wid))
        con.commit()
        flash("Magazyn zaktualizowany.", "ok")
    except Exception:
        flash("Magazyn o tej nazwie już istnieje.", "error")
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/partners/add", methods=["POST"])
@login_required
@admin_required
def partner_add():
    con = get_db()
    try:
        con.execute("INSERT INTO logistics_partners (name, phone, email) VALUES (?,?,?)",
                    (request.form["name"].strip(), request.form.get("phone", "").strip(),
                     request.form.get("email", "").strip()))
        con.commit()
        flash("Podwykonawca dodany.", "ok")
    except Exception:
        flash("Podwykonawca o tej nazwie już istnieje.", "error")
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/partners/<int:pid>/toggle", methods=["POST"])
@login_required
@admin_required
def partner_toggle(pid):
    con = get_db()
    con.execute("UPDATE logistics_partners SET active = 1 - active WHERE id=?", (pid,))
    con.commit()
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/partners/<int:pid>/edit", methods=["POST"])
@login_required
@admin_required
def partner_edit(pid):
    con = get_db()
    try:
        con.execute("UPDATE logistics_partners SET name=?, phone=?, email=? WHERE id=?",
                    (request.form["name"].strip(), request.form.get("phone", "").strip(),
                     request.form.get("email", "").strip(), pid))
        con.commit()
        flash("Podwykonawca zaktualizowany.", "ok")
    except Exception:
        flash("Podwykonawca o tej nazwie już istnieje.", "error")
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/departments/add", methods=["POST"])
@login_required
@admin_required
def department_add():
    con = get_db()
    try:
        con.execute("INSERT INTO departments (name) VALUES (?)",
                    (request.form["name"].strip(),))
        con.commit()
        flash("Dział dodany.", "ok")
    except Exception:
        flash("Dział o tej nazwie już istnieje.", "error")
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/departments/<int:did>/toggle", methods=["POST"])
@login_required
@admin_required
def department_toggle(did):
    con = get_db()
    con.execute("UPDATE departments SET active = 1 - active WHERE id=?", (did,))
    con.commit()
    con.close()
    return redirect(url_for("dictionaries"))


@app.route("/departments/<int:did>/edit", methods=["POST"])
@login_required
@admin_required
def department_edit(did):
    con = get_db()
    name = request.form["name"].strip()
    old = con.execute("SELECT name FROM departments WHERE id=?", (did,)).fetchone()
    try:
        con.execute("UPDATE departments SET name=? WHERE id=?", (name, did))
        if old and old["name"] != name:
            con.execute("UPDATE users SET department=? WHERE department=?", (name, old["name"]))
        con.commit()
        flash("Dział zaktualizowany.", "ok")
    except Exception:
        flash("Dział o tej nazwie już istnieje.", "error")
    con.close()
    return redirect(url_for("dictionaries"))


# ---------- użytkownicy (admin) ----------

@app.route("/users", methods=["GET", "POST"])
@login_required
@admin_required
def users():
    con = get_db()
    f_dept = request.args.get("department", "").strip()
    if request.method == "POST":
        fn = request.form.get("first_name", "").strip()
        ln = request.form.get("last_name", "").strip()
        dept = request.form.get("department", "").strip()
        if not fn or not ln:
            flash("Imię i nazwisko są wymagane – każde konto to konkretny PM.", "error")
        else:
            pw_err = password_policy_error(request.form.get("password", ""))
            if pw_err:
                flash(pw_err, "error")
            else:
                try:
                    con.execute("""INSERT INTO users (username, password_hash, role,
                                   first_name, last_name, department) VALUES (?,?,?,?,?,?)""",
                                (request.form["username"].strip(),
                                 generate_password_hash(request.form["password"], method="pbkdf2:sha256"),
                                 request.form.get("role", "user"), fn, ln, dept or None))
                    con.commit()
                    flash("Użytkownik dodany.", "ok")
                except Exception:
                    flash("Taki login już istnieje.", "error")
    sql = "SELECT * FROM users"
    params = []
    if f_dept == "__none__":
        sql += " WHERE IFNULL(department,'')=''"
    elif f_dept:
        sql += " WHERE department=?"
        params.append(f_dept)
    sql += " ORDER BY IFNULL(department, 'żżż'), last_name, username"
    rows = con.execute(sql, params).fetchall()
    departments = active_departments(con)
    con.close()
    return render_template("users.html", rows=rows, dn=display_name,
                           departments=departments, f_dept=f_dept)


@app.route("/users/<int:uid>/toggle", methods=["POST"])
@login_required
@admin_required
def user_toggle(uid):
    if uid == session["user_id"]:
        flash("Nie możesz dezaktywować własnego konta.", "error")
    else:
        con = get_db()
        con.execute("UPDATE users SET active = 1 - active WHERE id=?", (uid,))
        con.commit()
        con.close()
    return redirect(url_for("users"))


@app.route("/users/<int:uid>/password", methods=["POST"])
@login_required
@admin_required
def user_password(uid):
    pw = request.form["password"]
    pw_err = password_policy_error(pw)
    if pw_err:
        flash(pw_err, "error")
    else:
        con = get_db()
        con.execute("UPDATE users SET password_hash=? WHERE id=?",
                    (generate_password_hash(pw, method="pbkdf2:sha256"), uid))
        con.commit()
        con.close()
        flash("Hasło zmienione.", "ok")
    return redirect(url_for("users"))


@app.route("/users/<int:uid>/name", methods=["POST"])
@login_required
@admin_required
def user_name(uid):
    fn = request.form.get("first_name", "").strip()
    ln = request.form.get("last_name", "").strip()
    dept = request.form.get("department", "").strip()
    username = request.form.get("username", "").strip()
    role = request.form.get("role", "user").strip()
    if role not in ("user", "admin"):
        role = "user"
    if not fn or not ln:
        flash("Imię i nazwisko są wymagane.", "error")
    elif not username:
        flash("Login jest wymagany.", "error")
    elif uid == session["user_id"] and role != "admin":
        flash("Nie możesz odebrać sobie roli admin.", "error")
    else:
        con = get_db()
        try:
            con.execute(
                """UPDATE users SET username=?, first_name=?, last_name=?,
                   department=?, role=? WHERE id=?""",
                (username, fn, ln, dept or None, role, uid))
            con.commit()
            if uid == session["user_id"]:
                session["username"] = username
                session["department"] = dept
                session["full_name"] = f"{fn} {ln}".strip()
                session["role"] = role
            flash("Dane zaktualizowane.", "ok")
        except Exception:
            flash("Taki login już istnieje.", "error")
        con.close()
    return redirect(url_for("users"))


def _catalog_import_flow(catalog="main", endpoint="catalog_import"):
    """Wspólna logika importu Excel+ZIP dla katalogu main lub tcl."""
    status_name = "catalog_import_status.json" if catalog == "main" else "tcl_import_status.json"
    work_name = "catalog_import_work" if catalog == "main" else "tcl_import_work"
    status_path = Path(__file__).parent / "data" / status_name
    result = None
    if status_path.exists():
        try:
            import json
            result = json.loads(status_path.read_text(encoding="utf-8"))
        except Exception:
            result = None

    if request.method == "POST":
        xlsx = request.files.get("xlsx")
        photos_zip = request.files.get("photos_zip")
        do_update = bool(request.form.get("update"))
        also_new = bool(request.form.get("also_new_codes"))
        if not xlsx or not xlsx.filename:
            flash("Wybierz plik Excel (.xlsx).", "error")
            return redirect(url_for(endpoint))
        if not xlsx.filename.lower().endswith(".xlsx"):
            flash("Plik katalogu musi mieć rozszerzenie .xlsx.", "error")
            return redirect(url_for(endpoint))

        work = Path(__file__).parent / "data" / work_name
        if work.exists():
            shutil.rmtree(work, ignore_errors=True)
        work.mkdir(parents=True, exist_ok=True)
        xlsx_path = work / "katalog.xlsx"
        xlsx.save(xlsx_path)
        zip_path = None
        if photos_zip and photos_zip.filename:
            zip_path = work / "zdjecia.zip"
            photos_zip.save(zip_path)

        import json
        import threading

        status_path.write_text(
            json.dumps({"running": True, "messages": ["Import w toku…"]}, ensure_ascii=False),
            encoding="utf-8",
        )

        def job():
            messages = []
            try:
                photos_dir = None
                if zip_path and zip_path.exists():
                    extract_dir = work / "extracted"
                    extract_dir.mkdir(exist_ok=True)
                    try:
                        zf = zipfile.ZipFile(zip_path, "r", metadata_encoding="utf-8")
                    except TypeError:
                        zf = zipfile.ZipFile(zip_path, "r")
                    with zf:
                        zf.extractall(extract_dir)

                    img_ext = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
                    best, best_n = None, -1
                    for c in [extract_dir, *extract_dir.rglob("*")]:
                        if not c.is_dir() or "__MACOSX" in c.parts:
                            continue
                        n = sum(
                            1 for f in c.iterdir()
                            if f.is_file() and f.suffix.lower() in img_ext
                        )
                        if n > best_n:
                            best, best_n = c, n
                    photos_dir = best if best_n > 0 else None
                    if not photos_dir:
                        raise RuntimeError("ZIP nie zawiera folderu ze zdjęciami.")
                    messages.append(f"Folder zdjęć: {photos_dir.name} ({best_n} plików)")

                r1 = run_import(
                    xlsx_path, photos_dir=photos_dir, sheet="Import",
                    update=do_update, log=messages, catalog=catalog,
                )
                totals = dict(r1)
                if also_new:
                    r2 = run_import(
                        xlsx_path, photos_dir=photos_dir, sheet="Nowe kody",
                        update=False, log=messages, catalog=catalog,
                    )
                    totals = {
                        "added": r1["added"] + r2["added"],
                        "updated": r1["updated"] + r2["updated"],
                        "skipped": r1["skipped"] + r2["skipped"],
                        "no_photo": r1["no_photo"] + r2["no_photo"],
                        "messages": messages,
                    }
                totals["running"] = False
                totals["messages"] = messages
                status_path.write_text(
                    json.dumps(totals, ensure_ascii=False), encoding="utf-8"
                )
            except Exception as e:
                messages.append(f"Błąd: {e}")
                status_path.write_text(
                    json.dumps(
                        {"running": False, "error": str(e), "messages": messages,
                         "added": 0, "updated": 0, "skipped": 0, "no_photo": 0},
                        ensure_ascii=False,
                    ),
                    encoding="utf-8",
                )
            finally:
                shutil.rmtree(work, ignore_errors=True)

        threading.Thread(target=job, daemon=True).start()
        flash("Import uruchomiony w tle. Odśwież tę stronę za 1–3 minuty.", "ok")
        return redirect(url_for(endpoint))

    return render_template(
        "catalog_import.html", result=result, catalog=catalog, endpoint=endpoint)


@app.route("/catalog-import", methods=["GET", "POST"])
@login_required
@admin_required
def catalog_import():
    """Import katalogu głównego z Excela + ZIP (tylko admin)."""
    return _catalog_import_flow("main", "catalog_import")


@app.route("/tcl/import", methods=["GET", "POST"])
@login_required
@tcl_required
def tcl_import():
    """Import katalogu TCL – dział Warrens + admin."""
    return _catalog_import_flow("tcl", "tcl_import")


def _catalog_export_response(catalog="main"):
    """Eksport Excela z miniaturami: zaznaczone (POST) albo wg filtrów (GET)."""
    con = get_db()
    if request.method == "POST":
        raw_ids = request.form.getlist("eid")
        eids = []
        for x in raw_ids:
            try:
                eids.append(int(x))
            except (TypeError, ValueError):
                pass
        if not eids:
            con.close()
            flash("Zaznacz przynajmniej jedną pozycję do eksportu.", "error")
            return redirect(url_for("tcl_index" if catalog == "tcl" else "index"))
        placeholders = ",".join("?" * len(eids))
        rows = con.execute(
            f"""SELECT e.*, w.name AS warehouse_name FROM equipment e
                LEFT JOIN warehouses w ON w.id=e.warehouse_id
                WHERE e.id IN ({placeholders})
                  AND IFNULL(e.catalog,'main')=?
                  AND IFNULL(e.archived,0)=0
                ORDER BY e.code""",
            (*eids, catalog)).fetchall()
        suffix = "zaznaczone"
    else:
        filters = _equipment_list_filters(catalog)
        rows = _fetch_equipment_rows(con, filters)
        suffix = "filtry"
    con.close()
    if not rows:
        flash("Brak pozycji do eksportu.", "error")
        return redirect(url_for("tcl_index" if catalog == "tcl" else "index"))

    buf = build_catalog_miniatures_xlsx(rows, UPLOAD_DIR)
    stamp = local_now().strftime("%Y%m%d_%H%M")
    cat_label = "tcl" if catalog == "tcl" else "katalog"
    name = f"export_{cat_label}_{suffix}_{stamp}.xlsx"
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=name,
    )


@app.route("/catalog-export", methods=["GET", "POST"])
@login_required
def catalog_export():
    """Eksport katalogu głównego (filtry lub zaznaczenie) – Excel z miniaturami."""
    return _catalog_export_response("main")


@app.route("/tcl/export", methods=["GET", "POST"])
@login_required
@tcl_required
def tcl_export():
    """Eksport katalogu TCL (filtry lub zaznaczenie)."""
    return _catalog_export_response("tcl")


@app.route("/admin/export-db")
@login_required
@admin_required
def admin_export_db():
    """Eksport katalogu głównego w formacie importu: Excel + ZIP zdjęć (folder zdjecia/)."""
    stamp = local_now().strftime("%Y%m%d_%H%M")
    con = get_db()
    rows = con.execute(
        """SELECT e.*, w.name AS warehouse_name FROM equipment e
           LEFT JOIN warehouses w ON w.id=e.warehouse_id
           WHERE IFNULL(e.catalog,'main')='main' AND IFNULL(e.archived,0)=0
           ORDER BY e.code"""
    ).fetchall()

    export_rows = []
    photo_copies = []  # (src Path, arcname)
    for r in rows:
        photo_rows = [
            p for p in equipment_photo_rows(con, r["id"])
            if (p["kind"] or "normal") != "repaired"
        ]
        filenames = [p["filename"] for p in photo_rows]
        if not filenames and r["photo"]:
            filenames = [r["photo"]]
        names, copies = export_photo_zip_names(r["code"], filenames, UPLOAD_DIR)
        photo_copies.extend(copies)
        export_rows.append({
            "code": r["code"],
            "name": r["name"],
            "project_number": r["project_number"],
            "dimensions": r["dimensions"],
            "warehouse_name": r["warehouse_name"],
            "location": r["location"],
            "owner": r["owner"],
            "brand": r["brand"],
            "material_type": r["material_type"],
            "condition": r["condition"],
            "quantity": r["quantity"],
            "storage_instructions": r["storage_instructions"],
            "photo_file": ", ".join(names),
        })
    con.close()

    xlsx_buf = build_catalog_import_xlsx(export_rows)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    tmp_path = Path(tmp.name)
    tmp.close()
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"import_katalog_produktow_{stamp}.xlsx", xlsx_buf.getvalue())
            # osobny ZIP ze zdjęciami – jak przy imporcie (xlsx + zdjecia.zip)
            photos_buf = BytesIO()
            with zipfile.ZipFile(photos_buf, "w", zipfile.ZIP_DEFLATED) as pz:
                pz.writestr("zdjecia/", "")
                for src, arc in photo_copies:
                    pz.write(src, arcname=arc)
            zf.writestr(f"zdjecia_{stamp}.zip", photos_buf.getvalue())

        @after_this_request
        def _cleanup(response):
            try:
                tmp_path.unlink(missing_ok=True)
            except OSError:
                pass
            return response

        return send_file(
            tmp_path,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"export_katalog_import_{stamp}.zip",
        )
    except Exception:
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
