import os
import uuid
from datetime import datetime, date
from functools import wraps
from pathlib import Path

from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, send_file, abort, jsonify)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

from db import (get_db, init_db, reserved_qty, display_name, upsert_recipient,
                equipment_photo_list)
from pdf_gen import protocol_pdf, group_pdf

BASE = Path(__file__).parent
UPLOAD_DIR = BASE / "static" / "uploads"
ALLOWED_EXT = {"png", "jpg", "jpeg", "gif", "webp"}
MAX_PHOTOS = 5

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "zmien-mnie-w-produkcji")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB (do 5 zdjęć)

init_db()
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


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


def active_partners(con):
    return con.execute(
        "SELECT * FROM logistics_partners WHERE active=1 ORDER BY name").fetchall()


def active_warehouses(con):
    return con.execute(
        "SELECT * FROM warehouses WHERE active=1 ORDER BY name").fetchall()


def recent_recipients(con, limit=30):
    return con.execute(
        "SELECT * FROM recipients ORDER BY last_used DESC LIMIT ?", (limit,)).fetchall()


def recipient_form_fields(form):
    return dict(
        recipient_name=form.get("recipient_name", "").strip(),
        recipient_contact=form.get("recipient_contact", "").strip(),
        recipient_phone=form.get("recipient_phone", "").strip(),
        recipient_address=form.get("recipient_address", "").strip(),
        recipient_email=form.get("recipient_email", "").strip(),
    )


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
                           full_name=display_name(u))
            default = "dashboard" if u["role"] == "admin" else "index"
            return redirect(request.args.get("next") or url_for(default))
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
    if not check_password_hash(u["password_hash"], request.form.get("current_password", "")):
        flash("Obecne hasło jest nieprawidłowe.", "error")
    elif len(request.form.get("new_password", "")) < 6:
        flash("Nowe hasło musi mieć min. 6 znaków.", "error")
    else:
        con.execute("UPDATE users SET password_hash=? WHERE id=?",
                    (generate_password_hash(request.form["new_password"],
                                            method="pbkdf2:sha256"), u["id"]))
        con.commit()
        flash("Hasło zmienione.", "ok")
    con.close()
    default = "dashboard" if session.get("role") == "admin" else "index"
    return redirect(request.referrer or url_for(default))


# ---------- dashboard ----------

@app.route("/dashboard")
@login_required
@admin_required
def dashboard():
    today = date.today().isoformat()
    con = get_db()
    base_sql = """SELECT r.*, u.username, u.first_name, u.last_name,
                  e.code, e.name, e.location, w.name AS warehouse_name
                  FROM reservations r
                  JOIN users u ON u.id=r.user_id
                  JOIN equipment e ON e.id=r.equipment_id
                  LEFT JOIN warehouses w ON w.id=e.warehouse_id"""
    out_today = con.execute(
        base_sql + " WHERE r.status='rezerwacja' AND r.date_from=? ORDER BY e.code",
        (today,)).fetchall()
    back_today = con.execute(
        base_sql + " WHERE r.status='wydane' AND r.date_to=? ORDER BY e.code",
        (today,)).fetchall()
    overdue = con.execute(
        base_sql + " WHERE r.status='wydane' AND r.date_to<? ORDER BY r.date_to",
        (today,)).fetchall()
    con.close()
    days_overdue = {r["id"]: (date.today() - date.fromisoformat(r["date_to"])).days
                    for r in overdue}
    return render_template("dashboard.html", out_today=out_today,
                           back_today=back_today, overdue=overdue, today=today,
                           days_overdue=days_overdue, dn=display_name)


# ---------- rejestr sprzętu ----------

@app.route("/")
@login_required
def index():
    q = request.args.get("q", "").strip()
    f_project = request.args.get("project", "").strip()
    f_owner = request.args.get("owner", "").strip()
    f_brand = request.args.get("brand", "").strip()
    f_warehouse = request.args.get("warehouse", "").strip()
    f_own = request.args.get("own", "").strip()          # "1" = tylko materiały 00
    f_condition = request.args.get("condition", "").strip()

    con = get_db()
    sql = """SELECT e.*, w.name AS warehouse_name FROM equipment e
             LEFT JOIN warehouses w ON w.id=e.warehouse_id"""
    where, params = [], []
    if q:
        where.append("(e.code LIKE ? OR e.name LIKE ? OR e.location LIKE ?)")
        params += [f"%{q}%"] * 3
    if f_project:
        where.append("e.project_number = ?"); params.append(f_project)
    if f_owner:
        where.append("e.owner = ?"); params.append(f_owner)
    if f_brand:
        where.append("e.brand = ?"); params.append(f_brand)
    if f_warehouse.isdigit():
        where.append("e.warehouse_id = ?"); params.append(int(f_warehouse))
    if f_own == "1":
        where.append("e.material_type = 'wlasny'")
    if f_condition:
        where.append("e.condition = ?"); params.append(f_condition)
    if where:
        sql += " WHERE " + " AND ".join(where)
    items = con.execute(sql + " ORDER BY e.code", params).fetchall()

    # wartości do dropdownów filtrów
    projects = [r[0] for r in con.execute(
        "SELECT DISTINCT project_number FROM equipment WHERE IFNULL(project_number,'')!='' ORDER BY 1")]
    owners = [r[0] for r in con.execute(
        "SELECT DISTINCT owner FROM equipment WHERE IFNULL(owner,'')!='' ORDER BY 1")]
    brands = [r[0] for r in con.execute(
        "SELECT DISTINCT brand FROM equipment WHERE IFNULL(brand,'')!='' ORDER BY 1")]
    warehouses = active_warehouses(con)

    today = date.today().isoformat()
    availability = {
        it["id"]: it["quantity"] - reserved_qty(con, it["id"], today, today)
        for it in items
    }
    con.close()
    return render_template("index.html", items=items, availability=availability,
                           q=q, f_project=f_project, f_owner=f_owner, f_brand=f_brand,
                           f_warehouse=f_warehouse, f_own=f_own, f_condition=f_condition,
                           projects=projects, owners=owners, brands=brands,
                           warehouses=warehouses)


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
        max(1, int(form.get("quantity") or 1)), form["notes"].strip(),
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
@admin_required
def equipment_new():
    con = get_db()
    if request.method == "POST":
        try:
            new_photos = _save_new_photos(request.files)[:MAX_PHOTOS]
            primary = new_photos[0] if new_photos else None
            cur = con.execute(
                f"INSERT INTO equipment ({EQ_COLS}) VALUES ({','.join('?'*16)})",
                _equipment_form_values(request.form, request.files, primary_photo=primary))
            eid = cur.lastrowid
            for i, fn in enumerate(new_photos):
                con.execute(
                    "INSERT INTO equipment_photos (equipment_id, filename, sort_order) VALUES (?,?,?)",
                    (eid, fn, i))
            con.commit()
            flash("Sprzęt dodany.", "ok")
            return redirect(url_for("index"))
        except Exception as e:
            flash(f"Błąd: {'kod już istnieje' if 'UNIQUE' in str(e) else e}", "error")
        finally:
            con.close()
        return redirect(url_for("equipment_new"))
    warehouses = active_warehouses(con)
    con.close()
    return render_template("equipment_form.html", eq=None, warehouses=warehouses,
                           photos=[])


@app.route("/equipment/<int:eid>/edit", methods=["GET", "POST"])
@login_required
@admin_required
def equipment_edit(eid):
    con = get_db()
    eq = con.execute("SELECT * FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        abort(404)
    photos = equipment_photo_list(con, eid)
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
                con.execute(
                    "INSERT INTO equipment_photos (equipment_id, filename, sort_order) VALUES (?,?,?)",
                    (eid, fn, i))
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
                           photos=photos)


@app.route("/equipment/<int:eid>")
@login_required
def equipment_detail(eid):
    con = get_db()
    eq = con.execute("""SELECT e.*, w.name AS warehouse_name, w.address AS warehouse_address
                        FROM equipment e LEFT JOIN warehouses w ON w.id=e.warehouse_id
                        WHERE e.id=?""", (eid,)).fetchone()
    if not eq:
        abort(404)
    photos = equipment_photo_list(con, eid)
    if not photos and eq["photo"]:
        photos = [eq["photo"]]
    res = con.execute(
        """SELECT r.*, u.username, u.first_name, u.last_name FROM reservations r
           JOIN users u ON u.id=r.user_id
           WHERE r.equipment_id=? AND r.status != 'anulowana'
           ORDER BY r.date_from DESC""", (eid,)).fetchall()
    today = date.today().isoformat()
    avail_today = eq["quantity"] - reserved_qty(con, eid, today, today)
    con.close()
    return render_template("equipment_detail.html", eq=eq, reservations=res,
                           avail_today=avail_today, today=today, dn=display_name,
                           photos=photos)


@app.route("/equipment/<int:eid>/delete", methods=["POST"])
@login_required
@admin_required
def equipment_delete(eid):
    con = get_db()
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
    return redirect(url_for("index"))


# ---------- rezerwacje ----------

@app.route("/reservations")
@login_required
def reservations():
    f = request.args.get("status", "")
    mine = request.args.get("mine", "")
    overdue = request.args.get("overdue", "")
    today = date.today().isoformat()
    con = get_db()
    sql = """SELECT r.*, u.username, u.first_name, u.last_name, e.code, e.name, e.photo,
                    e.location, w.name AS warehouse_name
             FROM reservations r
             JOIN users u ON u.id=r.user_id JOIN equipment e ON e.id=r.equipment_id
             LEFT JOIN warehouses w ON w.id=e.warehouse_id"""
    where, params = [], []
    if f:
        where.append("r.status=?"); params.append(f)
    if mine == "1":
        where.append("r.user_id=?"); params.append(session["user_id"])
    if overdue == "1":
        where.append("r.status='wydane' AND r.date_to<?"); params.append(today)
    if where:
        sql += " WHERE " + " AND ".join(where)
    rows = con.execute(sql + " ORDER BY r.date_from DESC", params).fetchall()
    warehouses = active_warehouses(con)
    con.close()
    return render_template("reservations.html", rows=rows, f=f, mine=mine,
                           overdue=overdue, today=today, dn=display_name,
                           warehouses=warehouses)


@app.route("/equipment/<int:eid>/reserve", methods=["GET", "POST"])
@login_required
def reserve(eid):
    con = get_db()
    eq = con.execute("SELECT * FROM equipment WHERE id=?", (eid,)).fetchone()
    if not eq:
        abort(404)
    if request.method == "POST":
        d_from, d_to = request.form["date_from"], request.form["date_to"]
        qty = max(1, int(request.form.get("quantity") or 1))
        if not valid_dates(d_from, d_to):
            flash("Nieprawidłowy zakres dat.", "error")
        else:
            taken = reserved_qty(con, eid, d_from, d_to)
            free = eq["quantity"] - taken
            if qty > free:
                flash(f"Brak dostępności w tym terminie. Wolne sztuki: {free} z {eq['quantity']}.", "error")
            else:
                rec = recipient_form_fields(request.form)
                con.execute(
                    """INSERT INTO reservations (equipment_id, user_id, client,
                       date_from, date_to, quantity, notes, receiver,
                       recipient_name, recipient_contact, recipient_phone,
                       recipient_address, recipient_email)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (eid, session["user_id"], request.form["client"].strip(),
                     d_from, d_to, qty, request.form["notes"].strip(),
                     request.form.get("receiver", ""),
                     rec["recipient_name"], rec["recipient_contact"],
                     rec["recipient_phone"], rec["recipient_address"],
                     rec["recipient_email"]))
                upsert_recipient(con, rec["recipient_name"], rec["recipient_contact"],
                                 rec["recipient_phone"], rec["recipient_address"],
                                 rec["recipient_email"])
                con.commit()
                con.close()
                flash("Rezerwacja utworzona.", "ok")
                return redirect(url_for("equipment_detail", eid=eid))
    receivers = active_partners(con)
    recipients = recent_recipients(con)
    con.close()
    return render_template("reserve.html", eq=eq, receivers=receivers,
                           recipients=recipients)


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

    # ostrzeżenie: pozycje z różnych magazynów (odbiór z więcej niż jednego miejsca)
    wh_names = {it["warehouse_name"] for it in items if it["warehouse_name"]}
    multi_warehouse = len(wh_names) > 1

    if request.method == "POST" and "date_from" in request.form:
        d_from, d_to = request.form["date_from"], request.form["date_to"]
        if not valid_dates(d_from, d_to):
            flash("Nieprawidłowy zakres dat.", "error")
        else:
            errors = []
            wanted = {}
            for it in items:
                qty = max(1, int(request.form.get(f"qty_{it['id']}") or 1))
                free = it["quantity"] - reserved_qty(con, it["id"], d_from, d_to)
                if qty > free:
                    errors.append(f"{it['code']}: wolne {free} z {it['quantity']} szt.")
                wanted[it["id"]] = qty
            if errors:
                flash("Brak dostępności – " + "; ".join(errors), "error")
            else:
                rec = recipient_form_fields(request.form)
                gid = uuid.uuid4().hex[:8]
                for it in items:
                    con.execute(
                        """INSERT INTO reservations (equipment_id, user_id, client,
                           date_from, date_to, quantity, notes, group_id, receiver,
                           recipient_name, recipient_contact, recipient_phone,
                           recipient_address, recipient_email)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (it["id"], session["user_id"], request.form["client"].strip(),
                         d_from, d_to, wanted[it["id"]],
                         request.form["notes"].strip(), gid,
                         request.form.get("receiver", ""),
                         rec["recipient_name"], rec["recipient_contact"],
                         rec["recipient_phone"], rec["recipient_address"],
                         rec["recipient_email"]))
                upsert_recipient(con, rec["recipient_name"], rec["recipient_contact"],
                                 rec["recipient_phone"], rec["recipient_address"],
                                 rec["recipient_email"])
                con.commit()
                con.close()
                flash(f"Utworzono wspólną rezerwację ({len(items)} pozycji).", "ok")
                return redirect(url_for("reservations"))
    receivers = active_partners(con)
    recipients = recent_recipients(con)
    con.close()
    return render_template("reserve_multi.html", items=items, receivers=receivers,
                           recipients=recipients, multi_warehouse=multi_warehouse,
                           wh_names=sorted(wh_names))


def _selected_reservations(con, rids):
    rids = [int(r) for r in rids if str(r).isdigit()]
    if not rids:
        return []
    return con.execute(
        f"""SELECT r.*, u.username, u.first_name, u.last_name,
            e.code, e.name, e.location, e.owner, e.brand,
            e.photo, e.dimensions, e.project_number, e.storage_instructions,
            w.name AS warehouse_name, w.address AS warehouse_address
            FROM reservations r JOIN users u ON u.id=r.user_id
            JOIN equipment e ON e.id=r.equipment_id
            LEFT JOIN warehouses w ON w.id=e.warehouse_id
            WHERE r.id IN ({','.join('?'*len(rids))}) ORDER BY e.code""",
        rids).fetchall()


def _return_form_valid(form):
    """Magazyn i miejsce są wymagane przy zwrocie."""
    wid = (form.get("return_warehouse_id") or "").strip()
    loc = (form.get("return_location") or "").strip()
    return wid.isdigit() and bool(loc)


def _apply_return(con, r, user_id, form):
    """Przyjmuje zwrot jednej rezerwacji. Magazyn i miejsce są wymagane."""
    if r["status"] != "wydane":
        return False
    if not _return_form_valid(form):
        return False
    wid = int((form.get("return_warehouse_id") or "").strip())
    loc = (form.get("return_location") or "").strip()
    damage = 1 if form.get("damage") else 0
    damage_notes = (form.get("damage_notes") or "").strip()
    now = datetime.now().isoformat(timespec="seconds")
    con.execute("""UPDATE reservations SET status='zwrócone', returned_at=?,
                   returned_by=?, damage=?, damage_notes=? WHERE id=?""",
                (now, user_id, damage, damage_notes, r["id"]))
    eid = r["equipment_id"]
    con.execute("UPDATE equipment SET warehouse_id=?, location=? WHERE id=?",
                (wid, loc, eid))
    if damage:
        stamp = f"[{date.today().isoformat()}] zwrot rez. #{r['id']}: {damage_notes or 'uszkodzenie'}"
        con.execute("""UPDATE equipment SET condition='uszkodzony',
                       condition_notes=IFNULL(condition_notes,'') ||
                       CASE WHEN IFNULL(condition_notes,'')='' THEN '' ELSE char(10) END || ?
                       WHERE id=?""", (stamp, eid))
    return True


@app.route("/reservations/bulk/<action>", methods=["POST"])
@login_required
@admin_required
def bulk_action(action):
    if action not in ("issue", "return"):
        abort(404)
    con = get_db()
    rows = _selected_reservations(con, request.form.getlist("rid"))
    now = datetime.now().isoformat(timespec="seconds")
    n = 0
    if action == "return" and not _return_form_valid(request.form):
        con.close()
        flash("Wypełnij pole.", "error")
        return redirect(url_for("reservations"))
    for r in rows:
        if action == "issue" and r["status"] == "rezerwacja":
            con.execute("UPDATE reservations SET status='wydane', issued_at=?, issued_by=? WHERE id=?",
                        (now, session["user_id"], r["id"]))
            n += 1
        elif action == "return" and _apply_return(con, r, session["user_id"], request.form):
            n += 1
    con.commit()
    con.close()
    verb = "Wydano" if action == "issue" else "Przyjęto zwrot"
    flash(f"{verb}: {n} pozycji." if n else "Brak pozycji o odpowiednim statusie.", "ok" if n else "error")
    return redirect(url_for("reservations"))


@app.route("/reservations/pdf-group/<kind>")
@login_required
def pdf_group(kind):
    if kind not in ("wydanie", "przyjecie"):
        abort(404)
    con = get_db()
    rows = _selected_reservations(con, request.args.getlist("rid"))
    con.close()
    if not rows:
        flash("Zaznacz przynajmniej jedną rezerwację.", "error")
        return redirect(url_for("reservations"))
    buf = group_pdf(kind, rows)
    prefix = "WZ" if kind == "wydanie" else "PZ"
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                     download_name=f"{prefix}_zbiorczy_{datetime.now():%Y%m%d_%H%M}.pdf")


def _get_reservation(con, rid):
    r = con.execute(
        """SELECT r.*, u.username, u.first_name, u.last_name, e.code, e.name
           FROM reservations r
           JOIN users u ON u.id=r.user_id JOIN equipment e ON e.id=r.equipment_id
           WHERE r.id=?""", (rid,)).fetchone()
    if not r:
        abort(404)
    return r


@app.route("/reservations/<int:rid>/issue", methods=["POST"])
@login_required
@admin_required
def issue(rid):
    con = get_db()
    r = _get_reservation(con, rid)
    if r["status"] != "rezerwacja":
        flash("Można wydać tylko aktywną rezerwację.", "error")
    else:
        con.execute("UPDATE reservations SET status='wydane', issued_at=?, issued_by=? WHERE id=?",
                    (datetime.now().isoformat(timespec="seconds"), session["user_id"], rid))
        con.commit()
        flash("Oznaczono jako wydane.", "ok")
    con.close()
    return redirect(request.referrer or url_for("reservations"))


@app.route("/reservations/<int:rid>/return", methods=["POST"])
@login_required
@admin_required
def return_item(rid):
    con = get_db()
    r = _get_reservation(con, rid)
    if r["status"] != "wydane":
        flash("Można zwrócić tylko wydany sprzęt.", "error")
    elif not _return_form_valid(request.form):
        flash("Wypełnij pole.", "error")
    elif _apply_return(con, r, session["user_id"], request.form):
        con.commit()
        damaged = bool(request.form.get("damage"))
        flash("Oznaczono jako zwrócone." + (" Odnotowano uszkodzenie." if damaged else ""), "ok")
    else:
        flash("Nie udało się przyjąć zwrotu.", "error")
    con.close()
    return redirect(request.referrer or url_for("reservations"))


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


@app.route("/reservations/<int:rid>/pdf/<kind>")
@login_required
def reservation_pdf(rid, kind):
    if kind not in ("wydanie", "przyjecie"):
        abort(404)
    con = get_db()
    r = _get_reservation(con, rid)
    eq = con.execute("""SELECT e.*, w.name AS warehouse_name, w.address AS warehouse_address
                        FROM equipment e LEFT JOIN warehouses w ON w.id=e.warehouse_id
                        WHERE e.id=?""", (r["equipment_id"],)).fetchone()
    photos = equipment_photo_list(con, r["equipment_id"])
    if not photos and eq["photo"]:
        photos = [eq["photo"]]
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


# ---------- API: autouzupełnianie adresatów ----------

@app.route("/api/recipients")
@login_required
def api_recipients():
    con = get_db()
    rows = recent_recipients(con, limit=100)
    con.close()
    return jsonify([dict(r) for r in rows])


# ---------- słowniki (admin): magazyny + podwykonawcy ----------

@app.route("/dictionaries")
@login_required
@admin_required
def dictionaries():
    con = get_db()
    warehouses = con.execute("SELECT * FROM warehouses ORDER BY name").fetchall()
    partners = con.execute("SELECT * FROM logistics_partners ORDER BY name").fetchall()
    con.close()
    return render_template("dictionaries.html", warehouses=warehouses, partners=partners)


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


# ---------- użytkownicy (admin) ----------

@app.route("/users", methods=["GET", "POST"])
@login_required
@admin_required
def users():
    con = get_db()
    if request.method == "POST":
        fn = request.form.get("first_name", "").strip()
        ln = request.form.get("last_name", "").strip()
        if not fn or not ln:
            flash("Imię i nazwisko są wymagane – każde konto to konkretny PM.", "error")
        else:
            try:
                con.execute("""INSERT INTO users (username, password_hash, role,
                               first_name, last_name) VALUES (?,?,?,?,?)""",
                            (request.form["username"].strip(),
                             generate_password_hash(request.form["password"], method="pbkdf2:sha256"),
                             request.form.get("role", "user"), fn, ln))
                con.commit()
                flash("Użytkownik dodany.", "ok")
            except Exception:
                flash("Taki login już istnieje.", "error")
    rows = con.execute("SELECT * FROM users ORDER BY last_name, username").fetchall()
    con.close()
    return render_template("users.html", rows=rows, dn=display_name)


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
    if len(pw) < 6:
        flash("Hasło musi mieć min. 6 znaków.", "error")
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
    if not fn or not ln:
        flash("Imię i nazwisko są wymagane.", "error")
    else:
        con = get_db()
        con.execute("UPDATE users SET first_name=?, last_name=? WHERE id=?", (fn, ln, uid))
        con.commit()
        con.close()
        flash("Dane zaktualizowane.", "ok")
    return redirect(url_for("users"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
