"""
app.py — MUJI Label Report System Backend
Flask + SQLite + pdfplumber + python-docx

Deploy: Render.com (free tier)
  - Build: pip install -r requirements.txt
  - Start: python app.py
"""

import os, re, io, json, base64, sqlite3, shutil, zipfile
from datetime import datetime
from pathlib import Path
from flask import Flask, request, jsonify, send_file, send_from_directory, g
from werkzeug.utils import secure_filename
import pdfplumber
import openpyxl
from PIL import Image as PILImage

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
DB_PATH    = BASE_DIR / "data" / "products.db"
IMG_DIR    = BASE_DIR / "data" / "images"
REPORT_DIR = BASE_DIR / "data" / "reports"
STATIC_DIR = BASE_DIR / "static"

for d in [DB_PATH.parent, IMG_DIR, REPORT_DIR, STATIC_DIR]:
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, static_folder=str(STATIC_DIR))
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 100MB

# ── Database ──────────────────────────────────────────────────────────────────
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(str(DB_PATH))
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def init_db():
    with sqlite3.connect(str(DB_PATH)) as db:
        db.executescript("""
        CREATE TABLE IF NOT EXISTS products (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name_en     TEXT NOT NULL,
            name_th     TEXT,
            ui_number   TEXT NOT NULL,
            img1_path   TEXT,
            img2_path   TEXT,
            img3_path   TEXT,
            brand       TEXT DEFAULT 'MUJI',
            category    TEXT,
            status      TEXT DEFAULT 'active',
            created_at  TEXT DEFAULT (datetime('now','localtime')),
            updated_at  TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_ui ON products(ui_number) WHERE status='active';
        CREATE INDEX IF NOT EXISTS idx_name ON products(name_en);

        CREATE TABLE IF NOT EXISTS reports (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            filename     TEXT NOT NULL,
            customs_no   TEXT,
            invoice_no   TEXT,
            item_count   INTEGER,
            created_at   TEXT DEFAULT (datetime('now','localtime'))
        );
        """)

init_db()

# ── Helpers ───────────────────────────────────────────────────────────────────
def resize_save(img_bytes: bytes, dest: Path, max_wh=800) -> Path:
    img = PILImage.open(io.BytesIO(img_bytes)).convert("RGBA")
    img.thumbnail((max_wh, max_wh), PILImage.LANCZOS)
    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(dest), "PNG")
    return dest

def row_to_dict(row):
    return dict(row) if row else None

def match_product(db, name: str):
    name_up = name.upper().strip()
    # 1. Exact match
    r = db.execute("SELECT * FROM products WHERE UPPER(name_en)=? AND status='active'", (name_up,)).fetchone()
    if r: return row_to_dict(r)
    # 2. Partial match (name contains DB name or vice versa)
    rows = db.execute("SELECT * FROM products WHERE status='active'").fetchall()
    for row in rows:
        k = row["name_en"].upper()
        if name_up in k or k in name_up:
            return row_to_dict(row)
    # 3. Word-level fuzzy (all words of shorter must be in longer)
    name_words = set(name_up.split())
    best, best_score = None, 0
    for row in rows:
        k_words = set(row["name_en"].upper().split())
        common = len(name_words & k_words)
        score = common / max(len(name_words), len(k_words))
        if score > best_score and score >= 0.6:
            best_score, best = score, row_to_dict(row)
    return best

# ── PDF Parser ────────────────────────────────────────────────────────────────
KEYWORDS = ["MUJI","BRAND","SNACK","BOLO","SQUID","CRACKERS","PIE","POPCORN",
            "BISCUIT","TEA","LATTE","CHAI","RAMEN","STICKS","CHIPS","DORAYAKI",
            "WAFER","COOKIE","CANDY","JELLY","GUMMY","RICE","PASTA","NOODLE"]

def parse_customs_pdf(pdf_bytes: bytes) -> dict:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    lines = text.split("\n")

    def find(pat):
        m = re.search(pat, text)
        return m.group(1).strip() if m and m.lastindex else (m.group(0).strip() if m else "")

    customs_no  = find(r"(A\d{3}-\d-\d{4}-\d+)")
    invoice_no  = find(r"(THA[\w\-]+\d{2}[A-Z0-9\-]+)")
    ship_name   = find(r"(YM[\s\w]+?)(?:\n|วัน)")
    import_date = find(r"วันที่นำเข้า\s*\n?([\d/]+)") or find(r"(\d{2}/\d{2}/25\d{2})")
    total_ctn   = find(r"(\d[\d,]+)\s+CARTONS\s*\(") or find(r"(\d[\d,]+)\s+CARTONS")

    # qty lines: contain KGM and NNN C62
    qty_map = {}
    for i, line in enumerate(lines):
        m = re.search(r"(\d[\d,]*)\s+C62\s", line)
        if m and "KGM" in line:
            qty_map[i] = int(m.group(1).replace(",", ""))

    # product name lines
    products, seen = [], set()
    for qi, qty in sorted(qty_map.items(), key=lambda x: x[0]):
        for back in range(1, 10):
            idx = qi - back
            if idx < 0: break
            line = lines[idx]
            if re.search(r"[A-Z]{3}", line) and any(k in line for k in KEYWORDS):
                name = re.sub(r'"MUJI"\s*JP.*', "", line)
                name = re.sub(r"^OR\s+", "", name)
                name = re.sub(r"รหัส.*", "", name).strip()
                key  = name[:30].upper()
                if len(name) > 4 and key not in seen:
                    seen.add(key)
                    products.append({"name": name, "qty": qty})
                break

    # catch patterns that slip through
    extras = [
        r"STRAWBERRY JAM PIE[^\n\"]+",
        r"INSTANT ROASTED GREEN TEA LATTE[^\n\"]+",
        r"DRIED SQUID[^\n\"]+",
    ]
    for pat in extras:
        m = re.search(pat, text)
        if m:
            nm = re.sub(r'"MUJI"\s*JP.*', "", m.group(0)).strip()
            key = nm[:30].upper()
            if key not in seen:
                seen.add(key)
                qm = re.search(rf"{re.escape(nm[:15])}.*?(\d[\d,]*)\s+C62", text[:6000], re.S)
                qty = int(qm.group(1).replace(",","")) if qm else 0
                products.insert(0, {"name": nm, "qty": qty})

    products.sort(key=lambda p: text.find(p["name"]) if p["name"] in text else 9999)

    return {
        "customs_no":  customs_no,
        "invoice_no":  invoice_no,
        "ship_name":   ship_name.strip(),
        "import_date": import_date,
        "total_cartons": total_ctn,
        "products":    products,
    }

# ── Word Generator ────────────────────────────────────────────────────────────
def generate_docx(pdf_data: dict, matched_items: list, db) -> bytes:
    """
    Build .docx using docx-js (Node.js) called as subprocess.
    We write a JSON payload, call the node script, read the result.
    """
    # Embed images as base64 in the payload
    items_payload = []
    for item in matched_items:
        imgs = []
        for k in ["img1_path", "img2_path", "img3_path"]:
            p = item.get(k)
            if p:
                fp = IMG_DIR / p
                if fp.exists():
                    with open(fp, "rb") as f:
                        imgs.append(base64.b64encode(f.read()).decode())
        items_payload.append({
            "no":     item["no"],
            "name":   item["name"],
            "qty":    item["qty"],
            "reg":    item.get("ui_number", ""),
            "imgs":   imgs,
        })

    payload = {
        "meta": {
            "customs_no":    pdf_data["customs_no"],
            "invoice_no":    pdf_data["invoice_no"],
            "ship_name":     pdf_data["ship_name"],
            "import_date":   pdf_data["import_date"],
            "total_cartons": pdf_data["total_cartons"],
            "total_items":   len(matched_items),
        },
        "items": items_payload,
    }

    payload_path = BASE_DIR / "data" / "_docx_payload.json"
    out_path     = BASE_DIR / "data" / "_docx_out.docx"
    script_path  = BASE_DIR / "generate_docx.js"

    with open(payload_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)

    ret = os.system(f'node "{script_path}" "{payload_path}" "{out_path}"')
    if ret != 0:
        raise RuntimeError("docx generation failed")

    with open(out_path, "rb") as f:
        return f.read()

# ── API Routes ────────────────────────────────────────────────────────────────

# --- Serve frontend ---
@app.route("/")
def index():
    return send_from_directory(str(STATIC_DIR), "index.html")

@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(str(STATIC_DIR), filename)

# --- Product CRUD ---
@app.route("/api/products", methods=["GET"])
def list_products():
    db = get_db()
    q = request.args.get("q", "")
    if q:
        rows = db.execute(
            "SELECT * FROM products WHERE status='active' AND (UPPER(name_en) LIKE ? OR ui_number LIKE ?) ORDER BY name_en",
            (f"%{q.upper()}%", f"%{q}%")
        ).fetchall()
    else:
        rows = db.execute("SELECT * FROM products WHERE status='active' ORDER BY name_en").fetchall()
    return jsonify([row_to_dict(r) for r in rows])

@app.route("/api/products/<int:pid>", methods=["GET"])
def get_product(pid):
    db = get_db()
    row = db.execute("SELECT * FROM products WHERE id=? AND status='active'", (pid,)).fetchone()
    if not row: return jsonify({"error": "Not found"}), 404
    return jsonify(row_to_dict(row))

@app.route("/api/products", methods=["POST"])
def create_product():
    db = get_db()
    data = request.form
    name_en   = data.get("name_en", "").strip()
    ui_number = data.get("ui_number", "").strip()
    if not name_en or not ui_number:
        return jsonify({"error": "name_en and ui_number required"}), 400

    img_paths = []
    for i in range(1, 4):
        file = request.files.get(f"img{i}")
        if file and file.filename:
            fn = f"{ui_number}_img{i}.png"
            resize_save(file.read(), IMG_DIR / fn)
            img_paths.append(fn)
        else:
            img_paths.append(None)

    cur = db.execute(
        """INSERT INTO products (name_en,name_th,ui_number,img1_path,img2_path,img3_path,brand,category)
           VALUES (?,?,?,?,?,?,?,?)""",
        (name_en, data.get("name_th",""), ui_number,
         img_paths[0], img_paths[1], img_paths[2],
         data.get("brand","MUJI"), data.get("category",""))
    )
    db.commit()
    return jsonify({"id": cur.lastrowid, "message": "Created"}), 201

@app.route("/api/products/<int:pid>", methods=["PUT"])
def update_product(pid):
    db = get_db()
    data = request.form
    fields, vals = [], []
    for col in ["name_en","name_th","ui_number","brand","category"]:
        if col in data:
            fields.append(f"{col}=?")
            vals.append(data[col])

    for i in range(1, 4):
        file = request.files.get(f"img{i}")
        if file and file.filename:
            row = db.execute("SELECT ui_number FROM products WHERE id=?", (pid,)).fetchone()
            fn  = f"{row['ui_number']}_img{i}.png"
            resize_save(file.read(), IMG_DIR / fn)
            fields.append(f"img{i}_path=?")
            vals.append(fn)

    if not fields:
        return jsonify({"error": "Nothing to update"}), 400

    fields.append("updated_at=datetime('now','localtime')")
    vals.append(pid)
    db.execute(f"UPDATE products SET {', '.join(fields)} WHERE id=?", vals)
    db.commit()
    return jsonify({"message": "Updated"})

@app.route("/api/products/<int:pid>", methods=["DELETE"])
def delete_product(pid):
    db = get_db()
    db.execute("UPDATE products SET status='deleted', updated_at=datetime('now','localtime') WHERE id=?", (pid,))
    db.commit()
    return jsonify({"message": "Deleted"})

# --- Product image ---
@app.route("/api/products/<int:pid>/image/<int:imgno>")
def product_image(pid, imgno):
    db = get_db()
    row = db.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    if not row: return "", 404
    fn = row[f"img{imgno}_path"]
    if not fn: return "", 404
    fp = IMG_DIR / fn
    if not fp.exists(): return "", 404
    return send_file(str(fp), mimetype="image/png")

# --- Stats ---
@app.route("/api/stats")
def stats():
    db = get_db()
    total   = db.execute("SELECT COUNT(*) FROM products WHERE status='active'").fetchone()[0]
    with_img= db.execute("SELECT COUNT(*) FROM products WHERE status='active' AND img1_path IS NOT NULL").fetchone()[0]
    reports = db.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
    recent  = db.execute("SELECT * FROM reports ORDER BY id DESC LIMIT 5").fetchall()
    return jsonify({
        "total_products":  total,
        "with_images":     with_img,
        "without_images":  total - with_img,
        "total_reports":   reports,
        "recent_reports":  [row_to_dict(r) for r in recent],
    })

# --- Parse PDF ---
@app.route("/api/parse-pdf", methods=["POST"])
def parse_pdf_route():
    file = request.files.get("pdf")
    if not file: return jsonify({"error": "No PDF"}), 400
    try:
        data = parse_customs_pdf(file.read())
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- Generate report ---
@app.route("/api/generate-report", methods=["POST"])
def generate_report():
    file = request.files.get("pdf")
    if not file: return jsonify({"error": "No PDF"}), 400

    db = get_db()
    try:
        pdf_data = parse_customs_pdf(file.read())
    except Exception as e:
        return jsonify({"error": f"PDF parse failed: {e}"}), 500

    matched_items = []
    for i, product in enumerate(pdf_data["products"], 1):
        m = match_product(db, product["name"])
        item = {
            "no":       i,
            "name":     product["name"],
            "qty":      product["qty"],
            "matched":  bool(m),
            "ui_number": m["ui_number"] if m else "",
            "img1_path": m["img1_path"] if m else None,
            "img2_path": m["img2_path"] if m else None,
            "img3_path": m["img3_path"] if m else None,
        }
        matched_items.append(item)

    try:
        docx_bytes = generate_docx(pdf_data, matched_items, db)
    except Exception as e:
        return jsonify({"error": f"Word generation failed: {e}"}), 500

    # Save report record
    cn = re.sub(r"[^\w\-]", "-", pdf_data["customs_no"])
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    fn = f"รายงาน_{cn}_{stamp}.docx"
    out_path = REPORT_DIR / fn
    with open(out_path, "wb") as f:
        f.write(docx_bytes)

    db.execute("INSERT INTO reports (filename,customs_no,invoice_no,item_count) VALUES (?,?,?,?)",
               (fn, pdf_data["customs_no"], pdf_data["invoice_no"], len(matched_items)))
    db.commit()

    return send_file(
        io.BytesIO(docx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=fn,
    )

# --- Bulk import from Picture List xlsx ---
@app.route("/api/import-picture-list", methods=["POST"])
def import_picture_list():
    file = request.files.get("xlsx")
    if not file: return jsonify({"error": "No file"}), 400

    db   = get_db()
    data = file.read()
    wb   = openpyxl.load_workbook(io.BytesIO(data))
    ws   = wb.active

    # Map row -> images (up to 3 per row, deduplicated)
    from collections import defaultdict
    row_imgs = defaultdict(list)
    for img in ws._images:
        try:
            r = img.anchor._from.row + 1
            seen_sizes = {len(d) for d in row_imgs[r]}
            raw = img._data()
            if len(raw) not in seen_sizes:
                row_imgs[r].append(raw)
        except Exception:
            pass

    added, updated, skipped = 0, 0, 0
    for row in ws.iter_rows(min_row=2):
        ui_number = str(row[0].value or "").strip()
        name_en   = str(row[1].value or "").strip()
        if not ui_number or not name_en:
            skipped += 1
            continue

        imgs = row_imgs.get(row[0].row, [])[:3]
        img_paths = []
        for idx, raw in enumerate(imgs, 1):
            fn = f"{ui_number}_img{idx}.png"
            try:
                resize_save(raw, IMG_DIR / fn)
                img_paths.append(fn)
            except Exception:
                img_paths.append(None)
        while len(img_paths) < 3:
            img_paths.append(None)

        existing = db.execute("SELECT id FROM products WHERE ui_number=?", (ui_number,)).fetchone()
        if existing:
            db.execute(
                """UPDATE products SET name_en=?,img1_path=?,img2_path=?,img3_path=?,
                   updated_at=datetime('now','localtime') WHERE ui_number=?""",
                (name_en, img_paths[0], img_paths[1], img_paths[2], ui_number)
            )
            updated += 1
        else:
            db.execute(
                """INSERT INTO products (name_en,ui_number,img1_path,img2_path,img3_path)
                   VALUES (?,?,?,?,?)""",
                (name_en, ui_number, img_paths[0], img_paths[1], img_paths[2])
            )
            added += 1

    db.commit()
    return jsonify({"added": added, "updated": updated, "skipped": skipped,
                    "total": added + updated})

# --- Export DB ---
@app.route("/api/export")
def export_db():
    fmt = request.args.get("format", "json")
    db  = get_db()
    rows = [row_to_dict(r) for r in
            db.execute("SELECT * FROM products WHERE status='active' ORDER BY name_en").fetchall()]

    if fmt == "csv":
        import csv, io as sio
        buf = sio.StringIO()
        if rows:
            w = csv.DictWriter(buf, fieldnames=rows[0].keys())
            w.writeheader()
            w.writerows(rows)
        return send_file(
            io.BytesIO(buf.getvalue().encode("utf-8-sig")),
            mimetype="text/csv",
            as_attachment=True,
            download_name="products_export.csv",
        )
    return jsonify(rows)

# --- Recent reports ---
@app.route("/api/reports")
def list_reports():
    db   = get_db()
    rows = db.execute("SELECT * FROM reports ORDER BY id DESC LIMIT 20").fetchall()
    return jsonify([row_to_dict(r) for r in rows])

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
