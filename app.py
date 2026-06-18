"""
app.py — MUJI Label Report System (complete single-file)
"""
import os, re, io, json, base64, sqlite3, subprocess, tempfile
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from flask import Flask, request, jsonify, send_file, g
import pdfplumber, openpyxl
from PIL import Image as PILImage

BASE   = Path(__file__).parent
DATA   = BASE / "data"
IMGDIR = DATA / "images"
RPTDIR = DATA / "reports"
DB     = DATA / "products.db"
JSGEN  = BASE / "generate_docx.js"
for d in [DATA, IMGDIR, RPTDIR]: d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(str(DB)); g.db.row_factory = sqlite3.Row
        g.db.executescript("""
        CREATE TABLE IF NOT EXISTS products(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name_en TEXT NOT NULL, ui_number TEXT NOT NULL,
          img1 TEXT, img2 TEXT, img3 TEXT,
          status TEXT DEFAULT 'active',
          created_at TEXT DEFAULT (datetime('now','localtime')),
          updated_at TEXT DEFAULT (datetime('now','localtime')));
        CREATE INDEX IF NOT EXISTS idx_name ON products(name_en);
        CREATE TABLE IF NOT EXISTS reports(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          customs_no TEXT, invoice_no TEXT, item_count INTEGER, filename TEXT,
          created_at TEXT DEFAULT (datetime('now','localtime')));
        """)
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def save_img(raw, fn):
    img = PILImage.open(io.BytesIO(raw)).convert("RGBA"); img.thumbnail((600,600))
    out = io.BytesIO(); img.save(out,"PNG"); (IMGDIR/fn).write_bytes(out.getvalue()); return fn

KEYWORDS=["MUJI","BRAND","SNACK","BOLO","SQUID","CRACKERS","PIE","POPCORN",
          "BISCUIT","TEA","LATTE","CHAI","RAMEN","STICKS","CHIPS","DORAYAKI",
          "WAFER","COOKIE","RICE","PASTA","SEAWEED","CARAMEL","EGG"]

def match(db, name):
    key=name.upper().strip()
    r=db.execute("SELECT * FROM products WHERE UPPER(name_en)=? AND status='active'",(key,)).fetchone()
    if r: return dict(r)
    for row in db.execute("SELECT * FROM products WHERE status='active'").fetchall():
        k=row["name_en"].upper()
        if key in k or k in key: return dict(row)
    kw=set(key.split()); best,bsc=None,0
    for row in db.execute("SELECT * FROM products WHERE status='active'").fetchall():
        rw=set(row["name_en"].upper().split()); sc=len(kw&rw)/max(len(kw),len(rw))
        if sc>bsc and sc>=0.6: bsc,best=sc,dict(row)
    return best

def parse_pdf(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text="\n".join(p.extract_text() or "" for p in pdf.pages)
    lines=text.split("\n")
    def find(pat):
        m=re.search(pat,text); return m.group(1).strip() if m and m.lastindex else ""
    customs_no=find(r"(A\d{3}-\d-\d{4}-\d+)")
    invoice_no=find(r"(THA[\w\-]+\d{2}[A-Z0-9\-]+)")
    ship_name=find(r"(YM[\s\w]+?)(?:\n|วัน)")
    import_date=find(r"(\d{2}/\d{2}/25\d{2})")
    total_ctn=find(r"(\d[\d,]+)\s+CARTONS\s*\(") or find(r"(\d[\d,]+)\s+CARTONS")
    qty_map={}
    for i,ln in enumerate(lines):
        m=re.search(r"(\d[\d,]*)\s+C62\s",ln)
        if m and "KGM" in ln: qty_map[i]=int(m.group(1).replace(",",""))
    products,seen=[],set()
    for qi,qty in sorted(qty_map.items()):
        for back in range(1,10):
            idx=qi-back
            if idx<0: break
            ln=lines[idx]
            if re.search(r"[A-Z]{3}",ln) and any(k in ln for k in KEYWORDS):
                name=re.sub(r'"MUJI"\s*JP.*',"",ln); name=re.sub(r"^OR\s+","",name)
                name=re.sub(r"รหัส.*","",name).strip(); key=name[:30].upper()
                if len(name)>4 and key not in seen: seen.add(key); products.append({"name":name,"qty":qty}); break
    for pat in [r"STRAWBERRY JAM PIE[^\n\"]+",r"INSTANT ROASTED GREEN TEA LATTE[^\n\"]+",r"DRIED SQUID[^\n\"]+"]:
        m=re.search(pat,text)
        if m:
            nm=re.sub(r'"MUJI"\s*JP.*',"",m.group(0)).strip(); key=nm[:30].upper()
            if key not in seen:
                seen.add(key); qm=re.search(rf"{re.escape(nm[:15])}.*?(\d[\d,]*)\s+C62",text[:6000],re.S)
                qty=int(qm.group(1).replace(",","")) if qm else 0; products.insert(0,{"name":nm,"qty":qty})
    products.sort(key=lambda p: text.find(p["name"]) if p["name"] in text else 9999)
    return dict(customs_no=customs_no,invoice_no=invoice_no,ship_name=ship_name.strip(),
                import_date=import_date,total_cartons=total_ctn,products=products)

def make_docx(pdf_data, items):
    payload={"meta":{
        "customs_no":pdf_data["customs_no"],"invoice_no":pdf_data["invoice_no"],
        "ship_name":pdf_data.get("ship_name",""),"import_date":pdf_data.get("import_date",""),
        "total_cartons":pdf_data.get("total_cartons",""),"total_items":len(items)},"items":[]}
    for it in items:
        imgs=[]
        for k in ["img1","img2","img3"]:
            fn=it.get(k)
            if fn:
                fp=IMGDIR/fn
                if fp.exists(): imgs.append("data:image/png;base64,"+base64.b64encode(fp.read_bytes()).decode())
        payload["items"].append({"no":it["no"],"name":it["name"],"qty":it["qty"],"reg":it.get("ui_number",""),"imgs":imgs})
    with tempfile.NamedTemporaryFile(suffix=".json",delete=False,mode="w",encoding="utf-8") as f:
        json.dump(payload,f,ensure_ascii=False); jpath=f.name
    opath=jpath.replace(".json",".docx")
    result=subprocess.run(["node",str(JSGEN),jpath,opath],capture_output=True,text=True,timeout=120)
    os.unlink(jpath)
    if result.returncode!=0: raise RuntimeError(result.stderr[:500])
    data=Path(opath).read_bytes(); os.unlink(opath); return data

# ── HTML (inline) ─────────────────────────────────────────────────────────────
HTML = open(BASE/"static"/"index.html", encoding="utf-8").read()

@app.route("/")
def index(): return HTML

@app.route("/api/stats")
def stats():
    db=get_db()
    return jsonify({"total":db.execute("SELECT COUNT(*) FROM products WHERE status='active'").fetchone()[0],
                    "with_img":db.execute("SELECT COUNT(*) FROM products WHERE status='active' AND img1 IS NOT NULL").fetchone()[0],
                    "reports":db.execute("SELECT COUNT(*) FROM reports").fetchone()[0],
                    "recent":[dict(r) for r in db.execute("SELECT * FROM reports ORDER BY id DESC LIMIT 5")]})

@app.route("/api/parse-pdf",methods=["POST"])
def api_parse_pdf():
    f=request.files.get("pdf")
    if not f: return jsonify({"error":"ไม่พบไฟล์ PDF"}),400
    try: return jsonify(parse_pdf(f.read()))
    except Exception as e: return jsonify({"error":str(e)}),500

@app.route("/api/generate",methods=["POST"])
def api_generate():
    f=request.files.get("pdf")
    if not f: return jsonify({"error":"ไม่พบไฟล์ PDF"}),400
    db=get_db()
    try: pdf_data=parse_pdf(f.read())
    except Exception as e: return jsonify({"error":f"อ่าน PDF ไม่สำเร็จ: {e}"}),500
    items=[]
    for i,p in enumerate(pdf_data["products"],1):
        m=match(db,p["name"])
        items.append({"no":i,"name":p["name"],"qty":p["qty"],
                      "ui_number":m["ui_number"] if m else "",
                      "img1":m["img1"] if m else None,"img2":m["img2"] if m else None,"img3":m["img3"] if m else None})
    try: docx_bytes=make_docx(pdf_data,items)
    except Exception as e: return jsonify({"error":f"สร้าง Word ไม่สำเร็จ: {e}"}),500
    cn=re.sub(r"[^\w\-]","-",pdf_data["customs_no"])
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
    fn=f"รายงานการติดฉลากสินค้า_{cn}_{stamp}.docx"
    (RPTDIR/fn).write_bytes(docx_bytes)
    db.execute("INSERT INTO reports (customs_no,invoice_no,item_count,filename) VALUES (?,?,?,?)",
               (pdf_data["customs_no"],pdf_data["invoice_no"],len(items),fn)); db.commit()
    return send_file(io.BytesIO(docx_bytes),
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     as_attachment=True,download_name=fn)

@app.route("/api/import-xl",methods=["POST"])
def api_import_xl():
    f=request.files.get("xlsx")
    if not f: return jsonify({"error":"ไม่พบไฟล์"}),400
    db=get_db(); data=f.read(); wb=openpyxl.load_workbook(io.BytesIO(data)); ws=wb.active
    row_imgs=defaultdict(list)
    for img in ws._images:
        try:
            r=img.anchor._from.row+1; raw=img._data()
            if not any(len(x)==len(raw) for x in row_imgs[r]): row_imgs[r].append(raw)
        except: pass
    added=updated=skipped=0
    for row in ws.iter_rows(min_row=2):
        ui=str(row[0].value or "").strip(); nm=str(row[1].value or "").strip()
        if not ui or not nm: skipped+=1; continue
        imgs=row_imgs.get(row[0].row,[])[:3]; paths=[]
        for idx,raw in enumerate(imgs,1):
            fn=f"{ui}_img{idx}.png"
            try: save_img(raw,fn); paths.append(fn)
            except: paths.append(None)
        while len(paths)<3: paths.append(None)
        ex=db.execute("SELECT id FROM products WHERE ui_number=?",(ui,)).fetchone()
        if ex: db.execute("UPDATE products SET name_en=?,img1=?,img2=?,img3=?,updated_at=datetime('now','localtime') WHERE ui_number=?",(nm,paths[0],paths[1],paths[2],ui)); updated+=1
        else: db.execute("INSERT INTO products (name_en,ui_number,img1,img2,img3) VALUES (?,?,?,?,?)",(nm,ui,paths[0],paths[1],paths[2])); added+=1
    db.commit()
    return jsonify({"added":added,"updated":updated,"skipped":skipped})

@app.route("/api/products")
def api_products():
    db=get_db(); q=request.args.get("q","")
    if q: rows=db.execute("SELECT * FROM products WHERE status='active' AND (UPPER(name_en) LIKE ? OR ui_number LIKE ?) ORDER BY name_en",(f"%{q.upper()}%",f"%{q}%")).fetchall()
    else: rows=db.execute("SELECT * FROM products WHERE status='active' ORDER BY name_en").fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/products",methods=["POST"])
def api_create():
    db=get_db(); nm=request.form.get("name_en","").strip(); ui=request.form.get("ui_number","").strip()
    if not nm or not ui: return jsonify({"error":"name_en และ ui_number จำเป็น"}),400
    paths=[]
    for i in range(1,4):
        f=request.files.get(f"img{i}")
        if f and f.filename: fn=f"{ui}_img{i}.png"; save_img(f.read(),fn); paths.append(fn)
        else: paths.append(None)
    cur=db.execute("INSERT INTO products (name_en,ui_number,img1,img2,img3) VALUES (?,?,?,?,?)",(nm,ui,paths[0],paths[1],paths[2])); db.commit()
    return jsonify({"id":cur.lastrowid}),201

@app.route("/api/products/<int:pid>",methods=["PUT"])
def api_update(pid):
    db=get_db(); nm=request.form.get("name_en","").strip(); ui=request.form.get("ui_number","").strip()
    upd=[]
    if nm: upd.append(("name_en",nm))
    if ui: upd.append(("ui_number",ui))
    for i in range(1,4):
        f=request.files.get(f"img{i}")
        if f and f.filename: fn=f"{ui or str(pid)}_img{i}.png"; save_img(f.read(),fn); upd.append((f"img{i}",fn))
    if upd:
        sets=", ".join(f"{k}=?" for k,_ in upd); vals=[v for _,v in upd]+[pid]
        db.execute(f"UPDATE products SET {sets}, updated_at=datetime('now','localtime') WHERE id=?",vals); db.commit()
    return jsonify({"ok":True})

@app.route("/api/products/<int:pid>",methods=["DELETE"])
def api_delete(pid):
    db=get_db(); db.execute("UPDATE products SET status='deleted' WHERE id=?",(pid,)); db.commit(); return jsonify({"ok":True})

@app.route("/api/img/<int:pid>/<int:n>")
def api_img(pid,n):
    db=get_db(); row=db.execute("SELECT * FROM products WHERE id=?",(pid,)).fetchone()
    if not row: return "",404
    fn=row[f"img{n}"]
    if not fn: return "",404
    fp=IMGDIR/fn
    if not fp.exists(): return "",404
    return send_file(str(fp),mimetype="image/png")

@app.route("/api/export")
def api_export():
    db=get_db(); rows=db.execute("SELECT id,name_en,ui_number,status,created_at FROM products WHERE status='active' ORDER BY name_en").fetchall()
    import csv, io as sio
    buf=sio.StringIO(); w=csv.DictWriter(buf,fieldnames=["id","name_en","ui_number","status","created_at"])
    w.writeheader(); w.writerows([dict(r) for r in rows])
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")),mimetype="text/csv",as_attachment=True,download_name="products.csv")

@app.route("/api/history")
def api_history():
    db=get_db(); return jsonify([dict(r) for r in db.execute("SELECT * FROM reports ORDER BY id DESC LIMIT 30")])

if __name__=="__main__":
    port=int(os.environ.get("PORT",5000)); app.run(host="0.0.0.0",port=port,debug=False)
