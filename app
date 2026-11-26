import os
import io
import math
import sqlite3
import urllib.request
from datetime import datetime
from functools import wraps
from flask import (
    Flask, request, redirect, url_for, session,
    render_template_string, send_file, jsonify, flash
)

# ไลบรารีสำหรับสร้างไฟล์รายงาน PDF, Word, Excel
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import Color

import docx
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "mj-stock-secret-key-2026")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "mj_stock.db")
BACKUP_DIR = os.path.join(BASE_DIR, "backup_exports")
FONT_DIR = os.path.join(BASE_DIR, "fonts")
FONT_PATH = os.path.join(FONT_DIR, "Kanit-Bold.ttf")

os.makedirs(BACKUP_DIR, exist_ok=True)
os.makedirs(FONT_DIR, exist_ok=True)

WATERMARK_TEXT = "ระบบนับสต็อก พัฒนาโดย MJ"
FOOTER_CREDIT = "ซอฟต์แวร์ โดย MJ"

# ดาวน์โหลดฟอนต์ Kanit (ไม่มีหัว ไม่มีลิขสิทธิ์ SIL Open Font License) อัตโนมัติสำหรับ PDF
def ensure_thai_font():
    if not os.path.exists(FONT_PATH):
        try:
            url = "https://github.com/google/fonts/raw/main/ofl/kanit/Kanit-Bold.ttf"
            urllib.request.urlretrieve(url, FONT_PATH)
        except Exception:
            pass
    if os.path.exists(FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont("KanitBold", FONT_PATH))
            return "KanitBold"
        except Exception:
            pass
    return "Helvetica-Bold"

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()

    # ตารางผู้ใช้งาน
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            display_name TEXT NOT NULL,
            role TEXT DEFAULT 'user',
            auth_provider TEXT DEFAULT 'local',
            consent_accepted INTEGER DEFAULT 1,
            created_at TEXT
        )
    """)

    # ตารางสินค้า
    c.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            barcode TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            price REAL DEFAULT 0,
            quantity INTEGER DEFAULT 0,
            unit TEXT DEFAULT 'ชิ้น',
            details TEXT DEFAULT ''
        )
    """)

    # ตารางประวัติการนับสต็อก
    c.execute("""
        CREATE TABLE IF NOT EXISTS stock_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            barcode TEXT NOT NULL,
            product_name TEXT NOT NULL,
            counted_qty INTEGER NOT NULL,
            recorder_name TEXT NOT NULL,
            username TEXT NOT NULL,
            recorded_at TEXT NOT NULL
        )
    """)

    # ตารางประวัติการส่งออกและการพิมพ์ + ไฟล์สำรอง (สำหรับแอดมิน)
    c.execute("""
        CREATE TABLE IF NOT EXISTS export_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            display_name TEXT NOT NULL,
            export_type TEXT NOT NULL,
            backup_filename TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)

    # ตารางรีวิว
    c.execute("""
        CREATE TABLE IF NOT EXISTS reviews (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            display_name TEXT NOT NULL,
            rating INTEGER NOT NULL,
            comment TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)

    # สร้างบัญชีผู้ดูแลระบบเริ่มต้น (admin / 1234)
    c.execute("SELECT id FROM users WHERE username = 'admin'")
    if not c.fetchone():
        c.execute("""
            INSERT INTO users (username, password, display_name, role, auth_provider, consent_accepted, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, ("admin", "1234", "ผู้ดูแลระบบ (Admin)", "admin", "local", 1, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

    # เพิ่มข้อมูลสินค้าตัวอย่างเริ่มต้นหากยังไม่มี
    c.execute("SELECT COUNT(*) FROM products")
    if c.fetchone()[0] == 0:
        sample_products = [
            ("8850001001", "น้ำดื่ม MJ ขนาด 600 มล.", 10.0, 120, "ขวด", "น้ำดื่มสะอาดบรรจุขวด แพ็ค 12 ขวด"),
            ("8850001002", "กาแฟสำเร็จรูป สูตรเข้มข้น", 85.0, 45, "แพ็ค", "บรรจุ 20 ซองต่อแพ็ค"),
            ("8850001003", "กระดาษถ่ายเอกสาร A4 80 แกรม", 115.0, 30, "รีม", "500 แผ่นต่อรีม")
        ]
        c.executemany("""
            INSERT INTO products (barcode, name, price, quantity, unit, details)
            VALUES (?, ?, ?, ?, ?, ?)
        """, sample_products)

    conn.commit()
    conn.close()

init_db()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

# เทมเพลตหลักที่มีแถบเมนูและคำว่า "ซอฟต์แวร์ โดย MJ" ทุกหน้า
BASE_HTML = """
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ระบบนับสต็อก - ซอฟต์แวร์ โดย MJ</title>
    <link href="https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        * { box-sizing: border-box; font-family: 'Kanit', sans-serif; }
        body { margin: 0; background: #f4f6f9; color: #222; display: flex; flex-direction: column; min-height: 100vh; }
        header { background: #1e293b; color: #fff; padding: 12px 20px; display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; }
        header .brand { font-size: 1.2rem; font-weight: 700; }
        header .brand span { font-size: 0.85rem; background: #3b82f6; padding: 2px 8px; border-radius: 4px; margin-left: 8px; }
        nav { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 6px; }
        nav a { color: #e2e8f0; text-decoration: none; padding: 6px 12px; border-radius: 6px; font-size: 0.95rem; background: #334155; }
        nav a:hover, nav a.active { background: #2563eb; color: #fff; }
        nav a.logout { background: #dc2626; }
        .container { max-width: 1050px; width: 100%; margin: 20px auto; padding: 0 16px; flex: 1; }
        .card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 20px; }
        h2 { margin-top: 0; color: #1e293b; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }
        label { display: block; font-weight: 600; margin-bottom: 4px; margin-top: 10px; }
        input[type="text"], input[type="password"], input[type="number"], textarea, select {
            width: 100%; padding: 10px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 1rem;
        }
        button, .btn {
            display: inline-block; padding: 10px 18px; border: none; border-radius: 6px; font-size: 1rem;
            cursor: pointer; font-weight: 600; text-decoration: none; text-align: center;
        }
        .btn-primary { background: #2563eb; color: #fff; }
        .btn-success { background: #16a34a; color: #fff; }
        .btn-warning { background: #d97706; color: #fff; }
        .btn-danger { background: #dc2626; color: #fff; }
        .btn-google { background: #fff; color: #333; border: 1px solid #cbd5e1; width: 100%; margin-top: 8px; }
        .alert { padding: 12px; border-radius: 6px; margin-bottom: 16px; background: #dbeafe; color: #1e40af; border-left: 4px solid #2563eb; }
        .alert-error { background: #fee2e2; color: #991b1b; border-left-color: #dc2626; }
        table { width: 100%; border-collapse: collapse; margin-top: 12px; }
        th, td { border: 1px solid #e2e8f0; padding: 10px; text-align: left; }
        th { background: #f1f5f9; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
        @media (max-width: 700px) { .grid-2 { grid-template-columns: 1fr; } }
        footer { background: #0f172a; color: #f8fafc; text-align: center; padding: 14px; font-weight: 600; letter-spacing: 0.5px; }
        @media print {
            header, nav, .no-print { display: none !important; }
            .print-watermark {
                position: fixed; top: 45%; left: 10%; transform: rotate(-30deg);
                font-size: 48px; font-weight: 700; color: rgba(150, 150, 150, 0.25);
                pointer-events: none; z-index: 9999; user-select: none;
            }
        }
    </style>
</head>
<body>
    <header>
        <div class="brand">ระบบจัดการและนับสต็อกสินค้า <span>ซอฟต์แวร์ โดย MJ</span></div>
        {% if session.get('user_id') %}
        <nav>
            <a href="{{ url_for('stock_count') }}" class="{{ 'active' if active_page=='stock' else '' }}">หน้าแรก / นับสต็อก</a>
            <a href="{{ url_for('products_page') }}" class="{{ 'active' if active_page=='products' else '' }}">ราคาและจำนวน</a>
            <a href="{{ url_for('export_page') }}" class="{{ 'active' if active_page=='export' else '' }}">ส่งออก / พิมพ์</a>
            <a href="{{ url_for('review_page') }}" class="{{ 'active' if active_page=='review' else '' }}">รีวิวและแชร์</a>
            <a href="{{ url_for('logout') }}" class="logout">ออกจากระบบ</a>
        </nav>
        {% endif %}
    </header>

    <div class="container">
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, msg in messages %}
                    <div class="alert {{ 'alert-error' if category == 'error' else '' }}">{{ msg }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        {{ content | safe }}
    </div>

    <footer>{{ footer_credit }} | ระบบนับสต็อก พัฒนาโดย MJ</footer>
</body>
</html>
"""

def render_page(content_html, active_page="", **context):
    inner = render_template_string(content_html, **context)
    return render_template_string(
        BASE_HTML,
        content=inner,
        active_page=active_page,
        footer_credit=FOOTER_CREDIT,
        watermark_text=WATERMARK_TEXT
    )

# ==========================================
# 1. หน้า Login / สร้างบัญชี / ลืมรหัสผ่าน / Google Auth
# ==========================================
@app.route("/", methods=["GET", "POST"])
@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("stock_count"))

    if request.method == "POST":
        action = request.form.get("action")
        consent = request.form.get("consent")

        # บังคับติ๊กให้ความยินยอมก่อนใช้งานทุกกรณี
        if not consent:
            flash("ไม่สามารถเข้าใช้งานเว็บไซต์ได้ กรุณาติ๊กถูกเพื่อให้ความยินยอมการเข้าถึงและใช้ข้อมูลก่อน", "error")
            return redirect(url_for("login"))

        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        display_name = request.form.get("display_name", "").strip() or username

        conn = get_db()
        c = conn.cursor()

        if action == "login":
            c.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
            user = c.fetchone()
            if user:
                session["user_id"] = user["id"]
                session["username"] = user["username"]
                session["display_name"] = user["display_name"]
                session["role"] = user["role"]
                conn.close()
                return redirect(url_for("stock_count"))
            else:
                flash("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง", "error")

        elif action == "register":
            if not username or not password:
                flash("กรุณากรอก User และ Password ให้ครบถ้วน", "error")
            else:
                try:
                    c.execute("""
                        INSERT INTO users (username, password, display_name, role, auth_provider, consent_accepted, created_at)
                        VALUES (?, ?, ?, 'user', 'local', 1, ?)
                    """, (username, password, display_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    conn.commit()
                    flash("สร้างบัญชีสำเร็จ สามารถเข้าสู่ระบบได้ทันที")
                except sqlite3.IntegrityError:
                    flash("ชื่อผู้ใช้นี้ถูกใช้งานแล้ว กรุณาใช้ชื่ออื่น", "error")

        elif action == "forgot":
            new_pass = request.form.get("new_password", "").strip()
            c.execute("SELECT id FROM users WHERE username = ?", (username,))
            if c.fetchone() and new_pass:
                c.execute("UPDATE users SET password = ? WHERE username = ?", (new_pass, username))
                conn.commit()
                flash("เปลี่ยนรหัสผ่านใหม่เรียบร้อยแล้ว")
            else:
                flash("ไม่พบบัญชีผู้ใช้ หรือยังไม่ได้กรอกรหัสผ่านใหม่", "error")

        elif action == "google_register":
            google_email = request.form.get("google_email", "").strip()
            google_name = request.form.get("google_name", "").strip() or google_email
            if not google_email:
                flash("กรุณากรอกอีเมล Google เพื่อสร้างบัญชีและเข้าสู่ระบบ", "error")
            else:
                c.execute("SELECT * FROM users WHERE username = ?", (google_email,))
                user = c.fetchone()
                if not user:
                    c.execute("""
                        INSERT INTO users (username, password, display_name, role, auth_provider, consent_accepted, created_at)
                        VALUES (?, 'GOOGLE_OAUTH', ?, 'user', 'google', 1, ?)
                    """, (google_email, google_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    conn.commit()
                    c.execute("SELECT * FROM users WHERE username = ?", (google_email,))
                    user = c.fetchone()
                session["user_id"] = user["id"]
                session["username"] = user["username"]
                session["display_name"] = user["display_name"]
                session["role"] = user["role"]
                conn.close()
                return redirect(url_for("stock_count"))

        conn.close()

    html = """
    <div style="max-width: 520px; margin: 20px auto;">
        <div class="card">
            <h2>เข้าสู่ระบบ / สร้างบัญชี</h2>
            <p style="font-size: 0.9rem; color: #475569;">บัญชีทดสอบผู้ดูแลระบบ: User <b>admin</b> / Password <b>1234</b></p>

            <form method="POST" id="authForm">
                <div style="background: #fef3c7; border: 1px solid #f59e0b; padding: 12px; border-radius: 6px; margin-bottom: 14px;">
                    <label style="margin:0; display:flex; align-items:flex-start; gap:8px; cursor:pointer; font-weight:600; color:#92400e;">
                        <input type="checkbox" name="consent" value="1" required style="width:20px; height:20px; margin-top:2px;">
                        <span>ข้าพเจ้ายินยอมให้ระบบเข้าถึงข้อมูลบัญชี ข้อมูลการนับสต็อก และบันทึกประวัติการใช้งาน (จำเป็นต้องติ๊กถูกเพื่อใช้งานเว็บไซต์)</span>
                    </label>
                </div>

                <label>ชื่อผู้ใช้งาน (User)</label>
                <input type="text" name="username" placeholder="กรอก User">

                <label>รหัสผ่าน (Password)</label>
                <input type="password" name="password" placeholder="กรอก Password">

                <label>ชื่อ-นามสกุลผู้ใช้งาน (สำหรับแสดงตอนบันทึกนับสต็อก - กรณีสร้างบัญชี)</label>
                <input type="text" name="display_name" placeholder="ระบุชื่อผู้ใช้งาน">

                <div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:16px;">
                    <button type="submit" name="action" value="login" class="btn-primary" style="flex:1;">เข้าสู่ระบบ</button>
                    <button type="submit" name="action" value="register" class="btn-success" style="flex:1;">สร้างบัญชี</button>
                </div>

                <div style="margin-top: 12px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
                    <details>
                        <summary style="cursor:pointer; color:#2563eb; font-weight:600;">ลืมรหัสผ่าน? คลิกที่นี่เพื่อตั้งรหัสผ่านใหม่</summary>
                        <label>รหัสผ่านใหม่</label>
                        <input type="password" name="new_password" placeholder="กรอกรหัสผ่านใหม่ที่ต้องการเปลี่ยน">
                        <button type="submit" name="action" value="forgot" class="btn-warning" style="margin-top:8px; width:100%;">ยืนยันเปลี่ยนรหัสผ่าน</button>
                    </details>
                </div>

                <div style="margin-top: 16px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
                    <p style="margin:0 0 8px 0; font-weight:600;">หรือสร้างบัญชีและเข้าสู่ระบบด้วยบัญชี Google (ขอความยินยอมการเข้าถึงข้อมูล)</p>
                    <input type="text" name="google_email" placeholder="กรอกอีเมล Google (เช่น name@gmail.com)" style="margin-bottom:6px;">
                    <input type="text" name="google_name" placeholder="ชื่อบัญชี Google">
                    <button type="submit" name="action" value="google_register" class="btn-google">G สร้างบัญชี / เข้าสู่ระบบด้วย Google</button>
                </div>
            </form>
        </div>
    </div>
    """
    return render_page(html)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ==========================================
# API ค้นหาข้อมูลสินค้าอัตโนมัติเมื่อสแกนบาร์โค้ด
# ==========================================
@app.route("/api/product/<barcode>")
@login_required
def api_get_product(barcode):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM products WHERE barcode = ?", (barcode.strip(),))
    row = c.fetchone()
    conn.close()
    if row:
        return jsonify({
            "found": True,
            "barcode": row["barcode"],
            "name": row["name"],
            "price": row["price"],
            "quantity": row["quantity"],
            "unit": row["unit"],
            "details": row["details"]
        })
    return jsonify({"found": False})

# ==========================================
# 2. หน้าแรก / หน้านับสต็อก
# ==========================================
@app.route("/stock", methods=["GET", "POST"])
@login_required
def stock_count():
    conn = get_db()
    c = conn.cursor()

    if request.method == "POST":
        barcode = request.form.get("barcode", "").strip()
        product_name = request.form.get("product_name", "").strip()
        counted_qty = int(request.form.get("counted_qty", 0) or 0)
        recorder_name = request.form.get("recorder_name", "").strip() or session["display_name"]

        if not barcode or not product_name:
            flash("กรุณากรอกหรือสแกนรหัสสินค้าและชื่อสินค้า", "error")
        else:
            # อัปเดตยอดสต็อกหรือเพิ่มสินค้าใหม่ถ้ายังไม่มีในฐานข้อมูล
            c.execute("SELECT id FROM products WHERE barcode = ?", (barcode,))
            existing = c.fetchone()
            if existing:
                c.execute("UPDATE products SET quantity = ?, name = ? WHERE barcode = ?", (counted_qty, product_name, barcode))
            else:
                c.execute("INSERT INTO products (barcode, name, price, quantity, unit) VALUES (?, ?, 0, ?, 'ชิ้น')",
                          (barcode, product_name, counted_qty))

            c.execute("""
                INSERT INTO stock_logs (barcode, product_name, counted_qty, recorder_name, username, recorded_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (barcode, product_name, counted_qty, recorder_name, session["username"], datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            flash(f"บันทึกการนับสต็อกสินค้า '{product_name}' จำนวน {counted_qty} เรียบร้อยแล้ว")

    c.execute("SELECT * FROM stock_logs ORDER BY id DESC LIMIT 20")
    logs = c.fetchall()
    conn.close()

    html = """
    <div class="card">
        <h2>หน้าแรก : นับสต็อกและตรวจสอบข้อมูลสินค้า</h2>
        <div style="background:#f8fafc; padding:10px 14px; border-radius:6px; border:1px solid #cbd5e1; margin-bottom:14px;">
            <b>ชื่อผู้ใช้งานที่เข้าสู่ระบบ:</b> {{ session['display_name'] }} ({{ session['username'] }})
        </div>

        <form method="POST">
            <div class="grid-2">
                <div>
                    <label>รหัสสินค้า (บาร์โค้ด) — สแกนหรือพิมพ์แล้วข้อมูลจะขึ้นอัตโนมัติ</label>
                    <input type="text" id="barcodeInput" name="barcode" placeholder="ยิงบาร์โค้ดหรือพิมพ์รหัสสินค้า" autofocus required>
                </div>
                <div>
                    <label>ชื่อสินค้า</label>
                    <input type="text" id="nameInput" name="product_name" placeholder="ชื่อสินค้าจะขึ้นอัตโนมัติเมื่อสแกน" required>
                </div>
            </div>

            <div class="grid-2">
                <div>
                    <label>จำนวนสินค้าที่นับได้</label>
                    <input type="number" id="qtyInput" name="counted_qty" value="1" min="0" required>
                </div>
                <div>
                    <label>ชื่อผู้บันทึก / ผู้นับสต็อก</label>
                    <input type="text" name="recorder_name" value="{{ session['display_name'] }}" required>
                </div>
            </div>

            <div id="autoProductInfo" style="margin-top:14px; padding:12px; background:#eff6ff; border-radius:6px; border:1px solid #93c5fd;">
                <b>ข้อมูลสินค้าอัตโนมัติ:</b> <span id="infoText">รอการสแกนบาร์โค้ดหรือกรอกรหัสสินค้า...</span>
            </div>

            <button type="submit" class="btn-success" style="margin-top:16px; width:100%;">บันทึก (การนับสต็อก)</button>
        </form>
    </div>

    <div class="card">
        <h2>ประวัติการนับสต็อกล่าสุด</h2>
        <table>
            <thead>
                <tr>
                    <th>เวลาที่บันทึก</th>
                    <th>รหัสสินค้า (บาร์โค้ด)</th>
                    <th>ชื่อสินค้า</th>
                    <th>จำนวนที่นับ</th>
                    <th>ชื่อผู้บันทึก/ผู้นับสต็อก</th>
                </tr>
            </thead>
            <tbody>
                {% for log in logs %}
                <tr>
                    <td>{{ log['recorded_at'] }}</td>
                    <td>{{ log['barcode'] }}</td>
                    <td>{{ log['product_name'] }}</td>
                    <td>{{ log['counted_qty'] }}</td>
                    <td>{{ log['recorder_name'] }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>

    <script>
        const barcodeInput = document.getElementById('barcodeInput');
        const nameInput = document.getElementById('nameInput');
        const qtyInput = document.getElementById('qtyInput');
        const infoText = document.getElementById('infoText');

        async function fetchProduct(code) {
            if (!code) return;
            const res = await fetch('/api/product/' + encodeURIComponent(code));
            const data = await res.json();
            if (data.found) {
                nameInput.value = data.name;
                qtyInput.value = data.quantity;
                infoText.innerHTML = `รหัสสินค้า: <b>${data.barcode}</b> | ชื่อสินค้า: <b>${data.name}</b> | ราคา: <b>${data.price.toFixed(2)} บาท</b> | จำนวนในระบบ: <b>${data.quantity} ${data.unit}</b>`;
            } else {
                infoText.innerHTML = `ไม่พบรหัสสินค้า <b>${code}</b> ในระบบ (สามารถกรอกชื่อและจำนวนเพื่อบันทึกใหม่ได้)`;
            }
        }

        barcodeInput.addEventListener('change', () => fetchProduct(barcodeInput.value));
        barcodeInput.addEventListener('keyup', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                fetchProduct(barcodeInput.value);
            }
        });
    </script>
    """
    return render_page(html, active_page="stock", logs=logs)

# ==========================================
# 3. หน้าราคาและจำนวน (เพิ่ม/ลบสินค้า เซ็ตราคา จำนวน รายละเอียด)
# ==========================================
@app.route("/products", methods=["GET", "POST"])
@login_required
def products_page():
    conn = get_db()
    c = conn.cursor()

    if request.method == "POST":
        action = request.form.get("action")
        if action == "save":
            barcode = request.form.get("barcode", "").strip()
            name = request.form.get("name", "").strip()
            price = float(request.form.get("price", 0) or 0)
            quantity = int(request.form.get("quantity", 0) or 0)
            unit = request.form.get("unit", "ชิ้น").strip() or "ชิ้น"
            details = request.form.get("details", "").strip()

            if barcode and name:
                c.execute("""
                    INSERT INTO products (barcode, name, price, quantity, unit, details)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(barcode) DO UPDATE SET
                        name=excluded.name,
                        price=excluded.price,
                        quantity=excluded.quantity,
                        unit=excluded.unit,
                        details=excluded.details
                """, (barcode, name, price, quantity, unit, details))
                conn.commit()
                flash(f"บันทึกข้อมูลสินค้า '{name}' เรียบร้อยแล้ว")
        elif action == "delete":
            prod_id = request.form.get("prod_id")
            c.execute("DELETE FROM products WHERE id = ?", (prod_id,))
            conn.commit()
            flash("ลบรายการสินค้าเรียบร้อยแล้ว")

    c.execute("SELECT * FROM products ORDER BY id DESC")
    products = c.fetchall()
    conn.close()

    html = """
    <div class="card">
        <h2>หน้าราคาและจำนวน : เพิ่ม / แก้ไขข้อมูลสินค้า</h2>
        <form method="POST">
            <input type="hidden" name="action" value="save">
            <div class="grid-2">
                <div>
                    <label>รหัสสินค้า (บาร์โค้ด)</label>
                    <input type="text" name="barcode" placeholder="เช่น 8850001004" required>
                </div>
                <div>
                    <label>ชื่อสินค้า</label>
                    <input type="text" name="name" placeholder="ระบุชื่อสินค้า" required>
                </div>
            </div>
            <div class="grid-2">
                <div>
                    <label>ราคา (บาท)</label>
                    <input type="number" step="0.01" name="price" value="0" required>
                </div>
                <div>
                    <label>จำนวนคงเหลือ</label>
                    <input type="number" name="quantity" value="0" required>
                </div>
            </div>
            <div class="grid-2">
                <div>
                    <label>หน่วยนับ (ต่อชิ้น / แพ็ค / กล่อง)</label>
                    <input type="text" name="unit" value="ชิ้น" required>
                </div>
                <div>
                    <label>รายละเอียดสินค้า</label>
                    <input type="text" name="details" placeholder="ระบุรายละเอียดเพิ่มเติมของสินค้า">
                </div>
            </div>
            <button type="submit" class="btn-primary" style="margin-top:14px;">เพิ่ม / บันทึกข้อมูลสินค้า</button>
        </form>
    </div>

    <div class="card">
        <h2>รายการสินค้าทั้งหมดในระบบ</h2>
        <table>
            <thead>
                <tr>
                    <th>รหัสสินค้า</th>
                    <th>ชื่อสินค้า</th>
                    <th>ราคา</th>
                    <th>จำนวน</th>
                    <th>รายละเอียดสินค้า</th>
                    <th>จัดการ</th>
                </tr>
            </thead>
            <tbody>
                {% for p in products %}
                <tr>
                    <td>{{ p['barcode'] }}</td>
                    <td>{{ p['name'] }}</td>
                    <td>{{ "%.2f"|format(p['price']) }}</td>
                    <td>{{ p['quantity'] }} {{ p['unit'] }}</td>
                    <td>{{ p['details'] }}</td>
                    <td>
                        <form method="POST" style="margin:0;">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="prod_id" value="{{ p['id'] }}">
                            <button type="submit" class="btn-danger" style="padding:6px 12px; font-size:0.85rem;">ลบสินค้า</button>
                        </form>
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
    """
    return render_page(html, active_page="products", products=products)

# ==========================================
# ฟังก์ชันสร้างไฟล์รายงานพร้อมลายน้ำบังคับ และสำรองไฟล์อัตโนมัติ
# ==========================================
def fetch_all_products():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM products ORDER BY barcode ASC")
    rows = c.fetchall()
    conn.close()
    return rows

def log_and_backup_export(export_type, file_bytes, ext):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_user = "".join(ch for ch in session["username"] if ch.isalnum()) or "user"
    backup_filename = f"backup_{safe_user}_{export_type}_{timestamp}.{ext}"
    backup_path = os.path.join(BACKUP_DIR, backup_filename)

    with open(backup_path, "wb") as f:
        f.write(file_bytes)

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO export_logs (username, display_name, export_type, backup_filename, created_at)
        VALUES (?, ?, ?, ?, ?)
    """, (session["username"], session["display_name"], export_type, backup_filename, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()

def build_pdf_bytes(products):
    font_name = ensure_thai_font()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    def draw_watermark_and_footer():
        # วาดลายน้ำตัวอักษรไม่มีหัวขนาดใหญ่เอียงกลางหน้ากระดาษ
        c.saveState()
        c.setFillColor(Color(0.6, 0.6, 0.6, alpha=0.28))
        c.setFont(font_name, 34)
        c.translate(width / 2, height / 2)
        c.rotate(35)
        c.drawCentredString(0, 0, WATERMARK_TEXT)
        c.restoreState()

        # ข้อความกำกับล่างหน้ากระดาษ
        c.setFillColor(Color(0.2, 0.2, 0.2, alpha=1))
        c.setFont(font_name, 10)
        c.drawString(40, 25, f"{FOOTER_CREDIT} | ผู้ส่งออก: {session['display_name']} | วันที่: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    draw_watermark_and_footer()
    c.setFont(font_name, 18)
    c.drawString(40, height - 50, "รายงานสรุปรายการสต็อกสินค้า")

    y = height - 85
    c.setFont(font_name, 11)
    c.drawString(40, y, "รหัสสินค้า")
    c.drawString(145, y, "ชื่อสินค้า")
    c.drawString(355, y, "ราคา")
    c.drawString(430, y, "จำนวน")
    y -= 8
    c.line(40, y, width - 40, y)
    y -= 20

    for p in products:
        if y < 70:
            c.showPage()
            draw_watermark_and_footer()
            y = height - 60
            c.setFont(font_name, 11)
        c.drawString(40, y, str(p["barcode"]))
        c.drawString(145, y, str(p["name"])[:32])
        c.drawString(355, y, f"{p['price']:.2f}")
        c.drawString(430, y, f"{p['quantity']} {p['unit']}")
        y -= 20

    c.save()
    return buf.getvalue()

def build_word_bytes(products):
    doc = docx.Document()
    style = doc.styles["Normal"]
    style.font.name = "Kanit"
    style.font.size = Pt(12)

    # ใส่ลายน้ำใน Header ของเอกสาร Word (ขนาดใหญ่ ฟอนต์ไม่มีหัว)
    section = doc.sections[0]
    header = section.header
    hp = header.paragraphs[0]
    hrun = hp.add_run(WATERMARK_TEXT)
    hrun.font.name = "Kanit"
    hrun.font.size = Pt(28)
    hrun.font.bold = True
    hrun.font.color.rgb = RGBColor(180, 180, 180)

    # ล็อกป้องกันการแก้ไขเอกสารใน Word (Read-only protection)
    settings_element = doc.settings.element
    protection = parse_xml(f'<w:documentProtection {nsdecls("w")} w:edit="readOnly" w:enforcement="1"/>')
    settings_element.append(protection)

    title = doc.add_heading("รายงานสรุปรายการสต็อกสินค้า", level=1)
    doc.add_paragraph(f"{FOOTER_CREDIT} | ผู้ส่งออก: {session['display_name']} | วันที่: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "รหัสสินค้า"
    hdr[1].text = "ชื่อสินค้า"
    hdr[2].text = "ราคา"
    hdr[3].text = "จำนวน"
    hdr[4].text = "รายละเอียด"

    for p in products:
        row = table.add_row().cells
        row[0].text = str(p["barcode"])
        row[1].text = str(p["name"])
        row[2].text = f"{p['price']:.2f}"
        row[3].text = f"{p['quantity']} {p['unit']}"
        row[4].text = str(p["details"] or "")

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.add_run(f"{FOOTER_CREDIT} — {WATERMARK_TEXT}")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def build_excel_bytes(products):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "StockReport"

    # แถวลายน้ำขนาดใหญ่ด้านบนและตั้งค่า Header/Footer สำหรับพิมพ์
    ws.merge_cells("A1:E1")
    ws["A1"] = WATERMARK_TEXT
    ws["A1"].font = Font(name="Kanit", size=24, bold=True, color="999999")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42

    ws.merge_cells("A2:E2")
    ws["A2"] = f"{FOOTER_CREDIT} | ผู้ส่งออก: {session['display_name']} | วันที่: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Kanit", size=11, bold=True)

    ws.oddHeader.center.text = f"&\"Kanit,Bold\"&24&K999999 {WATERMARK_TEXT}"
    ws.oddFooter.center.text = f"&\"Kanit,Bold\"&11 {FOOTER_CREDIT}"

    headers = ["รหัสสินค้า", "ชื่อสินค้า", "ราคา", "จำนวน", "รายละเอียด"]
    ws.append(headers)
    for col_idx in range(1, 6):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = Font(name="Kanit", size=12, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")

    for p in products:
        ws.append([p["barcode"], p["name"], p["price"], f"{p['quantity']} {p['unit']}", p["details"]])

    # เปิดการป้องกันแผ่นงาน (Sheet Protection) เพื่อไม่ให้แก้ไขลายน้ำได้
    ws.protection.sheet = True
    ws.protection.password = "mjstock2026"
    ws.protection.enable()

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ==========================================
# 4. หน้าส่งออก (PDF / Word / Excel / Print) + ประวัติสำหรับ Admin
# ==========================================
@app.route("/export")
@login_required
def export_page():
    products = fetch_all_products()
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM export_logs ORDER BY id DESC LIMIT 50")
    export_logs = c.fetchall()
    conn.close()

    html = """
    <div class="print-watermark">{{ watermark_text }}</div>

    <div class="card no-print">
        <h2>หน้าส่งออกและพิมพ์เอกสาร</h2>
        <p>ไฟล์ที่ส่งออกและเอกสารที่สั่งพิมพ์ทุกฉบับจะมีลายน้ำบังคับ <b>"{{ watermark_text }}"</b> (ฟอนต์ไม่มีหัว ขนาดใหญ่ ป้องกันการแก้ไข) และระบบจะสำรองไฟล์อัตโนมัติทุกครั้ง</p>
        <div style="display:flex; gap:10px; flex-wrap:wrap; margin-top:14px;">
            <a href="{{ url_for('download_export', fmt='pdf') }}" class="btn btn-danger">ส่งออกเป็น PDF</a>
            <a href="{{ url_for('download_export', fmt='word') }}" class="btn btn-primary">ส่งออกเป็น Word (.docx)</a>
            <a href="{{ url_for('download_export', fmt='excel') }}" class="btn btn-success">ส่งออกเป็น Excel (.xlsx)</a>
            <button type="button" onclick="triggerPrint()" class="btn btn-warning">พิมพ์เอกสาร (Print)</button>
        </div>
    </div>

    <div class="card">
        <h2>ตัวอย่างรายงานสต็อกสินค้า ({{ watermark_text }})</h2>
        <table>
            <thead>
                <tr>
                    <th>รหัสสินค้า</th>
                    <th>ชื่อสินค้า</th>
                    <th>ราคา</th>
                    <th>จำนวน</th>
                    <th>รายละเอียด</th>
                </tr>
            </thead>
            <tbody>
                {% for p in products %}
                <tr>
                    <td>{{ p['barcode'] }}</td>
                    <td>{{ p['name'] }}</td>
                    <td>{{ "%.2f"|format(p['price']) }}</td>
                    <td>{{ p['quantity'] }} {{ p['unit'] }}</td>
                    <td>{{ p['details'] }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>

    {% if session.get('role') == 'admin' %}
    <div class="card no-print" style="border: 2px solid #2563eb;">
        <h2>ส่วนผู้ดูแลระบบ (Admin) : ประวัติการพิมพ์ / ส่งออก และไฟล์สำรอง</h2>
        <table>
            <thead>
                <tr>
                    <th>วัน-เวลา</th>
                    <th>ผู้ใช้งาน</th>
                    <th>รูปแบบ</th>
                    <th>ไฟล์สำรองในระบบ</th>
                </tr>
            </thead>
            <tbody>
                {% for el in export_logs %}
                <tr>
                    <td>{{ el['created_at'] }}</td>
                    <td>{{ el['display_name'] }} ({{ el['username'] }})</td>
                    <td>{{ el['export_type'] }}</td>
                    <td>
                        <a href="{{ url_for('download_backup', filename=el['backup_filename']) }}">{{ el['backup_filename'] }}</a>
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
    {% endif %}

    <script>
        async function triggerPrint() {
            await fetch('/api/log_print', { method: 'POST' });
            window.print();
        }
    </script>
    """
    return render_page(html, active_page="export", products=products, export_logs=export_logs)

@app.route("/export/<fmt>")
@login_required
def download_export(fmt):
    products = fetch_all_products()
    if fmt == "pdf":
        data = build_pdf_bytes(products)
        log_and_backup_export("PDF", data, "pdf")
        return send_file(io.BytesIO(data), as_attachment=True, download_name="MJ_Stock_Report.pdf", mimetype="application/pdf")
    elif fmt == "word":
        data = build_word_bytes(products)
        log_and_backup_export("WORD", data, "docx")
        return send_file(io.BytesIO(data), as_attachment=True, download_name="MJ_Stock_Report.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    elif fmt == "excel":
        data = build_excel_bytes(products)
        log_and_backup_export("EXCEL", data, "xlsx")
        return send_file(io.BytesIO(data), as_attachment=True, download_name="MJ_Stock_Report.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return redirect(url_for("export_page"))

@app.route("/api/log_print", methods=["POST"])
@login_required
def api_log_print():
    products = fetch_all_products()
    data = build_pdf_bytes(products)
    log_and_backup_export("PRINT", data, "pdf")
    return jsonify({"status": "ok"})

@app.route("/backup/<filename>")
@login_required
def download_backup(filename):
    if session.get("role") != "admin":
        flash("เฉพาะผู้ดูแลระบบเท่านั้นที่เข้าถึงไฟล์สำรองได้", "error")
        return redirect(url_for("export_page"))
    safe_name = os.path.basename(filename)
    path = os.path.join(BACKUP_DIR, safe_name)
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=safe_name)
    flash("ไม่พบไฟล์สำรองดังกล่าว", "error")
    return redirect(url_for("export_page"))

# ==========================================
# 5. หน้ารีวิว (ให้คะแนน 1-5 ดาว และปุ่มแชร์ประชาสัมพันธ์)
# ==========================================
@app.route("/reviews", methods=["GET", "POST"])
@login_required
def review_page():
    conn = get_db()
    c = conn.cursor()

    if request.method == "POST":
        rating = int(request.form.get("rating", 5))
        comment = request.form.get("comment", "").strip()
        if comment:
            c.execute("""
                INSERT INTO reviews (display_name, rating, comment, created_at)
                VALUES (?, ?, ?, ?)
            """, (session["display_name"], rating, comment, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            flash("ขอบคุณสำหรับรีวิวและคะแนนความพึงพอใจครับ")

    c.execute("SELECT * FROM reviews ORDER BY id DESC")
    reviews = c.fetchall()
    conn.close()

    html = """
    <div class="card">
        <h2>หน้ารีวิวและให้คะแนนความพึงพอใจ</h2>
        <form method="POST">
            <label>ให้คะแนนความพึงพอใจ (1 - 5 ดาว)</label>
            <select name="rating" required>
                <option value="5">★★★★★ (5 ดาว - ดีเยี่ยม)</option>
                <option value="4">★★★★☆ (4 ดาว - ดีมาก)</option>
                <option value="3">★★★☆☆ (3 ดาว - ปานกลาง)</option>
                <option value="2">★★☆☆☆ (2 ดาว - พอใช้)</option>
                <option value="1">★☆☆☆☆ (1 ดาว - ควรปรับปรุง)</option>
            </select>

            <label>ข้อความรีวิว / ข้อเสนอแนะ</label>
            <textarea name="comment" rows="3" placeholder="เขียนรีวิวการใช้งานระบบนับสต็อก..." required></textarea>

            <button type="submit" class="btn-primary" style="margin-top:12px;">ส่งรีวิว</button>
        </form>

        <div style="margin-top:20px; padding-top:16px; border-top:1px solid #e2e8f0;">
            <h3 style="margin-top:0;">แชร์ประชาสัมพันธ์ซอฟต์แวร์ไปยังช่องทางอื่นๆ</h3>
            <div style="display:flex; gap:10px; flex-wrap:wrap;">
                <button type="button" class="btn-primary" onclick="shareSocial('facebook')">แชร์ไปยัง Facebook</button>
                <button type="button" class="btn-success" onclick="shareSocial('line')">แชร์ไปยัง LINE</button>
                <button type="button" class="btn-warning" onclick="shareSocial('copy')">คัดลอกลิงก์เว็บไซต์</button>
            </div>
        </div>
    </div>

    <div class="card">
        <h2>ความคิดเห็นจากผู้ใช้งาน</h2>
        {% for r in reviews %}
        <div style="border-bottom:1px solid #e2e8f0; padding:10px 0;">
            <b>{{ r['display_name'] }}</b> —
            <span style="color:#f59e0b; font-size:1.1rem;">{{ '★' * r['rating'] }}{{ '☆' * (5 - r['rating']) }}</span>
            <small style="color:#64748b;">({{ r['created_at'] }})</small>
            <p style="margin:6px 0 0 0;">{{ r['comment'] }}</p>
        </div>
        {% else %}
        <p style="color:#64748b;">ยังไม่มีรีวิว เป็นคนแรกที่รีวิวระบบได้เลย</p>
        {% endfor %}
    </div>

    <script>
        function shareSocial(platform) {
            const url = encodeURIComponent(window.location.origin);
            const text = encodeURIComponent('แนะนำระบบนับสต็อกสินค้าออนไลน์ ใช้งานง่าย ฟรี - ซอฟต์แวร์ โดย MJ');
            if (platform === 'facebook') {
                window.open(`https://www.facebook.com/sharer/sharer.php?u=${url}`, '_blank');
            } else if (platform === 'line') {
                window.open(`https://social-plugins.line.me/lineit/share?url=${url}&text=${text}`, '_blank');
            } else {
                navigator.clipboard.writeText(window.location.origin);
                alert('คัดลอกลิงก์เว็บไซต์เรียบร้อยแล้ว สามารถนำไปวางแชร์ต่อได้ทันที');
            }
        }
    </script>
    """
    return render_page(html, active_page="review", reviews=reviews)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)