import os
import sqlite3
from datetime import datetime, date
from io import BytesIO
from functools import wraps
import hmac
from flask import jsonify
import qrcode

from flask import Flask, request, redirect, url_for, session, flash, render_template_string, send_file
from werkzeug.security import check_password_hash, generate_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "attendance.db")

EVENT_TITLE = "THE ANNUAL CONFERENCE 2026"
EVENT_THEME = "Delivering Reliable Digital Infrastructure for Meaningful Connectivity"
EVENT_DATE = "Selasa, 24 November 2026"
EVENT_TIME = "08.00–13.00 WITA"
EVENT_PLACE = "Merumatta Senggigi Lombok, Pantai Senggigi, Batu Layar, Lombok Barat"

# Ganti ADMIN_PASSWORD sebelum aplikasi dipublikasikan.
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "GantiPasswordAdmin2026!")
SECRET_KEY = os.getenv("SECRET_KEY", "ganti-secret-key-acak-sebelum-deploy")
ENFORCE_EVENT_DATE = os.getenv("ENFORCE_EVENT_DATE", "0") == "1"

app = Flask(__name__)
app.secret_key = SECRET_KEY


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                institution TEXT NOT NULL,
                position TEXT,
                phone TEXT,
                email TEXT,
                created_at TEXT NOT NULL,
                attendance_date TEXT NOT NULL
            )
        """)
        # Migrasi otomatis untuk database versi sebelumnya.
        columns = [row["name"] for row in conn.execute("PRAGMA table_info(attendance)").fetchall()]
        if "attendance_type" not in columns:
            conn.execute("ALTER TABLE attendance ADD COLUMN attendance_type TEXT NOT NULL DEFAULT 'Luring'")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_attendance_name ON attendance(full_name)")


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))
        return fn(*args, **kwargs)
    return wrapper


PUBLIC_TEMPLATE = r"""
<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Absensi | Annual Conference 2026</title>
<style>
:root{--gold:#e9bd55;--blue:#08275d;--cyan:#35c5ee}
*{box-sizing:border-box}
body{margin:0;min-height:100vh;font-family:Arial,Helvetica,sans-serif;color:#fff;
 background:linear-gradient(110deg,rgba(2,14,43,.88),rgba(2,24,66,.76)),url('{{ url_for("static",filename="backdrop.png") }}') center/cover fixed no-repeat}
.wrap{width:min(680px,94%);margin:30px auto;padding:26px}
.panel{background:rgba(3,20,55,.88);border:1px solid rgba(233,189,85,.55);border-radius:20px;
 box-shadow:0 18px 55px rgba(0,0,0,.35);backdrop-filter:blur(8px);padding:30px}
.brand{text-align:center;margin-bottom:24px}
.brand .eyebrow{color:var(--gold);font-weight:700;letter-spacing:2px;font-size:12px}
h1{font-family:Georgia,serif;text-align:center;font-size:clamp(25px,5vw,38px);line-height:1.12;margin:8px 0;color:#fff}
.theme{text-align:center;color:#f5d17a;font-weight:700;line-height:1.45;margin:12px auto 18px;max-width:550px}
.meta{display:grid;grid-template-columns:1fr;gap:7px;text-align:center;color:#dce9ff;font-size:14px;margin-bottom:24px}
label{display:block;font-size:13px;font-weight:700;color:#e9c86e;margin:15px 0 7px}
input,select{width:100%;padding:13px 14px;border-radius:9px;border:1px solid #9bb3d6;background:#f7faff;color:#10264b;font-size:15px}
button{width:100%;border:0;border-radius:10px;padding:15px;margin-top:23px;background:linear-gradient(90deg,#d7a638,#ffe08b);
 color:#09234f;font-size:16px;font-weight:800;cursor:pointer}
button:hover{filter:brightness(1.06)}
.note{font-size:12px;color:#c6d7f2;text-align:center;line-height:1.5;margin-top:16px}
.flash{padding:12px;border-radius:9px;margin:12px 0;background:#163d70;border:1px solid #5da9d4}
.flash.error{background:#642d36;border-color:#e58c8c}
.footer{text-align:center;color:#c6d7f2;font-size:11px;margin:18px}
a{color:#f6d477}
@media(max-width:520px){.wrap{padding:12px 0}.panel{padding:22px 18px}.meta{font-size:13px}}
</style>
</head>
<body>
<div class="wrap"><main class="panel">
  <div class="brand"><div class="eyebrow">BALMON MATARAM</div>
  <h1>ABSENSI DIGITAL</h1><div class="eyebrow">THE ANNUAL CONFERENCE 2026</div></div>
  <div class="theme">{{ theme }}</div>
  <div class="meta"><div>📅 {{ event_date }} &nbsp; | &nbsp; 🕗 {{ event_time }}</div><div>📍 {{ event_place }}</div></div>
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category,message in messages %}<div class="flash {{ 'error' if category=='error' else '' }}">{{ message }}</div>{% endfor %}
  {% endwith %}
  <form method="post" autocomplete="on">
    <label for="full_name">Nama lengkap *</label>
    <input id="full_name" name="full_name" required maxlength="120" placeholder="Masukkan nama lengkap">
    <label for="institution">Instansi / Unit Kerja *</label>
    <input id="institution" name="institution" required maxlength="160" placeholder="Contoh: Balmon Mataram">
    <label for="attendance_type">Jenis Kehadiran *</label>
    <select id="attendance_type" name="attendance_type" required>
      <option value="Luring" selected>Luring (Hadir di lokasi acara)</option>
      <option value="Daring">Daring (Mengikuti secara online)</option>
    </select>
    <label for="position">Jabatan (opsional)</label>
    <input id="position" name="position" maxlength="120" placeholder="Jabatan">
    <label for="phone">Nomor WhatsApp (opsional)</label>
    <input id="phone" name="phone" maxlength="30" placeholder="08xxxxxxxxxx">
    <label for="email">Email (opsional)</label>
    <input id="email" name="email" type="email" maxlength="160" placeholder="nama@email.com">
    <button type="submit">KONFIRMASI KEHADIRAN</button>
  </form>
  <div class="note">Dengan menekan tombol di atas, Anda menyatakan bahwa data yang diisi benar dan menyetujui penggunaannya untuk keperluan administrasi kehadiran acara.</div>
</main><div class="footer">Meaningful Connectivity • Balmon Mataram</div></div>
</body></html>
"""

LOGIN_TEMPLATE = r"""
<!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin Absensi</title><style>
body{font-family:Arial;background:#071b43;color:#fff;display:grid;place-items:center;min-height:100vh;margin:0}
form{background:#102c5d;padding:30px;border:1px solid #d9b34e;border-radius:16px;width:min(360px,90%)}
input{width:100%;padding:12px;margin:8px 0 16px;box-sizing:border-box}button{padding:12px;width:100%;background:#e9c15d;border:0;font-weight:bold}
a{color:#f4d47d}.msg{color:#ffd1d1}
</style></head><body><form method="post"><h2>Admin Absensi</h2>
{% with messages = get_flashed_messages() %}{% for m in messages %}<p class="msg">{{m}}</p>{% endfor %}{% endwith %}
<label>Username</label><input name="username" required>
<label>Password</label><input name="password" type="password" required>
<button>MASUK</button><p><a href="{{url_for('index')}}">Kembali ke halaman absensi</a></p></form></body></html>
"""


ADMIN_TEMPLATE = r"""
<!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Live Attendance Dashboard</title><style>
*{box-sizing:border-box}body{font-family:Arial;margin:0;background:#eef2f8;color:#12264a}
header{background:#071e4a;color:white;padding:20px 4%;display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap}
header a{color:#ffe08b}main{width:94%;max-width:1250px;margin:24px auto}
.card{background:white;padding:20px;border-radius:12px;box-shadow:0 4px 18px #10264b12;margin-bottom:18px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px}
.stat{font-size:34px;font-weight:bold;color:#0c397b}.muted{color:#61718a;font-size:13px}
.tools{display:flex;gap:10px;flex-wrap:wrap;margin:12px 0}input{padding:10px;border:1px solid #ccd5e3;border-radius:7px}
.btn{display:inline-block;padding:10px 14px;border-radius:7px;background:#0c397b;color:white;text-decoration:none;border:0;cursor:pointer}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{text-align:left;padding:11px;border-bottom:1px solid #e5eaf1}th{background:#f1f5fb}
.tablewrap{overflow:auto}.topgrid{display:grid;grid-template-columns:1.5fr 1fr;gap:18px}
.qrbox{display:flex;gap:18px;align-items:center;flex-wrap:wrap}.qrbox img{width:190px;height:190px;object-fit:contain;border:1px solid #e2e8f0;border-radius:10px;padding:8px}
.live{display:inline-flex;align-items:center;gap:7px;color:#14804a;font-size:12px;font-weight:bold}
.dot{width:9px;height:9px;border-radius:50%;background:#16a05d;display:inline-block}
@media(max-width:760px){.topgrid{grid-template-columns:1fr}.qrbox img{width:160px;height:160px}}
</style></head><body>
<header><div><b>WANDAA • LIVE ATTENDANCE DASHBOARD</b><div class="muted" style="color:#c8d6ed">The Annual Conference 2026</div></div>
<div><a href="{{url_for('index')}}">Halaman peserta</a> &nbsp; | &nbsp; <a href="{{url_for('logout')}}">Keluar</a></div></header>
<main>
<div class="topgrid">
  <div class="card">
    <div class="muted">MONITORING KEHADIRAN</div><h2 style="margin:7px 0">Annual Conference 2026</h2>
    <div class="muted">{{ event_date }} • {{ event_time }}<br>{{ event_place }}</div>
    <p class="live"><span class="dot"></span> LIVE — pembaruan otomatis setiap 5 detik</p>
    <div class="stats">
      <div><div class="muted">TOTAL HADIR</div><div class="stat" id="totalCount">{{ total }}</div></div>
      <div><div class="muted">CHECK-IN HARI INI</div><div class="stat" id="todayCount">0</div></div>
    </div>
  </div>
  <div class="card">
    <h3 style="margin-top:0">QR Code Check-in</h3>
    <div class="qrbox"><img src="{{url_for('admin_qr')}}" alt="QR Code absensi">
      <div><p class="muted">Peserta scan QR Code ini untuk membuka form absensi melalui HP.</p>
      <a class="btn" href="{{url_for('admin_qr', download=1)}}">Unduh QR Code</a>
      <p class="muted" style="margin-top:12px">Alamat check-in:<br><b style="overflow-wrap:anywhere">{{ public_url }}</b></p>
      </div></div>
  </div>
</div>
<div class="card"><h3>Daftar Kehadiran</h3>
<form class="tools" method="get"><input name="q" value="{{q}}" placeholder="Cari nama / instansi">
<button class="btn" type="submit">Cari</button><a class="btn" href="{{url_for('export_excel')}}">Unduh Excel</a></form>
<div class="tablewrap"><table><thead><tr><th>No.</th><th>Waktu Check-in</th><th>Nama</th><th>Instansi / Unit</th><th>Jenis Kehadiran</th><th>Jabatan</th><th>WhatsApp</th><th>Email</th></tr></thead>
<tbody id="attendanceRows">
{% for r in rows %}<tr><td>{{loop.index}}</td><td>{{r['created_at']}}</td><td>{{r['full_name']}}</td><td>{{r['institution']}}</td><td>{{r['attendance_type']}}</td><td>{{r['position'] or '-'}}</td><td>{{r['phone'] or '-'}}</td><td>{{r['email'] or '-'}}</td></tr>
{% else %}<tr><td colspan="8">Belum ada data kehadiran.</td></tr>{% endfor %}
</tbody></table></div></div>
</main>
<script>
async function refreshAttendance(){
  try{
    const response = await fetch("{{url_for('admin_live_data')}}", {cache:"no-store"});
    if(!response.ok) return;
    const data = await response.json();
    document.getElementById("totalCount").textContent = data.total;
    document.getElementById("todayCount").textContent = data.today;
    const body = document.getElementById("attendanceRows");
    body.innerHTML = "";
    if(!data.rows.length){
      body.innerHTML = '<tr><td colspan="8">Belum ada data kehadiran.</td></tr>';
      return;
    }
    data.rows.forEach((r,i)=>{
      const tr=document.createElement("tr");
      [i+1,r.created_at,r.full_name,r.institution,r.attendance_type,r.position||"-",r.phone||"-",r.email||"-"].forEach(value=>{
        const td=document.createElement("td"); td.textContent=value; tr.appendChild(td);
      });
      body.appendChild(tr);
    });
  }catch(e){ console.warn("Gagal memperbarui dashboard",e); }
}
setInterval(refreshAttendance,5000);
</script>
</body></html>
"""



@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        institution = request.form.get("institution", "").strip()
        attendance_type = request.form.get("attendance_type", "Luring").strip()
        if attendance_type not in ("Luring", "Daring"):
            attendance_type = "Luring"
        position = request.form.get("position", "").strip()
        phone = request.form.get("phone", "").strip()
        email = request.form.get("email", "").strip().lower()

        if not full_name or not institution:
            flash("Nama lengkap dan instansi/unit kerja wajib diisi.", "error")
            return redirect(url_for("index"))

        if ENFORCE_EVENT_DATE and date.today().isoformat() != "2026-11-24":
            flash("Check-in hanya dibuka pada 24 November 2026.", "error")
            return redirect(url_for("index"))

        # Cegah absensi ganda berdasarkan nama + instansi (case-insensitive).
        with db() as conn:
            duplicate = conn.execute("""
                SELECT id FROM attendance
                WHERE lower(trim(full_name)) = lower(trim(?))
                AND lower(trim(institution)) = lower(trim(?))
            """, (full_name, institution)).fetchone()
            if duplicate:
                flash("Data kehadiran dengan nama dan instansi tersebut sudah tercatat.", "error")
                return redirect(url_for("index"))

            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute("""
                INSERT INTO attendance
                (full_name,institution,attendance_type,position,phone,email,created_at,attendance_date)
                VALUES (?,?,?,?,?,?,?,?)
            """, (full_name, institution, attendance_type, position, phone, email, now, date.today().isoformat()))

        return render_template_string("""
        <!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
        <title>Check-in Berhasil</title><style>
        *{box-sizing:border-box}
        body{font-family:Arial,Helvetica,sans-serif;background:radial-gradient(circle at top,#173f79,#061638 70%);color:white;min-height:100vh;display:grid;place-items:center;margin:0;padding:20px}
        .box{width:min(760px,96%);padding:30px 26px 24px;text-align:center;border:1px solid #e8bf58;border-radius:20px;background:rgba(11,40,88,.94);box-shadow:0 18px 55px rgba(0,0,0,.35);overflow:hidden}
        h1{color:#f4d477;font-size:clamp(25px,5vw,36px);margin:4px 0 12px}
        .check{font-size:44px;color:#f4d477;line-height:1}
        .thanks{font-size:17px;line-height:1.5;margin:10px 0 5px}
        .sub{color:#d5e3f8;line-height:1.5;margin:5px 0 18px}
        .greeting{display:flex;justify-content:center;align-items:flex-end;gap:10px;margin:6px auto 0;max-width:580px;height:300px;overflow:hidden}
        .greeting img{display:block;width:50%;height:100%;object-fit:contain;object-position:center bottom}
        .event{color:#f4d477;font-size:12px;letter-spacing:1.5px;font-weight:bold;margin-top:10px}
        a{display:inline-block;color:#071e4a;background:linear-gradient(90deg,#d7a638,#ffe08b);padding:12px 22px;border-radius:9px;text-decoration:none;font-weight:bold;margin-top:15px}
        @media(max-width:520px){.box{padding:24px 14px 20px}.greeting{height:220px;gap:0}.thanks{font-size:15px}}
        </style></head><body><main class="box">
        <div class="check">✓</div><h1>Kehadiran Tercatat!</h1>
        <p class="thanks">Terima kasih, <b>{{name}}</b>.</p>
        <p class="sub">Registrasi kehadiran Anda berhasil. Selamat mengikuti<br><b>The Annual Conference 2026</b>.</p>
        <div class="greeting">
          <img src="{{url_for('static',filename='greeting_man.png')}}" alt="Sambutan panitia">
          <img src="{{url_for('static',filename='greeting_woman.png')}}" alt="Sambutan panitia">
        </div>
        <div class="event">DELIVERING RELIABLE DIGITAL INFRASTRUCTURE FOR MEANINGFUL CONNECTIVITY</div>
        <a href="{{url_for('index')}}">Kembali ke Halaman Absensi</a>
        </main></body></html>
        """, name=full_name)

    return render_template_string(PUBLIC_TEMPLATE, theme=EVENT_THEME, event_date=EVENT_DATE,
                                  event_time=EVENT_TIME, event_place=EVENT_PLACE)


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if request.form.get("username") == ADMIN_USERNAME and hmac.compare_digest(
                request.form.get("password", ""), ADMIN_PASSWORD):
            session["admin_logged_in"] = True
            return redirect(url_for("admin_dashboard"))
        flash("Username atau password salah.")
    return render_template_string(LOGIN_TEMPLATE)


@app.route("/admin")
@admin_required
def admin_dashboard():
    q = request.args.get("q", "").strip()
    with db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM attendance").fetchone()[0]
        if q:
            rows = conn.execute("""
                SELECT * FROM attendance
                WHERE full_name LIKE ? OR institution LIKE ?
                ORDER BY id DESC
            """, (f"%{q}%", f"%{q}%")).fetchall()
        else:
            rows = conn.execute("SELECT * FROM attendance ORDER BY id DESC").fetchall()
    return render_template_string(ADMIN_TEMPLATE, rows=rows, total=total, q=q,
                                  event_date=EVENT_DATE, event_time=EVENT_TIME, event_place=EVENT_PLACE, public_url=request.url_root)


@app.route("/admin/qr.png")
@admin_required
def admin_qr():
    # QR Code berisi alamat form absensi berdasarkan host yang sedang diakses admin.
    checkin_url = request.url_root
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=12,
        border=4,
    )
    qr.add_data(checkin_url)
    qr.make(fit=True)
    image = qr.make_image(fill_color="#071e4a", back_color="white")
    output = BytesIO()
    image.save(output, format="PNG")
    output.seek(0)
    return send_file(output, mimetype="image/png",
                     as_attachment=request.args.get("download") == "1",
                     download_name="QR_Checkin_Annual_Conference_2026.png")


@app.route("/admin/live-data")
@admin_required
def admin_live_data():
    today = date.today().isoformat()
    with db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM attendance").fetchone()[0]
        today_total = conn.execute(
            "SELECT COUNT(*) FROM attendance WHERE attendance_date = ?", (today,)
        ).fetchone()[0]
        rows = conn.execute("""
            SELECT created_at,full_name,institution,attendance_type,position,phone,email
            FROM attendance ORDER BY id DESC LIMIT 100
        """).fetchall()
    return jsonify({
        "total": total,
        "today": today_total,
        "rows": [dict(r) for r in rows]
    })


@app.route("/admin/export")
@admin_required
def export_excel():
    with db() as conn:
        rows = conn.execute("""
            SELECT id,created_at,full_name,institution,attendance_type,position,phone,email
            FROM attendance ORDER BY id ASC
        """).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Daftar Kehadiran"
    ws.append(["No.", "Waktu Check-in", "Nama Lengkap", "Instansi / Unit Kerja", "Jenis Kehadiran",
               "Jabatan", "Nomor WhatsApp", "Email"])
    for i, r in enumerate(rows, 1):
        ws.append([i, r["created_at"], r["full_name"], r["institution"], r["attendance_type"],
                   r["position"], r["phone"], r["email"]])

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="12366B")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 14), 36)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True,
                     download_name="Absensi_Annual_Conference_2026.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/logout")
def logout():
    session.clear()
    return redirect(url_for("admin_login"))


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=False)
