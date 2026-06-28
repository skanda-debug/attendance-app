"""
Dynamic QR Code Attendance System - WEB APP VERSION
Deployable on Render / Railway / any cloud host, or run locally + ngrok.

Opening the root URL "/" goes DIRECTLY to the live QR code page.
No Tkinter, no desktop GUI - everything runs in the browser.
"""

import qrcode
import threading
import time
import os
import secrets
import string
import io
import base64
from datetime import datetime
from calendar import monthrange
from flask import Flask, request, Response, send_file

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ─────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────
QR_ROTATE_SECONDS  = 60
EXCEL_FILENAME     = "attendance.xlsx"
SUBJECT            = "DTL"
LATE_AFTER_MINUTES = 10
CLASS_START_TIME   = None
ADMIN_PASSWORD     = "admin123"
ROLL_LIST          = []

app = Flask(__name__)

# ─────────────────────────────────────────────
#  SHARED STATE
# ─────────────────────────────────────────────
state_lock     = threading.Lock()
current_token  = ""
token_expiry   = 0.0
attendance_log = []
class_start_dt = datetime.now()
submitted_ips  = set()


def generate_token(length=12):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))


def rotate_token():
    global current_token, token_expiry
    while True:
        new_token = generate_token()
        with state_lock:
            current_token = new_token
            token_expiry  = time.time() + QR_ROTATE_SECONDS
        time.sleep(QR_ROTATE_SECONDS)


def is_late():
    if class_start_dt is None:
        return False
    return (datetime.now() - class_start_dt).total_seconds() / 60 > LATE_AFTER_MINUTES


def get_client_ip():
    """Get real client IP even behind a proxy (Render/Railway use X-Forwarded-For)."""
    if request.headers.get("X-Forwarded-For"):
        return request.headers.get("X-Forwarded-For").split(",")[0].strip()
    return request.remote_addr


def make_qr_base64(url):
    """Generate a QR code and return it as a base64 PNG data URI."""
    qr = qrcode.QRCode(version=2, error_correction=qrcode.constants.ERROR_CORRECT_H,
                       box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


# ══════════════════════════════════════════════
#  ROUTE: "/"  -->  Teacher's live QR display page
#  Opening this URL goes DIRECTLY to the QR page.
# ══════════════════════════════════════════════
@app.route("/")
def teacher_qr_page():
    base_url = request.url_root.rstrip("/")
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{SUBJECT} - Attendance QR</title>
<link rel="manifest" href="/manifest.json">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Attendance">
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{
    font-family:-apple-system,'Segoe UI',sans-serif;
    background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
    min-height:100vh; display:flex; flex-direction:column;
    align-items:center; justify-content:center; padding:20px; color:white;
  }}
  h1 {{ font-size:18px; font-weight:700; margin-bottom:4px; text-align:center; }}
  .sub {{ font-size:13px; color:#94a3b8; margin-bottom:20px; text-align:center; }}
  .qr-box {{
    background:white; border-radius:20px; padding:16px;
    box-shadow:0 20px 60px rgba(0,0,0,0.5); margin-bottom:20px;
  }}
  .qr-box img {{ width:260px; height:260px; display:block; }}
  .timer {{ font-size:40px; font-weight:800; color:#38bdf8; margin-bottom:6px; }}
  .bar {{ width:260px; height:6px; background:#1e293b; border-radius:3px; overflow:hidden; margin-bottom:24px; }}
  .bar-fill {{ height:100%; background:#38bdf8; transition:width 0.25s linear; }}
  .stats {{ display:flex; gap:24px; margin-bottom:20px; }}
  .stat {{ text-align:center; }}
  .stat .n {{ font-size:28px; font-weight:800; }}
  .stat .l {{ font-size:11px; color:#94a3b8; text-transform:uppercase; }}
  .present {{ color:#4ade80; }} .late {{ color:#fbbf24; }} .device {{ color:#a78bfa; }}
  .footer-btn {{
    margin-top:10px; padding:10px 20px; background:rgba(255,255,255,0.1);
    border:1px solid rgba(255,255,255,0.2); border-radius:10px; color:white;
    text-decoration:none; font-size:13px; font-weight:600;
  }}
</style>
</head>
<body>
  <h1>&#x1F4CB; {SUBJECT}</h1>
  <div class="sub" id="dateline"></div>
  <div class="qr-box"><img id="qrimg" src="" alt="QR Code"></div>
  <div class="timer" id="timer">--</div>
  <div class="bar"><div class="bar-fill" id="barfill" style="width:100%"></div></div>
  <div class="stats">
    <div class="stat"><div class="n present" id="cnt-present">0</div><div class="l">Present</div></div>
    <div class="stat"><div class="n late" id="cnt-late">0</div><div class="l">Late</div></div>
    <div class="stat"><div class="n device" id="cnt-device">0</div><div class="l">Devices</div></div>
  </div>
  <a class="footer-btn" href="/admin">Admin Panel</a>

<script>
const ROTATE = {QR_ROTATE_SECONDS};
document.getElementById('dateline').textContent = new Date().toLocaleDateString('en-IN', {{weekday:'long', day:'numeric', month:'long', year:'numeric'}});

let serverRemaining = ROTATE;
let lastSync = Date.now();
let lastTokenQR = "";
let refreshing = false;

async function refresh() {{
  if (refreshing) return;   // prevent overlapping calls from piling up
  refreshing = true;
  try {{
    const res = await fetch('/api/state', {{cache: 'no-store'}});
    if (!res.ok) throw new Error('Bad response: ' + res.status);
    const data = await res.json();
    if (data.qr !== lastTokenQR) {{
      document.getElementById('qrimg').src = data.qr;
      lastTokenQR = data.qr;
    }}
    document.getElementById('cnt-present').textContent = data.present;
    document.getElementById('cnt-late').textContent = data.late;
    document.getElementById('cnt-device').textContent = data.devices;
    serverRemaining = data.remaining;
    lastSync = Date.now();
  }} catch(e) {{
    console.error('refresh failed, will retry:', e);
    // Don't freeze - force a retry shortly instead of leaving old state stuck
    setTimeout(refresh, 1000);
  }} finally {{
    refreshing = false;
  }}
}}

function tick() {{
  const elapsed = (Date.now() - lastSync) / 1000;
  let remaining = serverRemaining - elapsed;
  if (remaining <= 0.2 && !refreshing) {{
    refresh();
  }}
  document.getElementById('timer').textContent = Math.max(0, Math.ceil(remaining)) + 's';
  document.getElementById('barfill').style.width = (Math.max(0, Math.min(100, remaining / ROTATE * 100))) + '%';
}}

refresh();
setInterval(tick, 250);
// Hard safety net: force a refresh every 4s no matter what, so a stuck
// state can never persist for more than a few seconds.
setInterval(() => {{ if (!refreshing) refresh(); }}, 4000);

// re-sync immediately when the tab/screen becomes visible again
// (fixes the "QR expired" issue caused by phone screen sleep / background throttling)
document.addEventListener('visibilitychange', () => {{
  if (!document.hidden) refresh();
}});
</script>
</body>
</html>"""
    return html


@app.route("/manifest.json")
def manifest():
    return Response('''{
  "name": "Attendance QR",
  "short_name": "Attendance",
  "start_url": "/",
  "display": "standalone",
  "background_color": "#1a1a2e",
  "theme_color": "#0f3460"
}''', mimetype="application/json")


# ══════════════════════════════════════════════
#  API: live state for the teacher page (polled by JS)
# ══════════════════════════════════════════════
@app.route("/api/state")
def api_state():
    base_url = request.url_root.rstrip("/")
    with state_lock:
        token   = current_token
        expiry  = token_expiry
        today = datetime.now().strftime("%Y-%m-%d")
        present = sum(1 for r in attendance_log if r["Date"] == today and r["Status"] == "Present")
        late    = sum(1 for r in attendance_log if r["Date"] == today and r["Status"] == "Late")
        devices = len(submitted_ips)
    student_url = f"{base_url}/scan?token={token}"
    qr_data = make_qr_base64(student_url)
    remaining = max(0, round(expiry - time.time(), 1))
    return {"qr": qr_data, "present": present, "late": late, "devices": devices,
            "remaining": remaining, "rotate_seconds": QR_ROTATE_SECONDS}


# ══════════════════════════════════════════════
#  ROUTE: /scan  -->  Student attendance form
# ══════════════════════════════════════════════
@app.route("/scan")
def student_form():
    token_param = request.args.get("token", "")
    with state_lock:
        valid  = (token_param == current_token and time.time() < token_expiry)
        expiry = token_expiry

    if not valid:
        return EXPIRED_HTML, 403

    remaining = max(0, int(expiry - time.time()))
    late_banner = '<div class="late-badge">You are marking attendance late!</div>' if is_late() else ""
    return FORM_HTML.format(
        token=token_param,
        date=datetime.now().strftime("%d %B %Y"),
        subject=SUBJECT,
        interval=remaining,
        late_banner=late_banner
    )


@app.route("/submit", methods=["POST"])
def submit():
    token_param = request.args.get("token", "")
    with state_lock:
        valid = (token_param == current_token and time.time() < token_expiry)
    if not valid:
        return EXPIRED_HTML, 403

    client_ip = get_client_ip()
    with state_lock:
        if client_ip in submitted_ips:
            return DEVICE_BLOCKED_HTML.format(subject=SUBJECT)

    usn     = request.form.get("usn", "").strip().upper()
    name    = request.form.get("name", "").strip().title()
    branch  = request.form.get("branch", "").strip()
    section = request.form.get("section", "").strip()
    today   = datetime.now().strftime("%Y-%m-%d")

    if not usn or not name or not branch or not section:
        return "<h2 style='font-family:sans-serif;padding:20px'>Please fill all fields.</h2>", 400

    with state_lock:
        dup = next((r for r in attendance_log if r["USN"] == usn and r["Date"] == today), None)
        if dup:
            return DUPLICATE_HTML.format(usn=usn, time=dup["Time"])

        status = "Late" if is_late() else "Present"
        record = {
            "USN": usn, "Name": name, "Branch": branch, "Section": section,
            "Date": today, "Time": datetime.now().strftime("%H:%M:%S"),
            "Subject": SUBJECT, "Status": status
        }
        attendance_log.append(record)
        submitted_ips.add(client_ip)

    save_to_excel(record)

    return SUCCESS_HTML.format(
        icon="&#x2705;" if status == "Present" else "&#x23F0;",
        status_class="ok" if status == "Present" else "late",
        badge_class="badge-present" if status == "Present" else "badge-late",
        status=status, usn=usn, name=name, branch=branch, section=section,
        date=datetime.now().strftime("%d %B %Y"), subject=SUBJECT
    )


# ══════════════════════════════════════════════
#  ADMIN PANEL
# ══════════════════════════════════════════════
@app.route("/admin")
def admin_login_or_panel():
    pwd = request.args.get("pwd", "")
    if pwd != ADMIN_PASSWORD:
        return ADMIN_LOGIN_HTML
    return build_admin_html()


@app.route("/admin/delete")
def admin_delete():
    pwd = request.args.get("pwd", "")
    usn = request.args.get("usn", "")
    if pwd != ADMIN_PASSWORD:
        return "Unauthorized", 401
    today = datetime.now().strftime("%Y-%m-%d")
    with state_lock:
        before = len(attendance_log)
        attendance_log[:] = [r for r in attendance_log if not (r["USN"] == usn and r["Date"] == today)]
        removed = before - len(attendance_log)
    if removed:
        rebuild_excel_for_today()
    return f"<meta http-equiv='refresh' content='0;url=/admin?pwd={pwd}'>"


@app.route("/admin/reset_devices")
def admin_reset_devices():
    pwd = request.args.get("pwd", "")
    if pwd != ADMIN_PASSWORD:
        return "Unauthorized", 401
    with state_lock:
        submitted_ips.clear()
    return f"<meta http-equiv='refresh' content='0;url=/admin?pwd={pwd}'>"


@app.route("/admin/absent")
def admin_absent():
    pwd = request.args.get("pwd", "")
    if pwd != ADMIN_PASSWORD:
        return "Unauthorized", 401
    if not ROLL_LIST:
        return f"<meta http-equiv='refresh' content='2;url=/admin?pwd={pwd}'><p style='font-family:sans-serif;padding:20px;color:white;background:#0f172a'>Roll list is empty. Add USNs in Settings first.</p>"
    count = generate_absent_list()
    msg = f"{count} student(s) marked Absent." if count else "All students already marked attendance."
    return f"<meta http-equiv='refresh' content='2;url=/admin?pwd={pwd}'><p style='font-family:sans-serif;padding:20px;color:white;background:#0f172a'>{msg}</p>"


@app.route("/admin/monthly")
def admin_monthly():
    pwd = request.args.get("pwd", "")
    if pwd != ADMIN_PASSWORD:
        return "Unauthorized", 401
    ok = generate_monthly_summary()
    msg = "Monthly Summary created! Download Excel to view." if ok else "No attendance data for this month yet."
    return f"<meta http-equiv='refresh' content='2;url=/admin?pwd={pwd}'><p style='font-family:sans-serif;padding:20px;color:white;background:#0f172a'>{msg}</p>"


@app.route("/admin/settings", methods=["GET", "POST"])
def admin_settings():
    global SUBJECT, QR_ROTATE_SECONDS, LATE_AFTER_MINUTES, ADMIN_PASSWORD, CLASS_START_TIME, ROLL_LIST, class_start_dt
    pwd = request.args.get("pwd", "")
    if pwd != ADMIN_PASSWORD:
        return "Unauthorized", 401

    if request.method == "POST":
        try:
            new_subject  = request.form.get("subject", "").strip()
            new_rotate   = int(request.form.get("rotate", "30"))
            new_late     = int(request.form.get("late", "10"))
            new_password = request.form.get("password", "").strip()
            new_start    = request.form.get("start", "").strip()
            new_rolls    = [u.strip().upper() for u in request.form.get("roll_list", "").splitlines() if u.strip()]
            if not new_subject or not new_password:
                raise ValueError("Subject and password required")
            SUBJECT = new_subject; QR_ROTATE_SECONDS = new_rotate; LATE_AFTER_MINUTES = new_late
            ADMIN_PASSWORD = new_password; CLASS_START_TIME = new_start if new_start else None
            ROLL_LIST = new_rolls
            if CLASS_START_TIME:
                h, m = map(int, CLASS_START_TIME.split(":"))
                class_start_dt = datetime.now().replace(hour=h, minute=m, second=0, microsecond=0)
            else:
                class_start_dt = datetime.now()
            return f"<meta http-equiv='refresh' content='1;url=/admin?pwd={ADMIN_PASSWORD}&tab=settings'>"
        except Exception as e:
            return f"<p style='font-family:sans-serif;padding:20px;color:red'>Error: {e}</p>"

    return build_admin_html()


@app.route("/download/excel")
def download_excel():
    pwd = request.args.get("pwd", "")
    if pwd != ADMIN_PASSWORD:
        return "Unauthorized", 401
    if not os.path.exists(EXCEL_FILENAME):
        return "No attendance file yet.", 404
    return send_file(EXCEL_FILENAME, as_attachment=True, download_name="attendance.xlsx")


def build_admin_html():
    today = datetime.now().strftime("%Y-%m-%d")
    with state_lock:
        records      = [r for r in attendance_log if r["Date"] == today]
        device_count = len(submitted_ips)
    rows = ""
    for i, r in enumerate(records, 1):
        sc = "late" if r["Status"] == "Late" else ("absent" if r["Status"] == "Absent" else "present")
        del_btn = (f"<button class='del-btn' onclick=\"location.href='/admin/delete?pwd={ADMIN_PASSWORD}&usn={r['USN']}'\">Delete</button>"
                   if r["Status"] != "Absent" else "—")
        rows += (f"<tr><td>{i}</td><td>{r['USN']}</td><td>{r['Name']}</td>"
                 f"<td>{r['Branch']}</td><td>{r['Section']}</td><td>{r['Time']}</td>"
                 f"<td class='{sc}'>{r['Status']}</td><td>{del_btn}</td></tr>")
    present = sum(1 for r in records if r["Status"] == "Present")
    late    = sum(1 for r in records if r["Status"] == "Late")
    absent  = sum(1 for r in records if r["Status"] == "Absent")

    return ADMIN_HTML.format(
        subject=SUBJECT, date=datetime.now().strftime("%d %B %Y"), pwd=ADMIN_PASSWORD,
        rows=rows if rows else "<tr><td colspan='8' style='text-align:center;padding:20px;color:#64748b'>No records yet</td></tr>",
        total=len(records), present=present, late=late, absent=absent, device_count=device_count,
        subject_val=SUBJECT, rotate_val=QR_ROTATE_SECONDS, late_val=LATE_AFTER_MINUTES,
        password_val=ADMIN_PASSWORD, start_val=CLASS_START_TIME or "", roll_val="\n".join(ROLL_LIST)
    )


# ══════════════════════════════════════════════
#  HTML TEMPLATES (student-facing + admin)
# ══════════════════════════════════════════════
FORM_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Mark Attendance</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{
    font-family:-apple-system,'Segoe UI',sans-serif;
    background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
    min-height:100vh; display:flex; align-items:center;
    justify-content:center; padding:20px;
  }}
  .card {{ background:white; border-radius:16px; padding:36px 32px; width:100%; max-width:420px; box-shadow:0 20px 60px rgba(0,0,0,0.4); }}
  .logo {{ text-align:center; margin-bottom:20px; }}
  .logo h1 {{ font-size:22px; color:#1a1a2e; font-weight:700; }}
  .logo p {{ font-size:13px; color:#888; margin-top:4px; }}
  .late-badge {{ background:#fef3c7; color:#92400e; border:1px solid #fde68a; border-radius:8px; padding:8px 12px; font-size:13px; text-align:center; margin-bottom:16px; font-weight:600; }}
  label {{ display:block; font-size:13px; font-weight:600; color:#444; margin-bottom:6px; }}
  input, select {{ width:100%; padding:11px 14px; border:1.5px solid #ddd; border-radius:8px; font-size:15px; margin-bottom:16px; background:white; }}
  input:focus, select:focus {{ outline:none; border-color:#0f3460; }}
  .row {{ display:flex; gap:12px; }} .row > div {{ flex:1; }}
  button {{ width:100%; padding:13px; background:#0f3460; color:white; border:none; border-radius:8px; font-size:16px; font-weight:600; cursor:pointer; }}
  .timer-bar {{ height:4px; background:#e5e7eb; border-radius:2px; overflow:hidden; margin-bottom:20px; }}
  .timer-fill {{ height:100%; background:#0f3460; animation:drain {interval}s linear forwards; }}
  @keyframes drain {{ from{{width:100%}} to{{width:0%}} }}
</style>
</head>
<body>
<div class="card">
  <div class="logo"><h1>&#x1F4CB; Mark Attendance</h1><p>{date} &nbsp;|&nbsp; {subject}</p></div>
  <div class="timer-bar"><div class="timer-fill"></div></div>
  {late_banner}
  <form method="POST" action="/submit?token={token}">
    <label>USN / Roll Number</label>
    <input name="usn" type="text" placeholder="e.g. 1BM21CS001" required autofocus>
    <label>Full Name</label>
    <input name="name" type="text" placeholder="e.g. Rahul Sharma" required>
    <div class="row">
      <div><label>Branch</label>
        <select name="branch" required>
          <option value="" disabled selected>Select</option>
          <option value="CSE">CSE</option><option value="ETE">ETE</option>
          <option value="ISE">ISE</option><option value="ECE">ECE</option>
          <option value="EEE">EEE</option><option value="ME">ME</option>
          <option value="CV">CV</option><option value="AI&ML">AI&amp;ML</option>
          <option value="Other">Other</option>
        </select>
      </div>
      <div><label>Section</label>
        <select name="section" required>
          <option value="" disabled selected>Select</option>
          <option value="No Section">No Section</option>
          <option value="A">A</option><option value="B">B</option>
          <option value="C">C</option><option value="D">D</option><option value="E">E</option>
        </select>
      </div>
    </div>
    <button type="submit">Submit Attendance &#x2713;</button>
  </form>
</div>
</body>
</html>"""

SUCCESS_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Attendance Marked</title>
<style>
  body {{ font-family:-apple-system,sans-serif; background:linear-gradient(135deg,#1a1a2e,#0f3460); display:flex;align-items:center;justify-content:center;min-height:100vh; }}
  .card {{ background:white;border-radius:16px;padding:40px 32px;text-align:center;max-width:400px; }}
  .icon {{ font-size:60px; margin-bottom:12px; }} h2 {{ font-size:22px; }}
  .ok {{ color:#16a34a; }} .late {{ color:#d97706; }}
  p {{ color:#555; margin-top:8px; font-size:14px; }}
  .badge {{ display:inline-block;margin-top:12px;padding:4px 12px;border-radius:20px;font-size:13px;font-weight:600; }}
  .badge-present {{ background:#d1fae5; color:#065f46; }} .badge-late {{ background:#fef3c7; color:#92400e; }}
</style></head>
<body><div class="card">
  <div class="icon">{icon}</div>
  <h2 class="{status_class}">Attendance Marked!</h2>
  <p><strong>{usn}</strong> &mdash; {name}</p>
  <p>{branch} &nbsp;|&nbsp; Section: {section}</p>
  <p style="margin-top:8px;color:#888">{date} &nbsp; {subject}</p>
  <span class="badge {badge_class}">{status}</span>
</div></body></html>"""

EXPIRED_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>QR Expired</title>
<style>
  body {{ font-family:-apple-system,sans-serif;background:linear-gradient(135deg,#7f1d1d,#991b1b);display:flex;align-items:center;justify-content:center;min-height:100vh; }}
  .card {{ background:white;border-radius:16px;padding:40px;text-align:center;max-width:360px; }}
  .icon {{ font-size:60px; }} h2 {{ color:#dc2626;margin-top:12px; }} p {{ color:#555;margin-top:8px;font-size:14px; }}
</style></head><body><div class="card">
  <div class="icon">&#x23F1;&#xFE0F;</div><h2>QR Code Expired</h2>
  <p>Please scan the latest QR code shown on the screen.</p>
</div></body></html>"""

DUPLICATE_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Already Marked</title>
<style>
  body {{ font-family:-apple-system,sans-serif;background:linear-gradient(135deg,#1e3a5f,#0f3460);display:flex;align-items:center;justify-content:center;min-height:100vh; }}
  .card {{ background:white;border-radius:16px;padding:40px;text-align:center;max-width:360px; }}
  .icon {{ font-size:60px; }} h2 {{ color:#d97706;margin-top:12px; }} p {{ color:#555;margin-top:8px;font-size:14px; }}
</style></head><body><div class="card">
  <div class="icon">&#x26A0;&#xFE0F;</div><h2>Already Marked</h2>
  <p><strong>{usn}</strong> already marked at <strong>{time}</strong> today.</p>
</div></body></html>"""

DEVICE_BLOCKED_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Device Already Used</title>
<style>
  body {{ font-family:-apple-system,sans-serif;background:linear-gradient(135deg,#312e81,#4338ca);display:flex;align-items:center;justify-content:center;min-height:100vh; }}
  .card {{ background:white;border-radius:16px;padding:40px;text-align:center;max-width:380px; }}
  .icon {{ font-size:60px; }} h2 {{ color:#4338ca;margin-top:12px; }} p {{ color:#555;margin-top:8px;font-size:14px;line-height:1.6; }}
  .note {{ margin-top:16px;background:#ede9fe;color:#4338ca;padding:10px 14px;border-radius:8px;font-size:13px;font-weight:600; }}
</style></head><body><div class="card">
  <div class="icon">&#x1F4F5;</div><h2>Device Already Used</h2>
  <p>Attendance has already been submitted from this device for <strong>{subject}</strong> today.</p>
  <div class="note">Each device can only submit attendance once per session.</div>
</div></body></html>"""

ADMIN_LOGIN_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Admin Login</title>
<style>
  body {{ font-family:-apple-system,sans-serif;background:#0f172a;display:flex;align-items:center;justify-content:center;min-height:100vh;color:white; }}
  .card {{ background:#1e293b;padding:32px;border-radius:16px;width:300px; }}
  h2 {{ margin-bottom:16px;color:#38bdf8; }}
  input {{ width:100%;padding:11px;border-radius:8px;border:1px solid #334155;background:#0f172a;color:white;margin-bottom:12px; }}
  button {{ width:100%;padding:11px;background:#0f3460;color:white;border:none;border-radius:8px;font-weight:600;cursor:pointer; }}
</style></head><body>
<div class="card">
  <h2>Admin Login</h2>
  <form method="GET" action="/admin">
    <input type="password" name="pwd" placeholder="Enter password" required autofocus>
    <button type="submit">Login</button>
  </form>
</div></body></html>"""

ADMIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Admin Panel</title>
<style>
  * {{ box-sizing:border-box;margin:0;padding:0; }}
  body {{ font-family:-apple-system,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh; }}
  .topbar {{ background:#0f3460;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px; }}
  .topbar h1 {{ font-size:18px;color:#38bdf8; }}
  .topbar span {{ font-size:13px;color:#94a3b8; }}
  .tabs {{ display:flex;background:#1e293b;padding:0 24px;overflow-x:auto; }}
  .tab {{ padding:12px 20px;cursor:pointer;font-size:13px;font-weight:600;color:#94a3b8;border-bottom:3px solid transparent;white-space:nowrap; }}
  .tab.active {{ color:#38bdf8;border-bottom-color:#38bdf8; }}
  .section {{ display:none;padding:24px; }} .section.active {{ display:block; }}
  table {{ width:100%;border-collapse:collapse;font-size:13px; }}
  th {{ background:#0f3460;padding:10px 12px;text-align:left;color:white; }}
  td {{ padding:9px 12px;border-bottom:1px solid #1e293b; }}
  .present {{ color:#4ade80;font-weight:600; }} .late {{ color:#fbbf24;font-weight:600; }} .absent {{ color:#f87171;font-weight:600; }}
  .del-btn {{ background:#7f1d1d;color:white;border:none;padding:5px 12px;border-radius:6px;cursor:pointer;font-size:12px; }}
  .summary {{ margin-top:16px;font-size:14px;color:#94a3b8;background:#1e293b;padding:12px 16px;border-radius:8px; }}
  .sp {{ color:#4ade80; }} .sl {{ color:#fbbf24; }} .sa {{ color:#f87171; }}
  .device-count {{ font-size:13px;color:#94a3b8;margin-top:8px; }} .device-count span {{ color:#a78bfa;font-weight:700; }}
  .settings-grid {{ display:grid;grid-template-columns:1fr 1fr;gap:20px;max-width:700px; }}
  .field label {{ display:block;font-size:12px;font-weight:600;color:#94a3b8;margin-bottom:6px; }}
  .field input, .field textarea {{ width:100%;padding:10px 12px;background:#1e293b;border:1.5px solid #334155;border-radius:8px;color:#e2e8f0;font-size:14px; }}
  .field textarea {{ min-height:120px;font-family:monospace;font-size:12px; }}
  .save-btn, .reset-btn, .card-btn {{ display:inline-block;margin-top:16px;padding:10px 20px;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;text-decoration:none;color:white; }}
  .save-btn {{ background:#16a34a; }} .reset-btn {{ background:#312e81; }}
  .action-cards {{ display:flex;gap:20px;flex-wrap:wrap; }}
  .card {{ background:#1e293b;border-radius:12px;padding:24px;width:260px;border:1px solid #334155; }}
  .card h3 {{ font-size:15px;margin-bottom:8px; }} .card p {{ font-size:13px;color:#94a3b8;margin-bottom:8px; }}
  .btn-amber {{ background:#b45309; }} .btn-teal {{ background:#0e7490; }}
  h2 {{ font-size:16px;color:#94a3b8;margin-bottom:16px; }}
</style></head>
<body>
<div class="topbar"><h1>Admin Panel</h1><span>{subject} | {date}</span></div>
<div class="tabs">
  <div class="tab active" onclick="show('att',this)">Attendance</div>
  <div class="tab" onclick="show('set',this)">Settings</div>
  <div class="tab" onclick="show('rep',this)">Reports</div>
</div>
<div id="att" class="section active">
  <h2>Today's Attendance</h2>
  <table><tr><th>#</th><th>USN</th><th>Name</th><th>Branch</th><th>Sec</th><th>Time</th><th>Status</th><th>Action</th></tr>{rows}</table>
  <div class="summary">Total: <span class="sp">{total}</span> | Present: <span class="sp">{present}</span> | Late: <span class="sl">{late}</span> | Absent: <span class="sa">{absent}</span></div>
  <div class="device-count">Devices locked: <span>{device_count}</span></div>
  <a href="/admin/reset_devices?pwd={pwd}" class="reset-btn" onclick="return confirm('Reset device locks?')">Reset Device Locks</a>
  <a href="/download/excel?pwd={pwd}" class="reset-btn" style="background:#0f3460">Download Excel</a>
</div>
<div id="set" class="section">
  <h2>Settings</h2>
  <form method="POST" action="/admin/settings?pwd={pwd}">
    <div class="settings-grid">
      <div class="field"><label>Subject</label><input name="subject" value="{subject_val}" required></div>
      <div class="field"><label>QR Rotate (s)</label><input name="rotate" type="number" value="{rotate_val}" required></div>
      <div class="field"><label>Late After (min)</label><input name="late" type="number" value="{late_val}" required></div>
      <div class="field"><label>Admin Password</label><input name="password" type="password" value="{password_val}" required></div>
      <div class="field"><label>Class Start (HH:MM)</label><input name="start" value="{start_val}" placeholder="09:30"></div>
    </div>
    <div class="field" style="margin-top:16px;max-width:700px">
      <label>Roll List (one USN per line)</label>
      <textarea name="roll_list">{roll_val}</textarea>
    </div>
    <button type="submit" class="save-btn">Save Settings</button>
  </form>
</div>
<div id="rep" class="section">
  <h2>Reports</h2>
  <div class="action-cards">
    <div class="card"><h3>Absent List</h3><p>Mark missing students Absent in Excel.</p>
      <a href="/admin/absent?pwd={pwd}" class="card-btn btn-amber" onclick="return confirm('Generate absent list?')">Generate</a></div>
    <div class="card"><h3>Monthly Summary</h3><p>Create attendance % report for the month.</p>
      <a href="/admin/monthly?pwd={pwd}" class="card-btn btn-teal">Generate</a></div>
  </div>
</div>
<script>
function show(id, el) {{
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  el.classList.add('active');
}}
const p = new URLSearchParams(location.search);
if (p.get('tab') === 'settings') show('set', document.querySelectorAll('.tab')[1]);
</script>
</body></html>"""


# ══════════════════════════════════════════════
#  EXCEL FUNCTIONS (same as before)
# ══════════════════════════════════════════════
HEADERS      = ["USN", "Name", "Branch", "Section", "Date", "Time", "Subject", "Status"]
COL_WIDTHS   = [16, 22, 10, 12, 14, 12, 14, 12]
HEADER_FILL  = PatternFill("solid", start_color="0F3460")
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
PRESENT_FILL = PatternFill("solid", start_color="D1FAE5")
LATE_FILL    = PatternFill("solid", start_color="FEF3C7")
ABSENT_FILL  = PatternFill("solid", start_color="FEE2E2")
CELL_FONT    = Font(name="Arial", size=10)
THIN         = Side(style="thin", color="CCCCCC")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER       = Alignment(horizontal="center", vertical="center")


def _write_header(ws):
    ws.append(HEADERS)
    for col_idx, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = h; cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = CENTER; cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 22


def _load_or_create_wb(path):
    if os.path.exists(path):
        try:
            return openpyxl.load_workbook(path)
        except Exception:
            pass
    return openpyxl.Workbook()


def _style_row(ws, row_num):
    for col_idx in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.font = CELL_FONT; cell.border = BORDER; cell.alignment = CENTER
        val = cell.value
        if val == "Present":  cell.fill = PRESENT_FILL
        elif val == "Late":   cell.fill = LATE_FILL
        elif val == "Absent": cell.fill = ABSENT_FILL


def _write_summary(ws):
    for row in range(ws.max_row, 1, -1):
        v = ws.cell(row, 1).value
        if v in ("Total Present", "Total Late", "Total Absent", None):
            ws.delete_rows(row)
        else:
            break
    present = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r and r[7] == "Present")
    late    = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r and r[7] == "Late")
    absent  = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r and r[7] == "Absent")
    sr = ws.max_row + 2
    ws.cell(sr,   1, "Total Present").font = Font(bold=True, name="Arial")
    ws.cell(sr,   2, present).font         = Font(bold=True, name="Arial", color="16A34A")
    ws.cell(sr+1, 1, "Total Late").font    = Font(bold=True, name="Arial")
    ws.cell(sr+1, 2, late).font            = Font(bold=True, name="Arial", color="D97706")
    ws.cell(sr+2, 1, "Total Absent").font  = Font(bold=True, name="Arial")
    ws.cell(sr+2, 2, absent).font          = Font(bold=True, name="Arial", color="DC2626")


def _safe_save(wb, path, tmp_path):
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, path)
    except PermissionError:
        ts = datetime.now().strftime("%H%M%S")
        wb.save(f"attendance_backup_{ts}.xlsx")
    finally:
        if os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except: pass


def save_to_excel(record):
    path = EXCEL_FILENAME; tmp_path = EXCEL_FILENAME + ".tmp"
    wb   = _load_or_create_wb(path)
    date_sheet = record["Date"]
    if date_sheet not in wb.sheetnames:
        ws = wb.create_sheet(date_sheet)
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 2:
            del wb["Sheet"]
        _write_header(ws)
    else:
        ws = wb[date_sheet]
    ws.append([record[h] for h in HEADERS])
    _style_row(ws, ws.max_row)
    _write_summary(ws)
    _safe_save(wb, path, tmp_path)


def rebuild_excel_for_today():
    today = datetime.now().strftime("%Y-%m-%d")
    path  = EXCEL_FILENAME; tmp_path = EXCEL_FILENAME + ".tmp"
    wb    = _load_or_create_wb(path)
    if today in wb.sheetnames: del wb[today]
    ws = wb.create_sheet(today)
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    _write_header(ws)
    with state_lock:
        records = [r for r in attendance_log if r["Date"] == today]
    for record in records:
        ws.append([record[h] for h in HEADERS])
        _style_row(ws, ws.max_row)
    _write_summary(ws)
    _safe_save(wb, path, tmp_path)


def generate_absent_list():
    today = datetime.now().strftime("%Y-%m-%d")
    if not ROLL_LIST:
        return 0
    with state_lock:
        present_usns = {r["USN"] for r in attendance_log if r["Date"] == today}
    absent_usns = [u.strip().upper() for u in ROLL_LIST if u.strip().upper() not in present_usns]
    if not absent_usns:
        return 0
    path = EXCEL_FILENAME; tmp_path = EXCEL_FILENAME + ".tmp"
    wb   = _load_or_create_wb(path)
    if today not in wb.sheetnames:
        ws = wb.create_sheet(today)
        if "Sheet" in wb.sheetnames: del wb["Sheet"]
        _write_header(ws)
    else:
        ws = wb[today]
    for usn in absent_usns:
        record = {"USN": usn, "Name": "—", "Branch": "—", "Section": "—",
                  "Date": today, "Time": "—", "Subject": SUBJECT, "Status": "Absent"}
        ws.append([record[h] for h in HEADERS])
        _style_row(ws, ws.max_row)
        with state_lock:
            attendance_log.append(record)
    _write_summary(ws)
    _safe_save(wb, path, tmp_path)
    return len(absent_usns)


def generate_monthly_summary():
    path = EXCEL_FILENAME; tmp_path = EXCEL_FILENAME + ".tmp"
    if not os.path.exists(path):
        return False
    wb   = _load_or_create_wb(path)
    now  = datetime.now()
    month_prefix = f"{now.year}-{now.month:02d}-"
    date_sheets  = [s for s in wb.sheetnames if s.startswith(month_prefix)]
    if not date_sheets:
        return False
    all_usns = {}; daily = {}
    for sheet_name in date_sheets:
        ws = wb[sheet_name]; day = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            usn = str(row[0]); status = str(row[7]) if row[7] else "—"
            if status in ("Total Present", "Total Late", "Total Absent"): continue
            all_usns[usn] = {"Name": str(row[1] or "—"), "Branch": str(row[2] or "—"), "Section": str(row[3] or "—")}
            day[usn] = status
        daily[sheet_name] = day
    if "Monthly_Summary" in wb.sheetnames:
        del wb["Monthly_Summary"]
    ms = wb.create_sheet("Monthly_Summary", 0)
    active_dates = sorted(daily.keys())
    headers = ["USN", "Name", "Branch", "Section"] + active_dates + ["Present", "Late", "Absent", "Total Days", "Attendance %"]
    ms.append(headers)
    hdr_fill = PatternFill("solid", start_color="0F3460"); hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    for col_idx in range(1, len(headers) + 1):
        cell = ms.cell(row=1, column=col_idx)
        cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = CENTER; cell.border = BORDER
        ms.column_dimensions[cell.column_letter].width = 14 if col_idx > 4 else 18
    pct_red = PatternFill("solid", start_color="FEE2E2"); pct_yellow = PatternFill("solid", start_color="FEF3C7"); pct_green = PatternFill("solid", start_color="D1FAE5")
    for usn, info in sorted(all_usns.items()):
        row_data = [usn, info["Name"], info["Branch"], info["Section"]]
        p = l = a = 0
        for d in active_dates:
            st = daily[d].get(usn, "Absent"); row_data.append(st)
            if st == "Present": p += 1
            elif st == "Late": l += 1
            else: a += 1
        total = len(active_dates)
        pct = round((p + l) / total * 100, 1) if total else 0
        row_data += [p, l, a, total, f"{pct}%"]
        ms.append(row_data); rn = ms.max_row
        for ci in range(1, len(row_data) + 1):
            cell = ms.cell(row=rn, column=ci)
            cell.font = CELL_FONT; cell.border = BORDER; cell.alignment = CENTER
            val = cell.value
            if val == "Present": cell.fill = PRESENT_FILL
            elif val == "Late": cell.fill = LATE_FILL
            elif val == "Absent": cell.fill = ABSENT_FILL
        pct_cell = ms.cell(row=rn, column=len(headers))
        if pct < 75: pct_cell.fill = pct_red
        elif pct < 85: pct_cell.fill = pct_yellow
        else: pct_cell.fill = pct_green
        pct_cell.font = Font(bold=True, name="Arial", size=10)
    ms.freeze_panes = "E2"
    _safe_save(wb, path, tmp_path)
    return True


# ══════════════════════════════════════════════
#  START TOKEN ROTATION IN BACKGROUND
# ══════════════════════════════════════════════
rotator_thread = threading.Thread(target=rotate_token, daemon=True)
rotator_thread.start()
time.sleep(0.1)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
