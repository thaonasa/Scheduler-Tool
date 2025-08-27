import os
import json
import uuid
import datetime as dt
from io import BytesIO

from flask import Flask, request, render_template_string, send_file, redirect, url_for, jsonify
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, Text
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.styles import Font



app = Flask(__name__)

# ========== CẤU HÌNH CHUNG ==========
COMPANY_NAME = "Đồng Tiến Bakery"
DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "meeting_schedule.json")  # Lưu JSON
WEEK_DAYS = 6  # Thứ 2 -> Thứ 7

# Bảng màu Chủ trì
CHAIR_COLORS = {
    'TGĐ': '#fcba03',
    'CEO': '#FF9999',
    'COO': '#99CCFF',
    'GS.XD': '#CCFF99',
    'TPQC': '#FFFF99',
    'PPNSĐT': '#FFCC99',
    'CV.BGĐ': '#CC99FF',
    'ITPM_N.Nguyên': '#99FFFF',
    'NVISO': '#FF99FF',
    'TPKT': '#99FF99',
    'TBHSE': '#FFCCFF',
    'PP.KTTC': '#CCFFFF',
    'GS.IT': '#FFCC00',
    'TL.BGĐ_YP': '#99CC00',
    'PPKV': '#FF9900'
}

# Danh sách Loại & Phòng họp cho select
CATEGORIES = ["Họp định kỳ", "Họp nội bộ", "Đào tạo", "Phỏng vấn"]
ROOMS = ["Phòng họp 1", "Phòng họp 2", "Phòng họp 3", "Phòng Tổng Giám Đốc"]

# ========== TIỆN ÍCH ==========
def ensure_data_file():
    data_dir = os.path.dirname(DATA_PATH)
    os.makedirs(data_dir, exist_ok=True)
    if not os.path.exists(DATA_PATH):
        with open(DATA_PATH, "w", encoding="utf-8") as f:
            json.dump({"sessions": []}, f, ensure_ascii=False, indent=2)

def load_data():
    ensure_data_file()
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data):
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def monday_of_week(any_date: dt.date) -> dt.date:
    return any_date - dt.timedelta(days=any_date.weekday())

def saturday_of_week(any_date: dt.date) -> dt.date:
    return monday_of_week(any_date) + dt.timedelta(days=WEEK_DAYS - 1)

def session_id_from_date(any_date: dt.date) -> str:
    iso_year, iso_week, _ = any_date.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"

def excel_color(hex_color: str) -> str:
    c = hex_color.strip()
    if c.startswith("#"):
        c = c[1:]
    if len(c) == 6:
        return "FF" + c.upper()
    if len(c) == 8:
        return c.upper()
    return "FF000000"

def hhmm_to_minutes(hhmm: str) -> int:
    h, m = map(int, hhmm.split(":"))
    return h * 60 + m

def guess_buoi(start_hhmm: str) -> str:
    return "SÁNG" if hhmm_to_minutes(start_hhmm) < 12 * 60 else "CHIỀU"

def overlap(e1, e2) -> bool:
    if e1["date"] != e2["date"] or e1["session_buoi"] != e2["session_buoi"]:
        return False
    s1, e1t = hhmm_to_minutes(e1["start_time"]), hhmm_to_minutes(e1["end_time"])
    s2, e2t = hhmm_to_minutes(e2["start_time"]), hhmm_to_minutes(e2["end_time"])
    return max(s1, s2) < min(e1t, e2t)

def compute_conflicts(events):
    by_key = {}
    for ev in events:
        key = (ev["date"], ev["session_buoi"])
        by_key.setdefault(key, []).append(ev)

    for key, arr in by_key.items():
        arr.sort(key=lambda x: hhmm_to_minutes(x["start_time"]))
        for i in range(len(arr)):
            arr[i]["conflict"] = False
        for i in range(1, len(arr)):
            if overlap(arr[i - 1], arr[i]):
                arr[i]["conflict"] = True
                arr[i - 1]["conflict"] = True

def get_or_create_session(data, any_date: dt.date):
    sid = session_id_from_date(any_date)
    for s in data["sessions"]:
        if s["id"] == sid:
            return s
    new_session = {
        "id": sid,
        "week_start": monday_of_week(any_date).isoformat(),
        "week_end": saturday_of_week(any_date).isoformat(),
        "events": []
    }
    data["sessions"].append(new_session)
    save_data(data)
    return new_session

def find_session_by_id(data, sid: str):
    for s in data["sessions"]:
        if s["id"] == sid:
            return s
    return None

def upsert_event(session, payload):
    _id = payload.get("id") or str(uuid.uuid4())
    ev = {
        "id": _id,
        "date": payload["date"],
        "session_buoi": payload["buoi"],
        "start_time": payload["start_time"],
        "end_time": payload["end_time"],
        "title": payload["title"],
        "category": payload.get("category", ""),
        "chair": payload["chair"],
        "attendees": payload.get("attendees", ""),
        "location": payload.get("location", "")
    }
    if hhmm_to_minutes(ev["start_time"]) >= hhmm_to_minutes(ev["end_time"]):
        raise ValueError("Giờ kết thúc phải lớn hơn giờ bắt đầu.")
    for i, e in enumerate(session["events"]):
        if e["id"] == _id:
            session["events"][i] = ev
            return ev
    session["events"].append(ev)
    return ev

def delete_event(session, event_id: str):
    session["events"] = [e for e in session["events"] if e["id"] != event_id]

# ======= DỮ LIỆU GỘP THEO NGÀY/BUỔI (dùng cho Export & Preview) =======
def build_schedule(session):
    """Trả về:
    - dates: list[date]
    - schedule: { 'dd/mm/YYYY': { 'SÁNG': [ev...], 'CHIỀU':[ev...] } }
    """
    week_start = dt.date.fromisoformat(session["week_start"])
    dates = [week_start + dt.timedelta(days=i) for i in range(WEEK_DAYS)]
    events = sorted(session["events"], key=lambda e: (e["date"], hhmm_to_minutes(e["start_time"])))
    compute_conflicts(events)
    schedule = {}
    for ev in events:
        key = dt.date.fromisoformat(ev["date"]).strftime('%d/%m/%Y')
        schedule.setdefault(key, {'SÁNG': [], 'CHIỀU': []})
        schedule[key][ev["session_buoi"]].append(ev)
    return dates, schedule

# ========== XUẤT EXCEL DẠNG BẢNG LỊCH HỌP ==========
def export_session_to_excel(session):
    """
    Xuất Excel: bố cục giống preview
    - Mỗi sự kiện chiếm 2 hàng (header + details)
    - Header bôi đậm & căn giữa (highlight điểm chính)
    - Details căn trái
    - Tô màu theo Chủ trì cho cả 2 hàng của sự kiện
    """
    wb = Workbook()
    ws = wb.active

    week_start = dt.date.fromisoformat(session["week_start"])
    week_end   = dt.date.fromisoformat(session["week_end"])

    # build_schedule(session) phải trả về:
    #   dates: [date, ...] của Thứ 2 -> Thứ 7 (datetime.date)
    #   schedule: { 'dd/mm/YYYY': { 'SÁNG': [event...], 'CHIỀU': [event...] } }
    dates, schedule = build_schedule(session)

    weekdays = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7']

    # ----- Tiêu đề -----
    ws.merge_cells('A1:G1')
    ws['A1'] = f"LỊCH HỌP TUẦN {COMPANY_NAME.upper()}"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:G2')
    ws['A2'] = f"Tuần:  {week_start.strftime('%d/%m/%Y')} -> {week_end.strftime('%d/%m/%Y')}"
    ws['A2'].font = Font(bold=True)
    ws['A2'].alignment = Alignment(horizontal='center')

    # ----- Header -----
    ws['A3'] = "Buổi"
    ws['A3'].font = Font(bold=True)
    for i, d in enumerate(dates, start=2):
        c = ws.cell(row=3, column=i, value=f"{weekdays[i-2]}\n({d.strftime('%d.%m.%Y')})")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.font = Font(bold=True)

    thin   = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ----- Dữ liệu -----
    start_row = 4
    date_keys = [d.strftime('%d/%m/%Y') for d in dates]

    for buoi in ["SÁNG", "CHIỀU"]:
        # số sự kiện tối đa trong buổi này (trên các ngày)
        max_events = max([len(schedule.get(k, {}).get(buoi, [])) for k in date_keys] + [0])
        if max_events < 1:
            max_events = 1  # vẫn chừa block để hiện "—"

        # mỗi sự kiện 2 hàng
        block_height = max_events * 2
        end_row = start_row + block_height - 1

        # Cột A: gộp ô "Buổi"
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        a = ws.cell(row=start_row, column=1, value=buoi)
        a.font = Font(bold=True)
        a.alignment = Alignment(horizontal="center", vertical="center")
        a.border = border

        # Điền dữ liệu theo từng ngày
        for col_idx, k in enumerate(date_keys, start=2):
            evs = list(schedule.get(k, {}).get(buoi, []))
            evs.sort(key=lambda e: hhmm_to_minutes(e["start_time"]))

            for r_off in range(max_events):
                # Hàng header & details cho sự kiện thứ r_off
                header_row  = start_row + r_off * 2
                detail_row  = header_row + 1

                # Bắt buộc set viền + alignment mặc định
                for rr in (header_row, detail_row):
                    cc = ws.cell(row=rr, column=col_idx)
                    cc.border = border

                if r_off < len(evs):
                    ev = evs[r_off]

                    # ----- HEADER (in đậm & căn giữa) -----
                    header_text = f"• {ev['start_time']}–{ev['end_time']}: {ev['title']}\nChủ trì: {ev['chair']}"
                    hcell = ws.cell(row=header_row, column=col_idx, value=header_text)
                    hcell.font = Font(bold=True)
                    hcell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

                    # ----- DETAILS (căn trái) -----
                    details = []
                    if ev.get("attendees"):
                        details.append(f"- Tham dự: {ev['attendees']}")
                    if ev.get("location"):
                        details.append(f"- Địa điểm: {ev['location']}")
                    if ev.get("category"):
                        details.append(f"- Loại: {ev['category']}")
                    if ev.get("conflict"):
                        details.append("⚠ Trùng giờ")

                    dcell = ws.cell(row=detail_row, column=col_idx, value="\n".join(details))
                    dcell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

                    # ----- NỀN THEO CHỦ TRÌ -----
                    hexcol = CHAIR_COLORS.get(ev["chair"])
                    if hexcol:
                        fill = PatternFill(start_color=excel_color(hexcol),
                                           end_color=excel_color(hexcol),
                                           fill_type="solid")
                        hcell.fill = fill
                        dcell.fill = fill
                else:
                    # Không có sự kiện -> hàng đầu hiển thị "—"
                    hcell = ws.cell(row=header_row, column=col_idx, value="—")
                    hcell.alignment = Alignment(horizontal='center', vertical='center')
                    ws.cell(row=detail_row, column=col_idx, value="")

        # Chiều cao hàng (header thấp hơn details)
        for r in range(start_row, end_row + 1, 2):
            ws.row_dimensions[r].height     = 34   # header
            ws.row_dimensions[r + 1].height = 88   # details

        start_row = end_row + 1  # sang buổi tiếp theo

    # Viền hàng header
    for c in range(1, 8):
        ws.cell(row=3, column=c).border = border

    # Độ rộng cột
    ws.column_dimensions['A'].width = 12
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 36

    # ----- Khu ghi chú -----
    notes_row = start_row + 1
    ws.merge_cells(start_row=notes_row, start_column=1, end_row=notes_row, end_column=6)
    ws.cell(
        row=notes_row, column=1,
        value="Ghi chú: Các cuộc họp phát sinh TL.BGĐ xin ý kiến BGĐ thống nhất -> HV cập nhật lên phần mềm"
    )
    ws.cell(
        row=notes_row, column=7,
        value=f"Đà Nẵng, Ngày {week_end.strftime('%d')} tháng {week_end.strftime('%m')} năm {week_end.strftime('%Y')}"
    ).alignment = Alignment(horizontal='center')

    ws.cell(row=notes_row+1, column=2, value="BGĐ KIỂM TRA")
    ws.cell(row=notes_row+1, column=4, value="TP.NS&ĐT Kiểm tra")
    ws.cell(row=notes_row+1, column=6, value="Người lập biểu")

    ws.cell(row=notes_row+4, column=2, value="Phan Thị Yến Tuyết")
    ws.cell(row=notes_row+4, column=4, value="Trần Thị Kim Oanh")
    ws.cell(row=notes_row+4, column=6, value="Trần Thị Mỹ Tân")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, f"lich_hop_tuan_{session['id']}.xlsx"




# ========== XUẤT ICS ==========
def export_session_to_ics(session):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//{COMPANY_NAME}//Meeting Calendar//VN"
    ]
    for ev in session["events"]:
        date = dt.date.fromisoformat(ev["date"])
        start = dt.datetime.combine(date, dt.time.fromisoformat(ev["start_time"] + ":00"))
        end = dt.datetime.combine(date, dt.time.fromisoformat(ev["end_time"] + ":00"))
        uid = ev["id"]
        title = ev["title"].replace("\n", " ")
        desc = []
        if ev.get("chair"): desc.append(f"Chu tri: {ev['chair']}")
        if ev.get("attendees"): desc.append(f"Tham du: {ev['attendees']}")
        if ev.get("category"): desc.append(f"Loai: {ev['category']}")
        description = "\\n".join(desc)
        def fmt(dtobj): return dtobj.strftime("%Y%m%dT%H%M%SZ")
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{fmt(dt.datetime.utcnow())}",
            f"DTSTART:{fmt(start)}",
            f"DTEND:{fmt(end)}",
            f"SUMMARY:{title}",
            f"DESCRIPTION:{description}",
            f"LOCATION:{ev.get('location','')}",
            "END:VEVENT"
        ]
    lines.append("END:VCALENDAR")
    ics_bytes = "\r\n".join(lines).encode("utf-8")
    return BytesIO(ics_bytes), f"lich_hop_tuan_{session['id']}.ics"

# ========== ROUTES ==========
@app.route("/")
def home():
    data = load_data()
    qdate = request.args.get("date")
    today = dt.date.today() if not qdate else dt.date.fromisoformat(qdate)
    sess = get_or_create_session(data, today)

    sessions_sorted = sorted(data["sessions"], key=lambda s: s["week_start"], reverse=True)

    q = request.args.get("q", "").strip().lower()
    events = list(sess["events"])
    if q:
        events = [e for e in events if q in json.dumps(e, ensure_ascii=False).lower()]

    compute_conflicts(events)

    return render_template_string(
        TEMPLATE_INDEX,
        company=COMPANY_NAME,
        chair_colors=CHAIR_COLORS,
        categories=CATEGORIES,
        rooms=ROOMS,
        session=sess,
        sessions=sessions_sorted,
        events=sorted(events, key=lambda x: (x["date"], x["session_buoi"], x["start_time"])),
        week_start=dt.date.fromisoformat(sess["week_start"]),
        week_end=dt.date.fromisoformat(sess["week_end"]),
        today=today,
        q=q
    )

@app.route("/preview/<session_id>")
def preview(session_id):
    data = load_data()
    sess = find_session_by_id(data, session_id)
    if not sess:
        return "Không tìm thấy session", 404
    dates, schedule = build_schedule(sess)
    weekdays = ['Thứ 2','Thứ 3','Thứ 4','Thứ 5','Thứ 6','Thứ 7']
    return render_template_string(
        TEMPLATE_PREVIEW,
        company=COMPANY_NAME,
        chair_colors=CHAIR_COLORS,
        session=sess,
        dates=dates,
        schedule=schedule,
        weekdays=weekdays
    )

@app.route("/sessions")
def list_sessions():
    data = load_data()
    sessions_sorted = sorted(data["sessions"], key=lambda s: s["week_start"], reverse=True)
    return jsonify(sessions_sorted)

@app.route("/switch-session", methods=["POST"])
def switch_session():
    date_str = request.form.get("any_date")
    if not date_str:
        return redirect(url_for("home"))
    return redirect(url_for("home", date=date_str))

@app.route("/event", methods=["POST"])
def add_or_update_event():
    data = load_data()
    date_str = request.form["date"]
    buoi = request.form.get("buoi") or guess_buoi(request.form["start_time"])
    sess = get_or_create_session(data, dt.date.fromisoformat(date_str))

    payload = {
        "id": request.form.get("id", ""),  # rỗng -> tạo mới
        "date": date_str,
        "buoi": buoi,
        "start_time": request.form["start_time"],
        "end_time": request.form["end_time"],
        "title": request.form["title"],
        "category": request.form.get("category", ""),
        "chair": request.form["chair"],
        "attendees": request.form.get("attendees", ""),
        "location": request.form.get("location", "")
    }
    try:
        upsert_event(sess, payload)
        save_data(data)
        return redirect(url_for("home", date=date_str))
    except ValueError as e:
        return f"Lỗi: {e}", 400

@app.route("/event/<session_id>/<event_id>/delete", methods=["POST"])
def remove_event(session_id, event_id):
    data = load_data()
    sess = find_session_by_id(data, session_id)
    if not sess:
        return "Không tìm thấy session", 404
    delete_event(sess, event_id)
    save_data(data)
    return redirect(url_for("home", date=sess["week_start"]))

@app.route("/event/<session_id>/clear", methods=["POST"])
def clear_session(session_id):
    data = load_data()
    sess = find_session_by_id(data, session_id)
    if not sess:
        return "Không tìm thấy session", 404
    sess["events"] = []
    save_data(data)
    return redirect(url_for("home", date=sess["week_start"]))

@app.route("/export/<session_id>/excel", methods=["POST"])
def export_excel(session_id):
    data = load_data()
    sess = find_session_by_id(data, session_id)
    if not sess:
        return "Không tìm thấy session", 404
    output, fname = export_session_to_excel(sess)
    return send_file(output,
                     as_attachment=True,
                     download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export/<session_id>/ics", methods=["POST"])
def export_ics(session_id):
    data = load_data()
    sess = find_session_by_id(data, session_id)
    if not sess:
        return "Không tìm thấy session", 404
    output, fname = export_session_to_ics(sess)
    return send_file(output,
                     as_attachment=True,
                     download_name=fname,
                     mimetype="text/calendar")

@app.route("/backup/json", methods=["GET"])
def backup_json():
    ensure_data_file()
    return send_file(DATA_PATH, as_attachment=True, download_name="meeting_schedule_backup.json")

# ========== TEMPLATES ==========
TEMPLATE_INDEX = """
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <title>Quản lý Lịch Họp {{ company }}</title>
  <style>
    :root { --bg:#f5f7fb; --card:#fff; --text:#2c3e50; --muted:#6b7280; --primary:#2563eb; --danger:#ef4444; }
    * { box-sizing: border-box }
    body { font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, "Apple Color Emoji", "Segoe UI Emoji"; margin:0; background:var(--bg); color:var(--text); }
    header { padding:16px 24px; background:#fff; border-bottom:1px solid #e5e7eb; position:sticky; top:0; z-index:10 }
    h1 { font-size:20px; margin:0 }
    .layout { display:grid; grid-template-columns: 300px 1fr; gap:16px; padding:16px; }
    .card { background:var(--card); border:1px solid #e5e7eb; border-radius:10px; box-shadow:0 1px 2px rgba(0,0,0,.03) }
    .card h2 { font-size:16px; padding:14px 16px; margin:0; border-bottom:1px solid #e5e7eb; background:#fafafa }
    .card .content { padding:16px }
    .muted { color:var(--muted); }
    .row { display:flex; gap:8px; }
    .row > * { flex:1 }
    input, select, button { padding:10px 12px; border:1px solid #e5e7eb; border-radius:8px; font:inherit; }
    button.primary { background:var(--primary); color:#fff; border-color:transparent; cursor:pointer }
    button.danger { background:var(--danger); color:#fff; border-color:transparent; cursor:pointer }
    table { width:100%; border-collapse: collapse; }
    th, td { text-align:left; padding:8px; border-bottom:1px solid #eee; vertical-align: top; }
    .legend { display:flex; flex-wrap:wrap; gap:8px; }
    .tag { display:inline-flex; align-items:center; gap:6px; padding:6px 10px; border:1px solid #e5e7eb; border-radius:999px; background:#fff; }
    .dot { width:12px; height:12px; border-radius:999px; border:1px solid #e5e7eb; }
    .conf { color:#b45309; font-weight:600; }
    .grid2 { display:grid; grid-template-columns: 1fr 1fr; gap:8px; }
    .grid3 { display:grid; grid-template-columns: repeat(3,1fr); gap:8px; }
    .nowrap { white-space: nowrap; }
    .toolbar { display:flex; gap:8px; flex-wrap: wrap; }
  </style>
</head>
<body>
  <header>
    <h1>Lịch họp – {{ company }}</h1>
  </header>

  <div class="layout">
    <!-- Sidebar -->
    <aside class="card">
      <h2>Tuần &amp; Tính năng</h2>
      <div class="content">
        <form method="post" action="/switch-session" class="row" style="margin-bottom:12px">
          <div>
            <label class="muted">Chọn bất kỳ ngày trong tuần</label>
            <input type="date" name="any_date" value="{{ today.isoformat() }}">
          </div>
          <div class="nowrap" style="align-self:flex-end">
            <button class="primary" type="submit">Mở tuần</button>
          </div>
        </form>

        <div style="margin:8px 0 12px" class="muted">
          Tuần hiện tại: <b>{{ week_start.strftime('%d/%m/%Y') }}</b> → <b>{{ week_end.strftime('%d/%m/%Y') }}</b>
        </div>

        <div style="margin-top:6px">
          <form method="get" action="/" class="row">
            <input type="hidden" name="date" value="{{ week_start.isoformat() }}">
            <input type="text" name="q" placeholder="Tìm kiếm..." value="{{ q }}">
            <button type="submit">Lọc</button>
          </form>
        </div>

        <hr style="margin:14px 0">

        <div class="legend">
          {% for chair, color in chair_colors.items() %}
            <span class="tag"><span class="dot" style="background: {{ color }}"></span>{{ chair }}</span>
          {% endfor %}
        </div>

        <hr style="margin:14px 0">

        <div class="toolbar">
          <a href="/preview/{{ session.id }}" target="_blank"><button type="button">Xem trước (mẫu Excel)</button></a>
          <form method="post" action="/export/{{ session.id }}/excel">
            <button class="primary" type="submit">Export Excel</button>
          </form>
          <form method="post" action="/export/{{ session.id }}/ics">
            <button type="submit">Export ICS</button>
          </form>
          <a href="/backup/json"><button type="button">Backup JSON</button></a>
          <form method="post" action="/event/{{ session.id }}/clear" onsubmit="return confirm('Xoá toàn bộ sự kiện của tuần này?')">
            <button class="danger" type="submit">Xoá toàn tuần</button>
          </form>
        </div>

        <hr style="margin:14px 0">

        <div>
          <div class="muted" style="margin-bottom:6px">Các tuần gần đây</div>
          <div style="max-height:260px; overflow:auto">
            <table>
              <thead><tr><th>Tuần</th><th class="nowrap">Mở</th></tr></thead>
              <tbody>
              {% for s in sessions %}
                <tr>
                  <td><b>{{ s.id }}</b><br><span class="muted">{{ s.week_start }} → {{ s.week_end }}</span></td>
                  <td class="nowrap">
                    <a href="/?date={{ s.week_start }}"><button type="button">Xem</button></a>
                  </td>
                </tr>
              {% endfor %}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </aside>

    <!-- Main: Form & List -->
    <main class="card">
      <h2>Thêm/Chỉnh sửa sự kiện</h2>
      <div class="content">
        <form id="event-form" method="post" action="/event">
          <input type="hidden" name="id" id="fld-id">
          <div class="grid3">
            <div>
              <label>Ngày</label>
              <input type="date" name="date" id="fld-date" value="{{ week_start.isoformat() }}" required>
            </div>
            <div>
              <label>Buổi</label>
              <select name="buoi" id="fld-buoi">
                <option value="">(Tự nhận diện theo giờ)</option>
                <option value="SÁNG">SÁNG</option>
                <option value="CHIỀU">CHIỀU</option>
              </select>
            </div>
            <div>
              <label>Loại</label>
              <input type="hidden" name="category" id="fld-category">
              <select id="fld-category-select">
                {% for c in categories %}
                  <option value="{{ c }}">{{ c }}</option>
                {% endfor %}
                <option value="__OTHER__">Khác…</option>
              </select>
              <input type="text" id="fld-category-other" placeholder="Nhập loại khác" style="display:none; margin-top:6px">
            </div>
          </div>

          <div class="grid3" style="margin-top:8px">
            <div>
              <label>Giờ bắt đầu</label>
              <input type="time" name="start_time" id="fld-start" required>
            </div>
            <div>
              <label>Giờ kết thúc</label>
              <input type="time" name="end_time" id="fld-end" required>
            </div>
            <div>
              <label>Chủ trì</label>
              <select name="chair" id="fld-chair" required>
                {% for chair, color in chair_colors.items() %}
                  <option value="{{ chair }}">{{ chair }}</option>
                {% endfor %}
              </select>
            </div>
          </div>

          <div class="grid2" style="margin-top:8px">
            <div>
              <label>Tên họp</label>
              <input type="text" name="title" id="fld-title" placeholder="VD: Họp giao ban tuần" required>
            </div>
            <div>
              <label>Địa điểm</label>
              <input type="hidden" name="location" id="fld-location">
              <select id="fld-location-select">
                {% for r in rooms %}
                  <option value="{{ r }}">{{ r }}</option>
                {% endfor %}
                <option value="__OTHER__">Khác…</option>
              </select>
              <input type="text" id="fld-location-other" placeholder="Nhập địa điểm" style="display:none; margin-top:6px">
            </div>
          </div>

          <div style="margin-top:8px">
            <label>Tham dự</label>
            <input type="text" name="attendees" id="fld-attendees" placeholder="VD: CEO, COO">
          </div>

          <div class="row" style="margin-top:12px">
            <button class="primary" type="submit">Lưu sự kiện</button>
            <button type="reset" onclick="document.getElementById('fld-id').value=''">Xoá nhập</button>
          </div>
        </form>
      </div>

      <h2 style="border-top:1px solid #e5e7eb">Sự kiện trong tuần</h2>
      <div class="content">
        {% if events %}
          <table>
            <thead>
              <tr>
                <th class="nowrap">Ngày</th>
                <th>Buổi</th>
                <th>Giờ</th>
                <th>Tiêu đề</th>
                <th>Chủ trì</th>
                <th>Tham dự</th>
                <th>Địa điểm</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
            {% for ev in events %}
              <tr
                data-id="{{ ev.id }}"
                data-date="{{ ev.date }}"
                data-buoi="{{ ev.session_buoi }}"
                data-start="{{ ev.start_time }}"
                data-end="{{ ev.end_time }}"
                data-title="{{ ev.title|e }}"
                data-chair="{{ ev.chair }}"
                data-attendees="{{ ev.attendees|e }}"
                data-location="{{ ev.location|e }}"
                data-category="{{ ev.category|e }}"
              >
                <td class="nowrap">{{ ev.date }}</td>
                <td>{{ ev.session_buoi }}</td>
                <td class="nowrap">{{ ev.start_time }}–{{ ev.end_time }} {% if ev.conflict %}<span class="conf">⚠ Trùng</span>{% endif %}</td>
                <td>
                  <div style="font-weight:600">{{ ev.title }}</div>
                  {% if ev.category %}<div class="muted">Loại: {{ ev.category }}</div>{% endif %}
                </td>
                <td>{{ ev.chair }}</td>
                <td>{{ ev.attendees }}</td>
                <td>{{ ev.location }}</td>
                <td class="nowrap">
                  <div style="display:flex; gap:6px">
                    <button type="button" onclick="editEvent(this)">Sửa</button>
                    <form method="post" action="/event/{{ session.id }}/{{ ev.id }}/delete" onsubmit="return confirm('Xoá sự kiện này?')">
                      <button class="danger" type="submit">Xoá</button>
                    </form>
                  </div>
                </td>
              </tr>
            {% endfor %}
            </tbody>
          </table>
        {% else %}
          <div class="muted">Chưa có sự kiện nào trong tuần này.</div>
        {% endif %}
      </div>
    </main>
  </div>

  <script>
    // —— Select "Khác…" logic (category & location) ——
    function setupOther(selectId, otherId, hiddenId) {
      const sel = document.getElementById(selectId);
      const other = document.getElementById(otherId);
      const hidden = document.getElementById(hiddenId);

      function syncHidden() {
        if (sel.value === '__OTHER__') {
          other.style.display = '';
          hidden.value = other.value.trim();
        } else {
          other.style.display = 'none';
          hidden.value = sel.value;
        }
      }
      sel.addEventListener('change', syncHidden);
      other.addEventListener('input', syncHidden);
      syncHidden();
    }
    setupOther('fld-category-select', 'fld-category-other', 'fld-category');
    setupOther('fld-location-select', 'fld-location-other', 'fld-location');

    // —— Edit button: đổ dữ liệu vào form ——
    function setSelectOrOther(selectId, otherId, hiddenId, value) {
      const sel = document.getElementById(selectId);
      const other = document.getElementById(otherId);
      const hidden = document.getElementById(hiddenId);
      const exists = Array.from(sel.options).some(opt => opt.value === value);
      if (exists) {
        sel.value = value;
        other.style.display = 'none';
        hidden.value = value;
      } else {
        sel.value = '__OTHER__';
        other.style.display = '';
        other.value = value || '';
        hidden.value = other.value;
      }
    }

    function editEvent(btn) {
      const tr = btn.closest('tr');
      const g = (k) => tr.dataset[k] || '';

      document.getElementById('fld-id').value = g('id');
      document.getElementById('fld-date').value = g('date');
      document.getElementById('fld-buoi').value = g('buoi');
      document.getElementById('fld-start').value = g('start');
      document.getElementById('fld-end').value = g('end');
      document.getElementById('fld-title').value = g('title');
      document.getElementById('fld-chair').value = g('chair');
      document.getElementById('fld-attendees').value = g('attendees');

      setSelectOrOther('fld-category-select', 'fld-category-other', 'fld-category', g('category'));
      setSelectOrOther('fld-location-select', 'fld-location-other', 'fld-location', g('location'));

      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  </script>
</body>
</html>
"""

# —— PREVIEW TEMPLATE: mỗi sự kiện là 1 khối màu theo Chủ trì ——
TEMPLATE_PREVIEW = """
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <title>Preview lịch tuần – {{ company }}</title>
  <style>
    body { font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial; background:#f6f7fb; margin:0; padding:24px; color:#111827; }
    h1 { margin:0 0 16px 0; font-size:20px; }
    .table { border:1px solid #e5e7eb; background:#fff; border-radius:10px; overflow:hidden }
    .head { display:grid; grid-template-columns: 120px repeat(6, 1fr); background:#f9fafb; border-bottom:1px solid #e5e7eb; }
    .head div { padding:10px 12px; font-weight:600; text-align:center; }
    .row { display:grid; grid-template-columns: 120px repeat(6, 1fr); border-bottom:1px solid #f3f4f6; }
    .buoi { background:#f9fafb; padding:10px 12px; font-weight:700; text-align:center; border-right:1px solid #e5e7eb }
    .cell { padding:8px; min-height:180px; border-right:1px solid #f3f4f6; }
    .ev { padding:8px; margin-bottom:8px; border-radius:8px; box-shadow: inset 0 0 0 1px rgba(0,0,0,.05); background:#f3f4f6 }
    .ev .tt { font-weight:700 }
    .muted { color:#6b7280; }
    .warn { color:#b45309; font-weight:700 }
    .legend { margin-top:16px; display:flex; gap:8px; flex-wrap:wrap }
    .tag { display:inline-flex; align-items:center; gap:6px; padding:6px 10px; border:1px solid #e5e7eb; border-radius:999px; background:#fff; }
    .dot { width:12px; height:12px; border-radius:999px; border:1px solid #e5e7eb; display:inline-block }
    @media print {
      body { padding:0; }
      .legend { display:none }
    }
  </style>
</head>
<body>
  <h1>Preview lịch tuần – {{ company }} ({{ session.week_start }} → {{ session.week_end }})</h1>

  <div class="table">
    <div class="head">
      <div>Buổi</div>
      {% for d in dates %}
        <div>{{ weekdays[loop.index0] }}<br><span class="muted">({{ d.strftime('%d.%m.%Y') }})</span></div>
      {% endfor %}
    </div>


    {% for buoi in ['SÁNG','CHIỀU'] %}
      <div class="row">
        <div class="buoi">{{ buoi }}</div>
        {% for d in dates %}
          {% set key = d.strftime('%d/%m/%Y') %}
          <div class="cell">
            {% for ev in schedule.get(key, {}).get(buoi, []) %}
              {% set bg = chair_colors.get(ev.chair, '#f3f4f6') %}
              <div class="ev" style="background: {{ bg }}">
                <div class="tt">• {{ ev.start_time }}–{{ ev.end_time }}: {{ ev.title }}</div>
                <div>Chủ trì: <b>{{ ev.chair }}</b></div>
                {% if ev.attendees %}<div>- Tham dự: {{ ev.attendees }}</div>{% endif %}
                {% if ev.location %}<div>- Địa điểm: {{ ev.location }}</div>{% endif %}
                {% if ev.category %}<div>- Loại: {{ ev.category }}</div>{% endif %}
                {% if ev.conflict %}<div class="warn">⚠ Trùng giờ</div>{% endif %}
              </div>
            {% else %}
              <div class="muted" style="font-style:italic">—</div>
            {% endfor %}
          </div>
        {% endfor %}
      </div>
    {% endfor %}
  </div>

  <div class="legend">
    {% for chair, color in chair_colors.items() %}
      <span class="tag"><span class="dot" style="background: {{ color }}"></span>{{ chair }}</span>
    {% endfor %}
  </div>
</body>
</html>
"""

# ========== MAIN ==========
if __name__ == "__main__":
    ensure_data_file()
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
