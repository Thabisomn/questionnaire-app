from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import sqlite3, io, os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = "change-this-secret-key-in-production"
app.jinja_env.globals.update(enumerate=enumerate)
app.jinja_env.globals.update(enumerate=enumerate)
DB = "responses.db"

# ── Database ──────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            submitted_at TEXT,
            consent TEXT,
            gender TEXT,
            age_group TEXT,
            district TEXT,
            stakeholder_category TEXT,
            education TEXT,
            -- Section B: Participation
            b1_planning INTEGER,
            b2_decision_making INTEGER,
            b3_implementation INTEGER,
            b4_consulted INTEGER,
            -- Section B: Communication
            b5_info_communicated INTEGER,
            b6_regular_updates INTEGER,
            b7_effective_communication INTEGER,
            b8_informed_decisions INTEGER,
            b9_feedback_considered INTEGER,
            -- Section B: Collaboration
            b10_work_together INTEGER,
            b11_cooperation INTEGER,
            b12_share_responsibilities INTEGER,
            b13_partnerships INTEGER,
            -- Section B: Empowerment
            b14_influence INTEGER,
            b15_leadership_roles INTEGER,
            b16_trained INTEGER,
            b17_capacity INTEGER,
            -- Section C: Project Sustainability
            c1_continues_after_funding INTEGER,
            c2_maintains_infrastructure INTEGER,
            c3_meets_needs INTEGER,
            c4_systems_in_place INTEGER,
            -- Section D: Stakeholder Perceptions
            d1_feel_included INTEGER,
            d2_trust_leadership INTEGER,
            d3_satisfied_participation INTEGER,
            d4_engagement_fair INTEGER,
            -- Section E: Barriers
            e1_lack_resources INTEGER,
            e2_power_imbalances INTEGER,
            e3_lack_information INTEGER,
            e4_weak_institutions INTEGER,
            -- Section F: Enablers
            f1_strong_leadership INTEGER,
            f2_training INTEGER,
            f3_good_communication INTEGER,
            f4_government_support INTEGER
        )""")

init_db()

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET"])
def index():
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
    f = request.form

    consent = f.get("consent", "No")
    if consent != "Yes":
        flash("You must consent to participate in this study.", "warning")
        return redirect(url_for("index"))

    data = (
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        consent,
        f.get("gender"), f.get("age_group"), f.get("district"),
        f.get("stakeholder_category"), f.get("education"),
        # Section B Participation
        f.get("b1"), f.get("b2"), f.get("b3"), f.get("b4"),
        # Section B Communication
        f.get("b5"), f.get("b6"), f.get("b7"), f.get("b8"), f.get("b9"),
        # Section B Collaboration
        f.get("b10"), f.get("b11"), f.get("b12"), f.get("b13"),
        # Section B Empowerment
        f.get("b14"), f.get("b15"), f.get("b16"), f.get("b17"),
        # Section C
        f.get("c1"), f.get("c2"), f.get("c3"), f.get("c4"),
        # Section D
        f.get("d1"), f.get("d2"), f.get("d3"), f.get("d4"),
        # Section E
        f.get("e1"), f.get("e2"), f.get("e3"), f.get("e4"),
        # Section F
        f.get("f1"), f.get("f2"), f.get("f3"), f.get("f4"),
    )

    with get_db() as conn:
        conn.execute("""
            INSERT INTO responses (
                submitted_at, consent, gender, age_group, district,
                stakeholder_category, education,
                b1_planning, b2_decision_making, b3_implementation, b4_consulted,
                b5_info_communicated, b6_regular_updates, b7_effective_communication,
                b8_informed_decisions, b9_feedback_considered,
                b10_work_together, b11_cooperation, b12_share_responsibilities, b13_partnerships,
                b14_influence, b15_leadership_roles, b16_trained, b17_capacity,
                c1_continues_after_funding, c2_maintains_infrastructure, c3_meets_needs, c4_systems_in_place,
                d1_feel_included, d2_trust_leadership, d3_satisfied_participation, d4_engagement_fair,
                e1_lack_resources, e2_power_imbalances, e3_lack_information, e4_weak_institutions,
                f1_strong_leadership, f2_training, f3_good_communication, f4_government_support
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, data)

    return render_template("thankyou.html")

@app.route("/admin")
def admin():
    with get_db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM responses").fetchone()[0]
        today = conn.execute("SELECT COUNT(*) FROM responses WHERE DATE(submitted_at) = DATE('now')").fetchone()[0]
        this_week = conn.execute("SELECT COUNT(*) FROM responses WHERE submitted_at >= DATE('now', '-7 days')").fetchone()[0]
        gender_rows = conn.execute("SELECT gender, COUNT(*) as cnt FROM responses GROUP BY gender").fetchall()
        gender = {r['gender']: r['cnt'] for r in gender_rows}
        district_rows = conn.execute("SELECT district, COUNT(*) as cnt FROM responses GROUP BY district ORDER BY cnt DESC").fetchall()
        districts = [(r['district'], r['cnt']) for r in district_rows]
        cat_rows = conn.execute("SELECT stakeholder_category, COUNT(*) as cnt FROM responses GROUP BY stakeholder_category ORDER BY cnt DESC").fetchall()
        categories = [(r['stakeholder_category'], r['cnt']) for r in cat_rows]
        recent = conn.execute("SELECT submitted_at, gender, age_group, district, stakeholder_category FROM responses ORDER BY submitted_at DESC LIMIT 10").fetchall()
    from datetime import datetime as dt
    hour = dt.now().hour
    greeting = "Good morning" if hour < 12 else "Good afternoon" if hour < 17 else "Good evening"
    now_str = dt.now().strftime("%A, %d %B %Y")
    return render_template("admin.html",
        total=total, today=today, this_week=this_week,
        gender=gender, districts=districts, categories=categories,
        recent=recent, greeting=greeting, now=now_str
    )

@app.route("/download")
def download():
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM responses ORDER BY submitted_at DESC").fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Responses"

    # Header styling
    header_fill = PatternFill("solid", fgColor="1F4E79")
    section_fill = PatternFill("solid", fgColor="2E75B6")
    header_font = Font(color="FFFFFF", bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = [
        ("ID", "Meta"), ("Submitted At", "Meta"), ("Consent", "Meta"),
        ("Gender", "Section A"), ("Age Group", "Section A"), ("District", "Section A"),
        ("Stakeholder Category", "Section A"), ("Education Level", "Section A"),
        # B - Participation
        ("Involved in project planning", "B: Participation"),
        ("Community participates in decisions", "B: Participation"),
        ("Stakeholders in implementation", "B: Participation"),
        ("Community consulted regularly", "B: Participation"),
        # B - Communication
        ("Info clearly communicated", "B: Communication"),
        ("Regular updates received", "B: Communication"),
        ("Effective communication", "B: Communication"),
        ("Informed about decisions", "B: Communication"),
        ("Feedback considered", "B: Communication"),
        # B - Collaboration
        ("Work together in activities", "B: Collaboration"),
        ("Cooperation with leaders", "B: Collaboration"),
        ("Share responsibilities", "B: Collaboration"),
        ("Govt-NGO-Community partnerships", "B: Collaboration"),
        # B - Empowerment
        ("Influence over decisions", "B: Empowerment"),
        ("Community given leadership roles", "B: Empowerment"),
        ("Stakeholders trained", "B: Empowerment"),
        ("Local capacity to sustain", "B: Empowerment"),
        # C
        ("Continues after funding ends", "C: Sustainability"),
        ("Maintains infrastructure", "C: Sustainability"),
        ("Meets long-term needs", "C: Sustainability"),
        ("Systems to sustain project", "C: Sustainability"),
        # D
        ("Feel included in decisions", "D: Perceptions"),
        ("Trust project leadership", "D: Perceptions"),
        ("Satisfied with participation", "D: Perceptions"),
        ("Engagement is fair", "D: Perceptions"),
        # E
        ("Lack of resources limits participation", "E: Barriers"),
        ("Power imbalances affect participation", "E: Barriers"),
        ("Lack of info limits engagement", "E: Barriers"),
        ("Weak institutions affect sustainability", "E: Barriers"),
        # F
        ("Strong leadership improves engagement", "F: Enablers"),
        ("Training improves participation", "F: Enablers"),
        ("Good communication enhances success", "F: Enablers"),
        ("Government support improves sustainability", "F: Enablers"),
    ]

    # Row 1: Section groups
    # Row 2: Column headers
    sections = {}
    for col_idx, (label, section) in enumerate(headers, start=1):
        if section not in sections:
            sections[section] = [col_idx, col_idx]
        else:
            sections[section][1] = col_idx

    # Write section row
    ws.append([""] * len(headers))
    for section, (start, end) in sections.items():
        cell = ws.cell(row=1, column=start, value=section)
        cell.fill = section_fill
        cell.font = header_font
        cell.alignment = center
        if start != end:
            ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)

    # Write column header row
    for col_idx, (label, _) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    # Write data rows
    for row in rows:
        ws.append(list(row))

    # Column widths
    for col_idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 22

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 50

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"questionnaire_responses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True, port=5001)
