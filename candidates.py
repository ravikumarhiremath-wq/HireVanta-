
from flask import Blueprint, request, render_template_string, send_file
from pathlib import Path
from io import BytesIO
from openpyxl import Workbook
import sqlite3

candidates_bp = Blueprint("candidates", __name__)

BASE_DIR = Path(__file__).resolve().parent
DATABASE = BASE_DIR / "hirevanta.db"


def get_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


def initialize_candidates():
    with get_connection() as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS candidates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT,
                phone TEXT,
                skills TEXT,
                experience TEXT,
                location TEXT,
                source TEXT,
                profile_url TEXT,
                open_to_work TEXT
            )
        """)


initialize_candidates()


@candidates_bp.route("/candidates", methods=["GET", "POST"])
def candidate_page():
    message = ""

    if request.method == "POST":
        name = request.form.get("name", "").strip()

        if name:
            with get_connection() as connection:
                connection.execute("""
                    INSERT INTO candidates
                    (name, email, phone, skills, experience, location,
                     source, profile_url, open_to_work)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    name,
                    request.form.get("email", "").strip(),
                    request.form.get("phone", "").strip(),
                    request.form.get("skills", "").strip(),
                    request.form.get("experience", "").strip(),
                    request.form.get("location", "").strip(),
                    request.form.get("source", "").strip(),
                    request.form.get("profile_url", "").strip(),
                    request.form.get("open_to_work", "Not confirmed")
                ))
            message = "Candidate saved successfully!"
        else:
            message = "Candidate name is required."

    with get_connection() as connection:
        rows = connection.execute(
            "SELECT * FROM candidates ORDER BY id DESC"
        ).fetchall()

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>HireVanta | Candidates</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background: #f4f6fb;
                margin: 0;
                padding: 30px;
                color: #18213a;
            }
            main {
                max-width: 1000px;
                margin: auto;
            }
            .panel {
                background: white;
                padding: 24px;
                margin-bottom: 22px;
                border-radius: 14px;
                box-shadow: 0 4px 18px #18213a10;
            }
            .fields {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
                gap: 12px;
            }
            input, select {
                width: 100%;
                box-sizing: border-box;
                padding: 12px;
                border: 1px solid #d8ddeb;
                border-radius: 7px;
                margin-top: 5px;
            }
            label { font-size: 13px; font-weight: bold; }
            button, .download {
                display: inline-block;
                border: 0;
                border-radius: 7px;
                padding: 12px 18px;
                background: #5145cd;
                color: white;
                text-decoration: none;
                cursor: pointer;
                margin-top: 16px;
            }
            .message { color: #18794e; }
            .table-wrap { overflow-x: auto; }
            table { border-collapse: collapse; width: 100%; }
            th, td {
                padding: 12px;
                border-bottom: 1px solid #e6e9f1;
                text-align: left;
                font-size: 13px;
            }
            th { background: #f4f6fb; }
        </style>
    </head>
    <body>
    <main>
        <h1>HireVanta</h1>
        <p>Candidate sourcing workspace</p>

        <section class="panel">
            <h2>Add a candidate</h2>
            <p>Enter profile and contact details provided by the candidate or obtained from a source you are authorized to use.</p>
            <p class="message">{{ message }}</p>

            <form method="POST">
                <div class="fields">
                    <label>Full name *
                        <input name="name" required>
                    </label>
                    <label>Email
                        <input name="email" type="email">
                    </label>
                    <label>Phone
                        <input name="phone" type="tel">
                    </label>
                    <label>Skills
                        <input name="skills" placeholder="Python, SQL, Excel">
                    </label>
                    <label>Experience
                        <input name="experience" placeholder="e.g. Fresher, 2 years">
                    </label>
                    <label>Location
                        <input name="location" placeholder="City">
                    </label>
                    <label>Profile source
                        <select name="source">
                            <option value="">Select source</option>
                            <option>LinkedIn</option>
                            <option>Naukri</option>
                            <option>Indeed</option>
                            <option>Other</option>
                        </select>
                    </label>
                    <label>Profile URL
                        <input name="profile_url" type="url" placeholder="https://...">
                    </label>
                    <label>Open to work status
                        <select name="open_to_work">
                            <option>Not confirmed</option>
                            <option>Yes</option>
                            <option>No</option>
                        </select>
                    </label>
                </div>
                <button type="submit">Save candidate</button>
            </form>
        </section>

        <section class="panel">
            <h2>Saved candidates ({{ rows|length }})</h2>
            <p>This spreadsheet contains candidates added manually. For profiles from a job search, open <a href="/searches">Saved searches</a>.</p>
            <a class="download" href="/candidates/export">Download manually saved candidates (Excel)</a>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th>Name</th><th>Email</th><th>Phone</th><th>Skills</th>
                            <th>Experience</th><th>Location</th><th>Source</th><th>Profile</th>
                            <th>Open to work</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for c in rows %}
                        <tr>
                            <td>{{ c["name"] }}</td>
                            <td>{{ c["email"] }}</td>
                            <td>{{ c["phone"] }}</td>
                            <td>{{ c["skills"] }}</td>
                            <td>{{ c["experience"] }}</td>
                            <td>{{ c["location"] }}</td>
                            <td>{{ c["source"] }}</td>
                            <td>
                                {% if c["profile_url"] %}
                                    <a href="{{ c['profile_url'] }}" target="_blank" rel="noopener noreferrer">Open profile</a>
                                {% else %}
                                    -
                                {% endif %}
                            </td>
                            <td>{{ c["open_to_work"] }}</td>
                        </tr>
                    {% else %}
                        <tr><td colspan="9">No candidates added yet.</td></tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
        </section>
        <p><a href="/">Back to JD upload</a></p>
    </main>
    </body>
    </html>
    """, rows=rows, message=message)


@candidates_bp.route("/candidates/export")
def export_candidates():
    with get_connection() as connection:
        rows = connection.execute("""
            SELECT name, email, phone, skills, experience, location,
                   source, profile_url, open_to_work
            FROM candidates ORDER BY id DESC
        """).fetchall()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Candidates"

    sheet.append([
        "Name", "Email", "Phone", "Skills", "Experience",
        "Location", "Source", "Profile URL", "Open to Work"
    ])

    for row in rows:
        sheet.append([row[key] or "" for key in row.keys()])

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column in sheet.columns:
        letter = column[0].column_letter
        sheet.column_dimensions[letter].width = min(
            max(max(len(str(cell.value or "")) for cell in column) + 2, 14),
            35
        )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="HireVanta_Candidates.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
