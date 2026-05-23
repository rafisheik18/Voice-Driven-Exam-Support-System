from flask import Flask, render_template, request, jsonify, redirect, url_for, session, send_file
from datetime import datetime
import os, sqlite3, json, re
from utils.audio_utils import generate_pdf_answer
from cryptography.fernet import Fernet
import openpyxl   # ✅ Excel support

app = Flask(__name__)
app.secret_key = "supersecret"

DB_FILE = "exam_results.db"
KEY_FILE = "secret.key"
USERS_FILE = "users.xlsx"   # ✅ Excel file with student/teacher accounts

# -------------------
# Load Users from Excel
# -------------------
def load_users():
    users = {}
    wb = openpyxl.load_workbook(USERS_FILE)
    sheet = wb.active
    for row in sheet.iter_rows(min_row=2, values_only=True):  # skip header
        username, password, role = row
        if username and password and role:
            users[str(username).strip()] = {
                "password": str(password).strip(),
                "role": str(role).strip().lower()
            }
    return users

# preload users into memory
users = load_users()
print("✅ Loaded users from Excel:", users)

# -------------------
# Encryption Setup
# -------------------
if not os.path.exists(KEY_FILE):
    with open(KEY_FILE, "wb") as f:
        f.write(Fernet.generate_key())
fernet = Fernet(open(KEY_FILE, "rb").read())

# -------------------
# DB Setup
# -------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS results (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student TEXT,
                    score TEXT,
                    total TEXT,
                    timestamp TEXT
                )""")
    conn.commit()
    conn.close()

init_db()

# -------------------
# Load Questions
# -------------------
def load_questions():
    with open("exam_questions.json", "r", encoding="utf-8") as f:
        return json.load(f)

# -------------------
# Routes
# -------------------
@app.route('/')
def home():
    return render_template("index.html")

@app.route('/login', methods=['POST'])
def login():
    username = request.form['username'].strip()
    password = request.form['password'].strip()

    if username in users and users[username]["password"] == password:
        session['user'] = username
        role = users[username]["role"]

        # Teacher login
        if role == "teacher":
            if request.is_json or request.headers.get("X-Requested-With") == "XMLHttpRequest":
                return jsonify({"status": "ok", "redirect": url_for('teacher_dashboard')})
            else:
                return redirect(url_for('teacher_dashboard'))

        # Student login
        elif role == "student":
            if request.is_json or request.headers.get("X-Requested-With") == "XMLHttpRequest":
                return jsonify({"status": "ok", "redirect": url_for('student_exam')})
            else:
                return redirect(url_for('student_exam'))

    # Invalid credentials
    if request.is_json or request.headers.get("X-Requested-With") == "XMLHttpRequest":
        return jsonify({"status": "error", "message": "Invalid credentials"})
    else:
        return render_template("index.html", error="Invalid credentials")

# -------------------
# Student Exam
# -------------------
@app.route('/student')
def student_exam():
    if 'user' not in session or users[session['user']]["role"] != "student":
        return redirect("/")
    questions = load_questions()
    return render_template("student.html", questions=questions)

@app.route('/submit-exam', methods=['POST'])
def submit_exam():
    if 'user' not in session:
        return redirect("/")
    answers = request.json.get("answers", {})
    student = session['user']
    questions = load_questions()

    score = 0
    total = len(questions)

    # Save answers to file
    os.makedirs("answers", exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    file_txt = f"answers/{student}_{timestamp}.txt"
    file_pdf = f"answers/{student}_{timestamp}.pdf"

    with open(file_txt, "w", encoding="utf-8") as f:
        for q in questions:
            qid = str(q["id"])
            q_text = q["question"]
            correct = q["answer"]
            ans = answers.get(qid, "Not Answered")
            f.write(f"Q{qid}: {q_text}\nAnswer: {ans}\nCorrect: {correct}\n\n")
            if ans.lower() == correct.lower():
                score += 1

    # Generate PDF
    generate_pdf_answer(student, questions, answers, timestamp, file_pdf)

    # Encrypt before saving in DB
    enc_score = fernet.encrypt(str(score).encode()).decode()
    enc_total = fernet.encrypt(str(total).encode()).decode()

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO results (student, score, total, timestamp) VALUES (?,?,?,?)",
              (student, enc_score, enc_total, timestamp))
    conn.commit()
    conn.close()

    return jsonify({"status": "ok", "score": score, "total": total})

# -------------------
# Teacher Dashboard
# -------------------
@app.route('/teacher')
def teacher_dashboard():
    if 'user' not in session or users[session['user']]["role"] != "teacher":
        return redirect("/")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT student, score, total, timestamp FROM results ORDER BY timestamp DESC")
    rows = c.fetchall()
    conn.close()

    # decrypt scores
    decrypted = []
    for r in rows:
        dec_score = fernet.decrypt(r[1].encode()).decode()
        dec_total = fernet.decrypt(r[2].encode()).decode()
        decrypted.append((r[0], dec_score, dec_total, r[3]))

    return render_template("teacher.html", results=decrypted)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join("answers", filename), as_attachment=True)

@app.route('/teacher/view/<student>/<timestamp>')
def view_answers(student, timestamp):
    file_txt = f"answers/{student}_{timestamp}.txt"
    if os.path.exists(file_txt):
        with open(file_txt, "r", encoding="utf-8") as f:
            content = f.read()
        return f"<pre>{content}</pre>"
    else:
        return "Answer file not found"

@app.route('/teacher/analytics')
def teacher_analytics():
    questions = load_questions()
    q_stats = {q["id"]: {"question": q["question"], "correct": 0, "wrong": 0, "not_answered": 0} for q in questions}

    global_correct = global_wrong = global_not_answered = 0

    for fname in os.listdir("answers"):
        if fname.endswith(".txt"):
            with open(os.path.join("answers", fname), "r", encoding="utf-8") as f:
                lines = f.read().splitlines()

            qid = None
            student_ans = None
            correct_ans = None
            for line in lines:
                if re.match(r"^Q\d+:", line):
                    qid = int(line.split(":")[0].replace("Q", "").strip())
                elif line.startswith("Answer:"):
                    student_ans = line.replace("Answer:", "").strip()
                elif line.startswith("Correct:"):
                    correct_ans = line.replace("Correct:", "").strip()
                    if qid in q_stats:
                        if student_ans == "Not Answered":
                            q_stats[qid]["not_answered"] += 1
                            global_not_answered += 1
                        elif student_ans.lower() == correct_ans.lower():
                            q_stats[qid]["correct"] += 1
                            global_correct += 1
                        else:
                            q_stats[qid]["wrong"] += 1
                            global_wrong += 1

    summary = {
        "correct": global_correct,
        "wrong": global_wrong,
        "not_answered": global_not_answered
    }

    return render_template("analytics.html", stats=q_stats, summary=summary)

# -------------------
# Logout
# -------------------
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('home'))

# -------------------
# Run Server
# -------------------
if __name__ == "__main__":
    app.run(debug=True, port=5500)
