from flask import Flask, render_template, request, redirect, session, flash
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash

from routes.research_routes import research_bp


app = Flask(__name__)
app.secret_key = "secret123"

app.register_blueprint(research_bp)

# ---------------- HOME ----------------
@app.route("/")
def home():
    return render_template("home.html")


# ---------------- DASHBOARD (Protected) ----------------
@app.route("/dashboard")
def dashboard():

    if "user_id" not in session:
        flash("Please login first.", "error")
        return redirect("/")

    return render_template("dashboard.html")


# ---------------- TEXT PAGE (Protected) ----------------
@app.route("/text")
def text():

    if "user_id" not in session:
        flash("Please login first.", "error")
        return redirect("/")

    return render_template("text.html")


# ---------------- RESEARCH PAGE (Protected) ----------------
@app.route("/research")
def research():

    if "user_id" not in session:
        flash("Please login first.", "error")
        return redirect("/")

    return render_template("research.html")


# ---------------- PDF PAGE (Protected) ----------------
@app.route("/pdf")
def pdf():

    if "user_id" not in session:
        flash("Please login first.", "error")
        return redirect("/")

    return render_template("pdf.html")


# ---------------- SIGNUP ----------------
@app.route("/signup", methods=["POST"])
def signup():

    name = request.form.get("name")
    email = request.form.get("email")
    password = request.form.get("password")

    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()

    # check if user exists
    cursor.execute("SELECT * FROM users WHERE email = ?", (email,))
    existing_user = cursor.fetchone()

    if existing_user:
        conn.close()
        flash("Account already exists. Please login.", "error")
        return redirect("/")

    # hash password
    hashed_password = generate_password_hash(password)

    # insert user
    cursor.execute(
        "INSERT INTO users (name, email, password) VALUES (?, ?, ?)",
        (name, email, hashed_password)
    )

    conn.commit()
    conn.close()

    flash("Account created successfully! Please login.", "success")

    return redirect("/")


# ---------------- LOGIN ----------------
@app.route("/login", methods=["POST"])
def login():

    email = request.form.get("email")
    password = request.form.get("password")

    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM users WHERE email = ?", (email,))
    user = cursor.fetchone()

    conn.close()

    # user not found
    if not user:
        flash("First sign up. Account not found.", "error")
        return redirect("/")

    # wrong password
    if not check_password_hash(user[3], password):
        flash("Incorrect password. Try again.", "error")
        return redirect("/")

    # login success
    session["user_id"] = user[0]
    session["user_name"] = user[1]

    flash("Login successful!", "success")

    return redirect("/dashboard")


# ---------------- LOGOUT ----------------
@app.route("/logout")
def logout():

    session.clear()

    flash("Logged out successfully.", "success")

    return redirect("/")


# ---------------- GENERATE TEXT ----------------
@app.route("/generate/text", methods=["POST"])
def gen_text():

    if "user_id" not in session:
        flash("Please login first.", "error")
        return redirect("/")

    data = request.form
    return "Text PPT Generated (logic here)"


# ---------------- GENERATE RESEARCH ----------------
#@app.route("/generate/research", methods=["POST"])
#def gen_research():
#
 #   if "user_id" not in session:
#        flash("Please login first.", "error")
#        return redirect("/")
#
 #   return "Research PPT Generated"


# ---------------- GENERATE PDF ----------------
@app.route("/generate/pdf", methods=["POST"])
def gen_pdf():

    if "user_id" not in session:
        flash("Please login first.", "error")
        return redirect("/")

    pdf = request.files["pdf"]

    return "PDF PPT Generated"


# ---------------- RUN SERVER ----------------
if __name__ == "__main__":
    print("Starting Flask Server...")
    app.run(debug=True)
