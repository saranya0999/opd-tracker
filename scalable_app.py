from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

DATA_DIR = "data"
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

HEADERS = ["Date", "Session", "New Case", "Old Case", "Name", "Age", "Gender", "Diagnosis", "User Name"]

@app.route("/", methods=["GET", "POST"])
def form():
    if request.method == "POST":
        date = request.form["date"]
        session = request.form["session"]
        new_case = request.form["new_case"]
        old_case = request.form["old_case"]
        name = request.form["name"]
        age = request.form["age"]
        gender = request.form["gender"]
        diagnosis = request.form["diagnosis"]
        username = request.form["username"].strip().lower().replace(" ", "")

        filename = f"{DATA_DIR}/julycensus_{username}.xlsx"

        entry = pd.DataFrame([[date, session, new_case, old_case, name, age, gender, diagnosis, username]], columns=HEADERS)

        if os.path.exists(filename):
            df = pd.read_excel(filename)
            df = pd.concat([df, entry], ignore_index=True)
        else:
            df = entry

        df.to_excel(filename, index=False)
        return redirect(url_for("view", username=username))

    return render_template("form_cases_sessions.html")

@app.route("/view", methods=["GET", "POST"])
def view():
    if request.method == "POST":
        username = request.form["username"].strip().lower().replace(" ", "")
        filepath = f"{DATA_DIR}/julycensus_{username}.xlsx"
        if os.path.exists(filepath):
            df = pd.read_excel(filepath)
            return render_template("view_entries.html", tables=[df.to_html(classes='data')], titles=df.columns.values, username=username)
        else:
            return f"No data found for user '{username}'"
    return render_template("view_input.html")