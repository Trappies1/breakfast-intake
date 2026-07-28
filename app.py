import os
import json
import base64
import sqlite3
import datetime
from io import BytesIO

import requests
from flask import Flask, request, render_template, send_file, redirect, url_for, flash
import openpyxl

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me")

# ---------------------------------------------------------------------------
# Config - set these as Environment Variables in Render, not in this file.
# ---------------------------------------------------------------------------
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
STAFF_PASSCODE = os.environ.get("STAFF_PASSCODE", "")      # optional PIN for the /orders page
DB_PATH = os.environ.get("DB_PATH", "orders.db")

MODEL = "claude-sonnet-4-6"

# The exact La Provence breakfast-order spreadsheet columns, in order.
COLUMNS = [
    "ID", "Start time", "Completion time", "Email", "Name", "Name2", "Date",
    "Room number", "If at 43 AM", "Specify", "I will enjoy my breakfast ...",
    "Special packing instructions", "Time breakfast must be served",
    "Can we include some fruit juice?", "Yoghurt - plain, fruit or no yoghurt",
    "Would you like to add a piece of fruit?",
    "Please indicate if there is any fruit, yoghurt or juice flavor that should not be included",
    "Also included on our breakfast is your choice of muesli, Corn Flakes, Pronutro, Rice Crispies or Bran Flakes (All Bran) . Please choose one:",
    "Bread", "Please choose one of the following",
    "Two eggs prepared the way you like it:", "Choose your first meat option",
    "Choose your second meat option", "Bread for the sandwich",
    "Breakfast burger option", "How must the eggs be done for the breakfast burger?",
    "Dagwood ", "Croissant, omelette, toasted sandwich, wrap filling 1",
    "Croissant, omelette, toasted sandwich, wrap filling 2",
    "Choose your vegetable/starch", "Special requirements:", "Add a lunch pack",
    "Bread option\n", "Filling", "Specify other", "Extra\n", "Spercify other",
    "Extra 2", "Specify other2", "Extra 3", "Specify other3", "Cold Drink",
    "Specify other4", "Last modified time",
]

EXTRACTION_PROMPT = """You are reading a photographed paper "La Provence Guesthouse - Breakfast Menu" \
order form. It is a checkbox form, hand-marked (circled or ticked) by a guest or staff \
member, sometimes with handwritten corrections in the margins.

Read ONLY the clearly marked selections. Return a single JSON object (no other text, \
no markdown fences) with exactly these keys. Use null for anything not marked or not \
applicable. Use the string "Please confirm" for anything marked in a way that is \
genuinely ambiguous (e.g. two contradictory boxes ticked in the same question, or a \
mark that could apply to either of two options).

Keys:
- "guest_name": guest's name if visible on the form or its label/sticker, else null
- "room_number": room number as a string
- "service_style": one of "In the dining room", "As a takeaway - to be fetched in the dining room",
  "Yoghurt, cereal & fruit in dining room + warm breakfast take-away",
  "Warm breakfast in dining room + yoghurt, cereal & fruit take-away", or "Please confirm"
- "time": the ticked time (e.g. "6:45", "7:00", "7:30", "8:00", "8:30")
- "fruit_juice": "Yes, fruit juice;" or "No"
- "yoghurt": "Plain", "Fruit Flavoured", or "No yoghurt"
- "add_fruit": "Yes" or "No"
- "cereal": one of "Muesli without milk - I will enjoy it with my yoghurt", "Muesli with milk",
  "Oats with milk", "Corn Flakes with milk", "Pronutro with milk", "Bran Flakes with milk", "None"
- "bread": one of "Plain white", "Toasted white", "Plain brown", "Toasted brown", "No bread"
- "option": which of the numbered breakfast options (1-7) was completed - one of
  "Create your own breakfast", "Dagwood", "Wrap", "Breakfast burger", "Toasted Sandwich",
  "Croissant", "Omelette"
- "eggs": egg style if the "Create your own breakfast" or burger option was chosen
  (e.g. "Soft fried", "Medium fried", "Hard fried", "Scrambled", "No eggs")
- "meat_1": first protein choice, if applicable
- "meat_2": second protein choice, if applicable
- "veg_starch": vegetable/starch or side choice (also used as the Dagwood/Wrap side)
- "dagwood_bread": "White" or "Brown", only if option is "Dagwood"
- "wrap_filling_1": first wrap/croissant/toasted-sandwich filling, if applicable
- "wrap_filling_2": second filling, if applicable
- "lunch_pack": true if a lunch pack section is filled in on this form, else false
- "lunch_pack_bread": lunch pack bread choice, if applicable
- "lunch_pack_filling": semicolon-joined lunch pack fillings, e.g. "Cheese;Mince;", if applicable
- "lunch_pack_extras": array of any lunch pack extra items written or ticked, if applicable
- "lunch_pack_drink": lunch pack drink choice, if applicable
- "special_requirements": a short plain-English note capturing: any dietary flag written on
  the form's label/sticker (e.g. "No onion"), any handwritten margin correction or note,
  and a clear flag for anything you marked "Please confirm" above, explaining what needs
  checking. Use null only if there is genuinely nothing to flag.

Be conservative: if a form is blank, incomplete, or contradictory in a section, say so in
special_requirements rather than guessing a specific answer.
"""


TEXT_EXTRACTION_PROMPT = """A staff member has typed a breakfast order in plain, informal language \
(often shorthand, sometimes with typos) - the same way they'd describe it out loud. It may be as \
short as a name and a couple of items, or as detailed as a full form. Some shorthand you may see, \
matching this guesthouse's own conventions:

- "standard pork" means: dining room service, fruit juice yes, fruit-flavoured yoghurt, a piece of \
  fruit, muesli, bread to confirm, "Create your own breakfast" option, eggs to confirm, Bacon as the \
  first meat, Mini Cheese Griller as the second meat, Chips as the veg/starch - unless the message \
  overrides specific parts of this (e.g. a different time, "no-pork" substitutes the meats, "sit \
  down" means dining room, "take away" means takeaway).
- "standard no-pork" is the same as above but with Chicken strips and Boerewors as the two meats
  instead of Bacon and any pork item.

Read the message below and return a single JSON object (no other text, no markdown fences) with \
exactly these keys - use null for anything not mentioned or not applicable, and the string \
"Please confirm" only for something genuinely ambiguous that truly needs a human to check (not \
simply because it wasn't mentioned - unmentioned fields should just be null).

Keys:
- "guest_name": guest's name
- "room_number": room number as a string, if given
- "service_style": one of "In the dining room", "As a takeaway - to be fetched in the dining room",
  "Yoghurt, cereal & fruit in dining room + warm breakfast take-away",
  "Warm breakfast in dining room + yoghurt, cereal & fruit take-away", or null
- "time": the breakfast time, if given
- "fruit_juice": "Yes, fruit juice;" or "No"
- "yoghurt": "Plain", "Fruit Flavoured", or "No yoghurt"
- "add_fruit": "Yes" or "No"
- "cereal": one of "Muesli without milk - I will enjoy it with my yoghurt", "Muesli with milk",
  "Oats with milk", "Corn Flakes with milk", "Pronutro with milk", "Bran Flakes with milk", "None"
- "bread": one of "Plain white", "Toasted white", "Plain brown", "Toasted brown", "No bread",
  "Please confirm"
- "option": "Create your own breakfast", "Dagwood", "Wrap", "Breakfast burger",
  "Toasted Sandwich", "Croissant", or "Omelette"
- "eggs": egg style, if applicable
- "meat_1": first protein choice, if applicable
- "meat_2": second protein choice, if applicable
- "veg_starch": vegetable/starch or side choice
- "dagwood_bread": "White" or "Brown", only if option is "Dagwood"
- "wrap_filling_1": first wrap/croissant/toasted-sandwich filling, if applicable
- "wrap_filling_2": second filling, if applicable
- "lunch_pack": true if a lunch pack is mentioned, else false
- "lunch_pack_bread": lunch pack bread choice, if applicable
- "lunch_pack_filling": semicolon-joined lunch pack fillings, e.g. "Steak;Tomato salad;", if applicable
- "lunch_pack_extras": array of any lunch pack extra items mentioned, if applicable
- "lunch_pack_drink": lunch pack drink choice, if applicable
- "special_requirements": a short plain-English note for any dietary flag, allergy, or anything you
  marked "Please confirm" above, explaining what needs checking. Use null if there's nothing to flag.

Message to read:
---
{message}
---
"""


def call_claude_text(message_text):
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": MODEL,
            "max_tokens": 1200,
            "messages": [
                {"role": "user", "content": TEXT_EXTRACTION_PROMPT.format(message=message_text)}
            ],
        },
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json()["content"]
    text = "".join(block.get("text", "") for block in content if block.get("type") == "text")
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            submitted_at TEXT,
            data_json TEXT
        )"""
    )
    conn.commit()
    conn.close()


def call_claude_vision(image_bytes, media_type):
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": MODEL,
            "max_tokens": 1500,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": b64,
                            },
                        },
                        {"type": "text", "text": EXTRACTION_PROMPT},
                    ],
                }
            ],
        },
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json()["content"]
    text = "".join(block.get("text", "") for block in content if block.get("type") == "text")
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)


@app.route("/", methods=["GET"])
def upload_page():
    return render_template("upload.html")


@app.route("/submit", methods=["POST"])
def submit():
    photo = request.files.get("photo")
    if not photo or photo.filename == "":
        flash("Please choose or take a photo first.")
        return redirect(url_for("upload_page"))

    image_bytes = photo.read()
    media_type = photo.mimetype or "image/jpeg"

    try:
        extracted = call_claude_vision(image_bytes, media_type)
    except Exception as exc:
        flash(f"Something went wrong reading that photo: {exc}. Please try again.")
        return redirect(url_for("upload_page"))

    init_db()
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO orders (submitted_at, data_json) VALUES (?, ?)",
        (datetime.datetime.now().isoformat(), json.dumps(extracted)),
    )
    conn.commit()
    conn.close()

    return render_template("success.html", data=extracted)


@app.route("/submit_text", methods=["POST"])
def submit_text():
    message = (request.form.get("message") or "").strip()
    if not message:
        flash("Please type the order details first.")
        return redirect(url_for("upload_page"))

    try:
        extracted = call_claude_text(message)
    except Exception as exc:
        flash(f"Something went wrong reading that order: {exc}. Please try again.")
        return redirect(url_for("upload_page"))

    init_db()
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO orders (submitted_at, data_json) VALUES (?, ?)",
        (datetime.datetime.now().isoformat(), json.dumps(extracted)),
    )
    conn.commit()
    conn.close()

    return render_template("success.html", data=extracted)


@app.route("/orders", methods=["GET", "POST"])
def orders():
    if STAFF_PASSCODE:
        if request.method == "POST":
            if request.form.get("passcode") != STAFF_PASSCODE:
                flash("Incorrect passcode.")
                return render_template("passcode.html")
        else:
            if request.args.get("passcode") != STAFF_PASSCODE:
                return render_template("passcode.html")

    init_db()
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT id, submitted_at, data_json FROM orders ORDER BY id DESC"
    ).fetchall()
    conn.close()

    parsed = [(r[0], r[1], json.loads(r[2])) for r in rows]
    return render_template("orders.html", orders=parsed, passcode=STAFF_PASSCODE)


@app.route("/download")
def download():
    if STAFF_PASSCODE and request.args.get("passcode") != STAFF_PASSCODE:
        return "Not authorised", 403

    init_db()
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT data_json FROM orders ORDER BY id ASC").fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(COLUMNS)

    for i, (data_json,) in enumerate(rows, start=1):
        d = json.loads(data_json)
        row = {c: None for c in COLUMNS}
        row["ID"] = 1000 + i
        row["Email"] = "anonymous"
        row["Name2"] = d.get("guest_name")
        row["Room number"] = d.get("room_number")
        row["I will enjoy my breakfast ..."] = d.get("service_style")
        row["Time breakfast must be served"] = d.get("time")
        row["Can we include some fruit juice?"] = d.get("fruit_juice")
        row["Yoghurt - plain, fruit or no yoghurt"] = d.get("yoghurt")
        row["Would you like to add a piece of fruit?"] = d.get("add_fruit")
        row["Also included on our breakfast is your choice of muesli, Corn Flakes, Pronutro, Rice Crispies or Bran Flakes (All Bran) . Please choose one:"] = d.get("cereal")
        row["Bread"] = d.get("bread")
        row["Please choose one of the following"] = d.get("option")
        row["Two eggs prepared the way you like it:"] = d.get("eggs")
        row["Choose your first meat option"] = d.get("meat_1")
        row["Choose your second meat option"] = d.get("meat_2")
        row["Dagwood "] = d.get("dagwood_bread")
        row["Croissant, omelette, toasted sandwich, wrap filling 1"] = d.get("wrap_filling_1")
        row["Croissant, omelette, toasted sandwich, wrap filling 2"] = d.get("wrap_filling_2")
        row["Choose your vegetable/starch"] = d.get("veg_starch")
        row["Special requirements:"] = d.get("special_requirements")
        row["Add a lunch pack"] = "Yes" if d.get("lunch_pack") else "No"
        row["Bread option\n"] = d.get("lunch_pack_bread")
        row["Filling"] = d.get("lunch_pack_filling")
        row["Cold Drink"] = d.get("lunch_pack_drink")

        extras = d.get("lunch_pack_extras") or []
        extra_headings = ["Extra\n", "Extra 2", "Extra 3"]
        specify_headings = ["Spercify other", "Specify other2", "Specify other3"]
        for extra_val, ext_head, spec_head in zip(extras, extra_headings, specify_headings):
            row[ext_head] = "Other"
            row[spec_head] = extra_val

        ws.append([row[c] for c in COLUMNS])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    today = datetime.date.today().isoformat()
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"La Provence Breakfast orders {today}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
