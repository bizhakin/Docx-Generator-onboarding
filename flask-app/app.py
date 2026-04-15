import os
import io
import re
import json
from datetime import date
from flask import Flask, render_template, request, jsonify, send_file, abort
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "contract_templates")
CONTRACTS_CONFIG = os.path.join(BASE_DIR, "contracts.json")

app = Flask(__name__)


def load_contracts():
    with open(CONTRACTS_CONFIG, "r") as f:
        return json.load(f)["contracts"]


def replace_in_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        placeholder = "{{" + key + "}}"
        if placeholder not in paragraph.text:
            continue
        # Try replacing run by run first (preserves formatting)
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, str(value))
        # If placeholder spans multiple runs, collapse and replace
        if placeholder in paragraph.text:
            full_text = "".join(r.text for r in paragraph.runs)
            if placeholder in full_text:
                new_text = full_text.replace(placeholder, str(value))
                for run in paragraph.runs:
                    run.text = ""
                if paragraph.runs:
                    paragraph.runs[0].text = new_text


def clear_remaining_placeholders(paragraph):
    """Remove any {{PLACEHOLDER}} that wasn't filled."""
    if "{{" not in paragraph.text:
        return
    for run in paragraph.runs:
        run.text = re.sub(r"\{\{[^}]+\}\}", "", run.text)
    if "{{" in paragraph.text:
        full_text = "".join(r.text for r in paragraph.runs)
        cleaned = re.sub(r"\{\{[^}]+\}\}", "", full_text)
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = cleaned


def fill_template(template_path, field_values):
    doc = Document(template_path)
    today = date.today().strftime("%B %d, %Y")
    field_values["DATE"] = today
    field_values["Date"] = today

    all_paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)

    for para in all_paragraphs:
        replace_in_paragraph(para, field_values)
        clear_remaining_placeholders(para)

    return doc


@app.route("/")
def index():
    contracts = load_contracts()
    return render_template("index.html", contracts=contracts)


@app.route("/fields/<contract_id>")
def get_contract_fields(contract_id):
    contracts = load_contracts()
    contract = next((c for c in contracts if c["id"] == contract_id), None)
    if not contract:
        abort(404)
    return jsonify(contract)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.form.to_dict()
    contract_id = data.pop("contract_id", None)
    if not contract_id:
        abort(400)

    contracts = load_contracts()
    contract = next((c for c in contracts if c["id"] == contract_id), None)
    if not contract:
        abort(404)

    template_path = os.path.join(TEMPLATES_DIR, contract["template"])
    if not os.path.exists(template_path):
        abort(500, description=f"Template file '{contract['template']}' not found.")

    # Handle deliverables count — blank out unused slots
    deliverables_count = int(data.pop("deliverables_count", 5))
    for i in range(1, 6):
        for field in ["Service", "Quantity", "Turnaround", "Revisions"]:
            key = f"{field}{i}"
            if i > deliverables_count:
                data[key] = ""

    doc = fill_template(template_path, data)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    filename = f"{contract_id}_{date.today().isoformat()}.docx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
