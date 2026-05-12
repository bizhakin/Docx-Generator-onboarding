import os
import io
import re
import json
import uuid
from datetime import date, datetime
from flask import Flask, render_template, request, jsonify, send_file, abort
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "contract_templates")
CONTRACTS_CONFIG = os.path.join(BASE_DIR, "contracts.json")
GENERATED_DIR = os.path.join(BASE_DIR, "generated")
os.makedirs(GENERATED_DIR, exist_ok=True)

app = Flask(__name__)


def load_contracts():
    with open(CONTRACTS_CONFIG, "r") as f:
        return json.load(f)["contracts"]


def replace_in_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        placeholder = "{{" + key + "}}"
        if placeholder not in paragraph.text:
            continue
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, str(value))
        if placeholder in paragraph.text:
            full_text = "".join(r.text for r in paragraph.runs)
            if placeholder in full_text:
                new_text = full_text.replace(placeholder, str(value))
                for run in paragraph.runs:
                    run.text = ""
                if paragraph.runs:
                    paragraph.runs[0].text = new_text


def clear_remaining_placeholders(paragraph):
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


def fill_template(template_path, field_values, clear_unfilled=True):
    doc = Document(template_path)
    today = date.today().strftime("%B %d, %Y")
    field_values = dict(field_values)
    field_values["DATE"] = today
    field_values["Date"] = today

    all_paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)

    for para in all_paragraphs:
        replace_in_paragraph(para, field_values)
        if clear_unfilled:
            clear_remaining_placeholders(para)

    return doc


def get_base_url():
    """Return the externally accessible base URL for sharing links."""
    domains = os.environ.get("REPLIT_DOMAINS", "")
    if domains:
        domain = domains.split(",")[0].strip()
        return f"https://{domain}"
    dev_domain = os.environ.get("REPL_SLUG") or ""
    replit_dev = os.environ.get("REPLIT_DEV_DOMAIN", "")
    if replit_dev:
        return f"https://{replit_dev}"
    return request.host_url.rstrip("/")


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
        abort(500, description=f"Template file not found.")

    # Handle deliverables count — blank out unused slots
    deliverables_count = int(data.pop("deliverables_count", 5))
    for i in range(1, 6):
        for field in ["Service", "Quantity", "Turnaround", "Revisions"]:
            key = f"{field}{i}"
            if i > deliverables_count:
                data[key] = ""

    # Save job metadata for the signing workflow
    token = str(uuid.uuid4())
    job_dir = os.path.join(GENERATED_DIR, token)
    os.makedirs(job_dir, exist_ok=True)

    # Extract client signature field IDs so we can keep them blank in provider copy
    client_sig_fields = [f["id"] for f in contract.get("client_signature_fields", [])]

    # Provider copy: fill everything, leave client signature fields blank
    provider_data = dict(data)
    for fid in client_sig_fields:
        provider_data[fid] = ""

    provider_doc = fill_template(template_path, provider_data)
    provider_doc.save(os.path.join(job_dir, "provider_copy.docx"))

    # Save metadata for client signing later
    meta = {
        "contract_id": contract_id,
        "contract_name": contract["name"],
        "form_data": data,
        "client_signature_fields": contract.get("client_signature_fields", []),
        "template": contract["template"],
        "created_at": datetime.utcnow().isoformat(),
    }
    with open(os.path.join(job_dir, "meta.json"), "w") as f:
        json.dump(meta, f)

    base = get_base_url()
    return jsonify({
        "success": True,
        "token": token,
        "download_url": f"/download/{token}",
        "sign_url": f"{base}/sign/{token}",
        "contract_name": contract["name"],
    })


@app.route("/download/<token>")
def download(token):
    token = re.sub(r"[^a-f0-9\-]", "", token)
    job_dir = os.path.join(GENERATED_DIR, token)
    path = os.path.join(job_dir, "provider_copy.docx")
    if not os.path.exists(path):
        abort(404)
    meta_path = os.path.join(job_dir, "meta.json")
    contract_id = "contract"
    if os.path.exists(meta_path):
        with open(meta_path) as f:
            meta = json.load(f)
        contract_id = meta.get("contract_id", "contract")
    filename = f"{contract_id}_{date.today().isoformat()}.docx"
    return send_file(
        path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.route("/sign/<token>")
def sign_page(token):
    token = re.sub(r"[^a-f0-9\-]", "", token)
    meta_path = os.path.join(GENERATED_DIR, token, "meta.json")
    if not os.path.exists(meta_path):
        abort(404)
    with open(meta_path) as f:
        meta = json.load(f)
    return render_template("sign.html", token=token, meta=meta)


@app.route("/sign/<token>/submit", methods=["POST"])
def sign_submit(token):
    token = re.sub(r"[^a-f0-9\-]", "", token)
    meta_path = os.path.join(GENERATED_DIR, token, "meta.json")
    if not os.path.exists(meta_path):
        abort(404)
    with open(meta_path) as f:
        meta = json.load(f)

    template_path = os.path.join(TEMPLATES_DIR, meta["template"])
    if not os.path.exists(template_path):
        abort(500)

    # Merge original form data + client signature fields
    all_data = dict(meta["form_data"])
    for sig_field in meta.get("client_signature_fields", []):
        fid = sig_field["id"]
        all_data[fid] = request.form.get(fid, "")

    doc = fill_template(template_path, all_data)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    contract_id = meta.get("contract_id", "contract")
    filename = f"{contract_id}_signed_{date.today().isoformat()}.docx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
