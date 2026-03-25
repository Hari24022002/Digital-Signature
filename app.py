from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory, jsonify
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from io import BytesIO
import qrcode
from PIL import Image
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer
from datetime import datetime
import os
import json
import logging
import uuid
import fitz
import pythoncom
import win32com.client

app = Flask(__name__)
app.secret_key = 'your_secret_key'

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'upload')
DOCS_FILE = 'documents.json'
HISTORY_FILE = 'history.json'
USERS_FILE = 'users.json'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
logging.basicConfig(level=logging.INFO)
ALLOWED_EXTENSIONS = {'.pdf', '.doc', '.docx'}

for path in [DOCS_FILE, HISTORY_FILE, USERS_FILE]:
    if not os.path.exists(path):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump([], f)
        logging.info(f"Initialized empty {path}")

def save_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4)
        f.flush()

def load_json(path):
    if not os.path.exists(path):
        save_json(path, [])
        return []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError:
        save_json(path, [])
        return []

def log_history(action, doc_id, username, pbno=None, remarks=None):
    history = load_json(HISTORY_FILE)
    entry = {
        "timestamp": datetime.now().isoformat(),
        "action": action,
        "document_id": doc_id,
        "username": username,
        "pbno": pbno,
        "remarks": remarks,
        "stamp": session.get("stamp", "")
    }
    history.append(entry)
    save_json(HISTORY_FILE, history)

def deduplicate_documents(doc_list):
    seen = set()
    unique_docs = []
    for doc in doc_list:
        if doc.get("id") not in seen:
            seen.add(doc.get("id"))
            unique_docs.append(doc)
    return unique_docs

def find_label_coordinates(pdf_path, label):
    label = label.lower()
    for page_layout in extract_pages(pdf_path):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                text = element.get_text().strip()
                if label in text.lower():
                    return float(element.x0), float(element.y0)
    return None, None

def allowed_file(filename):
    return '.' in filename and os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS

def convert_doc_to_pdf(doc_paths, pdf_paths):
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for doc_path, pdf_path in zip(doc_paths, pdf_paths):
            try:
                abs_doc_path = os.path.abspath(doc_path)
                abs_pdf_path = os.path.abspath(pdf_path)
                doc = word.Documents.Open(abs_doc_path)
                doc.SaveAs(abs_pdf_path, FileFormat=17)  # 17 = PDF format
                doc.Close(False)
            except Exception as e:
                logging.error(f"Failed to convert {doc_path}: {e}")
                continue
        word.Quit()
    except Exception as e:
        logging.error(f"COM conversion failed: {e}")
        raise RuntimeError(f"COM conversion failed: {e}")
    finally:
        pythoncom.CoUninitialize()

def draw_signature_block(canvas_obj, x, y, user_info, qr_img):
    canvas_obj.setFont("Helvetica", 8)
    qr_size = 60
    canvas_obj.drawInlineImage(qr_img, x, y, width=qr_size, height=qr_size)
    text_x = x + qr_size + 10
    text_y = y + qr_size - 10
    color = (0, 0, 0)
    if user_info.get('role', '').lower() == 'certifier':
        if user_info.get('action', '').lower() == 'approved':
            color = (0, 128, 0)
        elif user_info.get('action', '').lower() == 'reverted':
            color = (255, 0, 0)
    canvas_obj.setFillColorRGB(color[0]/255.0, color[1]/255.0, color[2]/255.0)
    lines = [
        f"Name: {user_info['username']}",
        f"Role: {user_info['role']}",
        f"PB No: {user_info['pbno']}",
        f"Action: {user_info.get('action', '')}",
        f"Timestamp: {user_info['timestamp']}"
    ]
    if user_info.get('stamp'):
        lines.insert(3, f"Stamp: {user_info['stamp']}")
    for line in lines:
        canvas_obj.drawString(text_x, text_y, line)
        text_y -= 11
    canvas_obj.setFont("Helvetica-Oblique", 8)
    canvas_obj.drawCentredString(x + qr_size / 2, y - 12, "Digitally Signed")
    canvas_obj.setFillColorRGB(0, 0, 0)

def insert_signature_to_pdf(pdf_path, output_path, user_info, role_context):
    label = "Tested By" if role_context.lower() == "prepared" else "Certified By"
    x_label, y_label = find_label_coordinates(pdf_path, label)
    if x_label is None:
        raise Exception(f"Label '{label}' not found in PDF.")
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    lines = [
        f"Name: {user_info['username']}",
        f"Role: {user_info['role']}",
        f"PB No: {user_info['pbno']}",
        f"Action: {user_info.get('action', '')}",
        f"Timestamp: {user_info['timestamp']}"
    ]
    if user_info.get('stamp'):
        lines.insert(3, f"Stamp: {user_info['stamp']}")
    qr_data = "\n".join(lines)
    qr = qrcode.make(qr_data)
    qr_buffer = BytesIO()
    qr.save(qr_buffer, format='PNG')
    qr_buffer.seek(0)
    qr_img = Image.open(qr_buffer)
    for i in range(len(reader.pages) - 1):
        writer.add_page(reader.pages[i])
    last_page = reader.pages[-1]
    packet = BytesIO()
    width = float(last_page.mediabox.width)
    height = float(last_page.mediabox.height)
    c = canvas.Canvas(packet, pagesize=(width, height))
    SIGNATURE_BLOCK_WIDTH = 180
    base_y = y_label - 70
    base_x = max(10, min(x_label - 60, width - SIGNATURE_BLOCK_WIDTH - 10))
    draw_signature_block(c, base_x, base_y, user_info, qr_img)
    c.save()
    packet.seek(0)
    overlay = PdfReader(packet)
    last_page.merge_page(overlay.pages[0])
    writer.add_page(last_page)
    with open(output_path, 'wb') as f:
        writer.write(f)

@app.route('/')
def home():
    history = load_json(HISTORY_FILE)
    documents = load_json(DOCS_FILE)
    return render_template('home.html', documents=documents, history=history)

@app.route('/login', methods=['POST'])
def login():
    pbno = request.form['pbno']
    password = request.form['password']
    users = load_json(USERS_FILE)
    user = next((u for u in users if u['pbno'] == pbno and u['password'] == password), None)
    if user:
        session['username'] = user['username']
        session['pbno'] = user['pbno']
        session['role'] = user['role']
        session['stamp'] = user.get('stamp', '')
        return redirect(url_for('dashboard'))
    flash("Invalid credentials", 'danger')
    return redirect(url_for('home'))

@app.route('/register', methods=['POST'])
def register():
    username = request.form['username']
    pbno = request.form['pbno']
    password = request.form['password']
    role = request.form['role']
    stamp = request.form.get('stamp', '')
    users = load_json(USERS_FILE)
    if not all([username, pbno, password, role]):
        flash("All fields except stamp are required.", "register:danger")
        return redirect(url_for('home', tab='register'))
    if any(u['username'] == username for u in users):
        flash("Username already exists.", "register:danger")
        return redirect(url_for('home', tab='register'))
    users.append({
        "username": username,
        "pbno": pbno,
        "password": password,
        "role": role,
        "stamp": stamp
    })
    save_json(USERS_FILE, users)
    flash("Registration successful. Please login.", "register:success")
    return redirect(url_for('home', tab='register'))

@app.route('/dashboard')
def dashboard():
    if 'username' not in session:
        return redirect(url_for('home'))
    documents = load_json(DOCS_FILE)
    role = session['role']
    username = session['username']
    if role.lower() == 'prepared':
        user_docs = [d for d in documents if d['uploaded_by'] == username]
        return render_template('dashboard_prepared.html', documents=user_docs, user=username)
    elif role.lower() == 'certifier':
        relevant_docs = [d for d in documents if d.get('status') in ['Pending','Approved','Rejected']]
        relevant_docs.sort(key=lambda x: x.get('timestamp_uploaded',''), reverse=True)
        return render_template('dashboard_certifier.html', documents=relevant_docs, user=username)
    elif role.lower() == 'manager':
        uploaded_docs = [d for d in documents if d['uploaded_by'] == username]
        pending_docs = [d for d in documents if d['status'] == 'Pending']
        actioned_docs = [d for d in documents if d['status'] in ['Approved', 'Rejected']]
        actioned_docs.sort(key=lambda d: d.get('timestamp_actioned',''), reverse=True)
        combined = deduplicate_documents(uploaded_docs + pending_docs + actioned_docs)
        return render_template('dashboard_manager.html', documents=combined, user=username)

@app.route('/upload', methods=['POST'])
def upload():
    if 'username' not in session:
        return redirect(url_for('home'))
    uploaded_files = request.files.getlist('doc')
    documents = load_json(DOCS_FILE)
    word_files = []
    word_output_paths = []
    pdf_files = []
    temp_paths = []
    doc_entries = []
    for file in uploaded_files:
        if file and allowed_file(file.filename):
            original_filename = file.filename
            name, ext = os.path.splitext(original_filename)
            ext = ext.lower()
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
            unique_filename = f"{timestamp}_{name}{ext}"
            temp_path = os.path.join(UPLOAD_FOLDER, unique_filename)
            file.save(temp_path)
            temp_paths.append(temp_path)
            pdf_filename = f"{timestamp}_{name}.pdf"
            pdf_path = os.path.join(UPLOAD_FOLDER, pdf_filename)
            if ext in ['.doc', '.docx']:
                word_files.append(temp_path)
                word_output_paths.append(pdf_path)
            else:
                pdf_files.append((temp_path, pdf_path))
            doc_entry = {
                "id": str(uuid.uuid4()),
                "original_filename": original_filename,
                "filename": os.path.basename(pdf_path),
                "signed_filename": "",
                "status": "Uploaded",
                "uploaded_by": session['username'],
                "pbno": session['pbno'],
                "certifier": "",
                "certifier_pbno": session['pbno'],
                "remarks": "",
                "timestamp_uploaded": datetime.now().isoformat(),
                "timestamp_actioned": ""
            }
            doc_entries.append(doc_entry)
    if word_files:
        try:
            convert_doc_to_pdf(word_files, word_output_paths)
            for temp_path in word_files:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        except Exception as e:
            flash(f"Failed to convert Word files to PDF: {str(e)}", "danger")
            for temp_path in word_files:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
    for temp_path, pdf_path in pdf_files:
        try:
            os.rename(temp_path, pdf_path)
        except Exception as e:
            flash(f"Failed to process PDF {os.path.basename(temp_path)}: {str(e)}", "danger")
            if os.path.exists(temp_path):
                os.remove(temp_path)
    for doc_entry in doc_entries:
        documents.append(doc_entry)
        log_history("Uploaded", doc_entry["id"], session['username'], session['pbno'])
    save_json(DOCS_FILE, documents)
    return redirect(url_for('dashboard'))

@app.route('/verify/<id>')
def verify(id):
    if 'username' not in session:
        return jsonify({'error': 'Unauthorized, please log in'}), 401
    documents = load_json(DOCS_FILE)
    doc = next((d for d in documents if str(d['id']) == id), None)
    if not doc or doc['status'] != 'Uploaded':
        return jsonify({'error': 'Document not found or not in uploadable state'}), 404
    input_path = os.path.join(UPLOAD_FOLDER, doc['filename'])
    signed_name = f"signed_{doc['filename']}"
    output_path = os.path.join(UPLOAD_FOLDER, signed_name)
    user_info = {
        "username": session['username'],
        "pbno": session['pbno'],
        "role": "Prepared",
        "action": "Verified",
        "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        "stamp": session.get('stamp', '')
    }
    try:
        insert_signature_to_pdf(input_path, output_path, user_info, "Prepared")
        doc['status'] = 'Verified'
        doc['signed_filename'] = signed_name
        save_json(DOCS_FILE, documents)
        log_history("Verified", doc["id"], session['username'], session['pbno'])
        return jsonify({'message': 'Document verified successfully'}), 200
    except Exception as e:
        if str(e).startswith("Label"):
            return jsonify({'error': "Could not find 'Tested By' or 'Certified By' label in the PDF. Please upload a valid file."}), 400
        return jsonify({'error': f"Verification failed: {str(e)}"}), 500

@app.route('/send/<id>')
def send(id):
    if 'username' not in session:
        return jsonify({'error': 'Unauthorized, please log in'}), 401
    documents = load_json(DOCS_FILE)
    doc = next((d for d in documents if str(d['id']) == id), None)
    if not doc:
        return jsonify({'error': 'Document not found'}), 404
    if doc['status'] != 'Verified':
        return jsonify({'error': 'Document not in verifiable state'}), 400
    if doc['uploaded_by'] != session['username']:
        return jsonify({'error': 'Unauthorized to send this document'}), 403
    doc['status'] = 'Pending'
    save_json(DOCS_FILE, documents)
    log_history("Sent to Certifier", doc["id"], session['username'], session['pbno'])
    return jsonify({'message': 'Document sent successfully'}), 200

@app.route('/action/<id>/<decision>', methods=['POST'])
def action(id, decision):
    if 'username' not in session:
        return jsonify({'error': 'Unauthorized, please log in'}), 401
    remarks = request.form.get('remarks', '')
    documents = load_json(DOCS_FILE)
    for doc in documents:
        if str(doc['id']) == id and doc['status'] == 'Pending':
            input_path = os.path.join(UPLOAD_FOLDER, doc.get('signed_filename') or doc.get('filename'))
            final_name = f"{decision.lower()}_{id}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, final_name)
            qr_action = "Reverted" if decision.lower() == "rejected" else decision
            user_info = {
                "username": session['username'],
                "pbno": session['pbno'],
                "role": "Certifier",
                "action": qr_action,
                "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "stamp": session.get('stamp', '')
            }
            try:
                insert_signature_to_pdf(input_path, output_path, user_info, "Certifier")
                doc.update({
                    'filename': final_name,
                    'signed_filename': final_name,
                    'status': decision,
                    'approver': session['username'],
                    'remarks': remarks,
                    'timestamp_actioned': user_info['timestamp']
                })
                save_json(DOCS_FILE, documents)
                log_history(f"{decision} by Certifier", doc["id"], session['username'], session['pbno'], remarks)
                return '', 204
            except Exception as e:
                return jsonify({'error': f"Action failed: {str(e)}"}), 500
    return jsonify({'error': 'Document not found or not in pending state'}), 404

@app.route('/delete/<id>', methods=['DELETE'])
def delete_document(id):
    if 'username' not in session:
        return jsonify({'error': 'Unauthorized, please log in'}), 401
    documents = load_json(DOCS_FILE)
    doc = next((d for d in documents if str(d['id']) == id), None)
    if not doc:
        return jsonify({'error': 'Document not found'}), 404
    if doc['uploaded_by'] != session['username']:
        return jsonify({'error': 'Unauthorized to delete this document'}), 403
    if doc['status'] not in ['Uploaded', 'Verified', 'Rejected']:
        return jsonify({'error': 'Cannot delete document in this status'}), 400
    try:
        file_path = os.path.join(UPLOAD_FOLDER, doc['filename'])
        if os.path.exists(file_path):
            os.remove(file_path)
        if doc['signed_filename']:
            signed_file_path = os.path.join(UPLOAD_FOLDER, doc['signed_filename'])
            if os.path.exists(signed_file_path):
                os.remove(signed_file_path)
        documents = [d for d in documents if str(d['id']) != id]
        save_json(DOCS_FILE, documents)
        history = load_json(HISTORY_FILE)
        history = [h for h in history if str(h['document_id']) != id]
        save_json(HISTORY_FILE, history)
        return jsonify({'message': 'Document deleted successfully'}), 200
    except Exception as e:
        return jsonify({'error': f'Deletion failed: {str(e)}'}), 500

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('home'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5057)
