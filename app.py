from flask import Flask, render_template, request, redirect, url_for, send_file, session, abort
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from PIL import Image
import fitz  # PyMuPDF
import os
import io
import base64
import json

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "supersecretkey")

# 📁 Konfigurasi Folder
TEMP_DIR = os.path.join(os.path.dirname(__file__), "temp")
os.makedirs(TEMP_DIR, exist_ok=True)

# 🔐 Autentikasi Google API
SCOPES = ["https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_INFO = json.loads(os.environ.get("GOOGLE_SERVICE_ACCOUNT", "{}"))
credentials = service_account.Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)

# 📂 Folder Root Google Drive
ROOT_FOLDER_ID = os.environ.get("FOLDER_ID")

# 🔐 Password per folder
FOLDER_PASSWORDS = json.loads(os.environ.get("FOLDER_PASSWORDS", "{}"))

# ✅ Fungsi Pendukung

def get_files(folder_id):
    query = f"'{folder_id}' in parents and trashed = false"
    result = drive_service.files().list(q=query, fields="files(id, name, mimeType, parents)").execute()
    return result.get("files", [])

def find_keyword_position(pdf_path, keyword):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    text_instances = page.search_for(keyword, quads=True)
    if text_instances:
        rect = text_instances[-1]  # Ambil posisi terakhir
        return rect.x1, rect.y1
    return None

def add_signature_to_pdf(pdf_path, signature_data, keyword):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    x, y = find_keyword_position(pdf_path, keyword) or (50, 50)
    img_bytes = base64.b64decode(signature_data.split(",")[1])
    img_path = os.path.join(TEMP_DIR, "signature.png")
    with open(img_path, "wb") as f:
        f.write(img_bytes)
    img = Image.open(img_path)
    img_width = 100
    img_height = int(img.height * (img_width / img.width))
    img.save(img_path)

    img_rect = fitz.Rect(x, y, x + img_width, y + img_height)
    page.insert_image(img_rect, filename=img_path)

    output_path = pdf_path.replace(".pdf", "_signed.pdf")
    doc.save(output_path)
    doc.close()
    return output_path

def move_file(file_id, folder_id):
    file = drive_service.files().get(fileId=file_id, fields="parents").execute()
    previous_parents = ",".join(file.get("parents"))
    drive_service.files().update(fileId=file_id, addParents=folder_id, removeParents=previous_parents, fields="id, parents").execute()

# ✅ Routing Utama

@app.route("/")
def index():
    folder_groups = {
        "Pengajuan Awal": ["01 - Pengajuan Awal"],
        "Rabat": ["02A - SPV HRGA", "03A - SPV", "03B - Manager", "03C - General"],
        "PRS": ["02B - PAMO", "04A - SPV", "04B - Manager", "04C - General"],
        "Final": ["05 - Final"]
    }

    result = drive_service.files().list(
        q=f"'{ROOT_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        fields="files(id, name)"
    ).execute()

    folders = {f['name']: f['id'] for f in result.get("files", [])}
    group_data = {}

    for group, folder_names in folder_groups.items():
        group_data[group] = []
        for name in folder_names:
            folder_id = folders.get(name)
            if folder_id:
                files = get_files(folder_id)
                group_data[group].append({"name": name, "id": folder_id, "count": len(files)})

    return render_template("index.html", group_data=group_data)

@app.route("/folder/<folder_name>", methods=["GET", "POST"])
def view_folder(folder_name):
    if request.method == "POST":
        password = request.form.get("password")
        if FOLDER_PASSWORDS.get(folder_name) == password:
            session[folder_name] = True
            return redirect(url_for("view_folder", folder_name=folder_name))
        return "🔒 Password salah.", 403

    if not session.get(folder_name):
        return render_template("login.html", folder_name=folder_name)

    result = drive_service.files().list(
        q=f"'{ROOT_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        fields="files(id, name)"
    ).execute()
    folder_id = next((f['id'] for f in result['files'] if f['name'] == folder_name), None)
    if not folder_id:
        return "Folder tidak ditemukan", 404

    files = get_files(folder_id)
    return render_template("folder.html", files=files, folder_name=folder_name)

@app.route("/preview/<file_id>")
def preview_file(file_id):
    request_file = drive_service.files().get_media(fileId=file_id)
    file_metadata = drive_service.files().get(fileId=file_id).execute()
    filename = file_metadata['name']
    temp_path = os.path.join(TEMP_DIR, filename)

    with io.FileIO(temp_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request_file)
        done = False
        while not done:
            _, done = downloader.next_chunk()

    if not filename.lower().endswith(".pdf"):
        return "Hanya file PDF yang bisa dipreview", 400

    return send_file(temp_path, mimetype="application/pdf")

@app.route("/save_signature", methods=["POST"])
def save_signature():
    file_id = request.form['file_id']
    folder_name = request.form['folder']
    signature_data = request.form['signature']
    file_metadata = drive_service.files().get(fileId=file_id).execute()
    filename = file_metadata['name']

    temp_pdf_path = os.path.join(TEMP_DIR, filename)
    request_file = drive_service.files().get_media(fileId=file_id)
    with io.FileIO(temp_pdf_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request_file)
        done = False
        while not done:
            _, done = downloader.next_chunk()

    keyword = folder_name.split()[-1]
    signed_path = add_signature_to_pdf(temp_pdf_path, signature_data, keyword)

    media = MediaFileUpload(signed_path, mimetype="application/pdf", resumable=True)
    drive_service.files().update(fileId=file_id, media_body=media).execute()

    # 🚚 Pindah ke folder berikutnya kalau perlu
    folder_mapping = {
        "01 - Pengajuan Awal": {"SR": "02A - SPV HRGA", "MR": "02A - SPV HRGA", "GR": "02A - SPV HRGA",
                                "SP": "02B - PAMO", "MP": "02B - PAMO", "GP": "02B - PAMO"},
        "02A - SPV HRGA": {"SR": "03A - SPV", "MR": "03B - Manager", "GR": "03C - General"},
        "02B - PAMO": {"SP": "04A - SPV", "MP": "04B - Manager", "GP": "04C - General"},
        "03A - SPV": {}, "03B - Manager": {}, "03C - General": {},
        "04A - SPV": {}, "04B - Manager": {}, "04C - General": {}
    }

    prefix = filename.split()[0][:2].upper()
    target_folder_name = folder_mapping.get(folder_name, {}).get(prefix)
    if target_folder_name:
        result = drive_service.files().list(
            q=f"'{ROOT_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            fields="files(id, name)"
        ).execute()
        target_id = next((f['id'] for f in result['files'] if f['name'] == target_folder_name), None)
        if target_id:
            move_file(file_id, target_id)

    return redirect(url_for("index"))

# ✅ Jalankan aplikasi di Railway
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
