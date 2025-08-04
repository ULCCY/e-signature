from flask import Flask, render_template, request, redirect, url_for, send_file, session, jsonify, abort
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from PIL import Image
from functools import wraps
from dotenv import load_dotenv
import fitz  # PyMuPDF
import os
import base64
import io
import re
from io import BytesIO
import pypandoc
from docx2pdf import convert as docx2pdf_convert
import shutil
import time
from datetime import datetime

# === INISIALISASI APP ===
app = Flask(__name__)
load_dotenv()
app.secret_key = os.getenv('SECRET_KEY')

# ✅ Inisialisasi status file yang ready
ready_files = {}

# === KONFIGURASI PATH ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, 'temp')
os.makedirs(TEMP_DIR, exist_ok=True)
print(f"📁 TEMP_DIR: {TEMP_DIR}")

# === KONFIGURASI GOOGLE DRIVE ===
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials.json')
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

# === KONFIGURASI FOLDER ===
FOLDERS = {
    "01": {"id": os.getenv("FOLDER_01"), "name": "01 - Vehicle Repair Request", "password": os.getenv("PASSWORD_01"), "group": "Vehicle Repair Request"},
    "02A": {"id": os.getenv("FOLDER_02A"), "name": "02A - HRGA Supervisor", "password": os.getenv("PASSWORD_02A"), "group": "Rabat"},
    "03A": {"id": os.getenv("FOLDER_03A"), "name": "03A - Finance Supervisor", "password": os.getenv("PASSWORD_03A"), "group": "Rabat"},
    "03B": {"id": os.getenv("FOLDER_03B"), "name": "03B - Finance Manager", "password": os.getenv("PASSWORD_03B"), "group": "Rabat"},
    "03C": {"id": os.getenv("FOLDER_03C"), "name": "03C - General Manager", "password": os.getenv("PASSWORD_03C"), "group": "Rabat"},
    # "01A": {"id": os.getenv("FOLDER_01A"), "name": "01A - RSM", "password": os.getenv("PASSWORD_01A"), "group": "PRS"},
    "02B": {"id": os.getenv("FOLDER_02B"), "name": "02B - PAMO", "password": os.getenv("PASSWORD_02B"), "group": "PRS"},
    "04A": {"id": os.getenv("FOLDER_04A"), "name": "04A - Finance Supervisor", "password": os.getenv("PASSWORD_04A"), "group": "PRS"},
    "04B": {"id": os.getenv("FOLDER_04B"), "name": "04B - Finance Manager", "password": os.getenv("PASSWORD_04B"), "group": "PRS"},
    "04C": {"id": os.getenv("FOLDER_04C"), "name": "04C - General Manager", "password": os.getenv("PASSWORD_04C"), "group": "PRS"},
    "05": {"id": os.getenv("FOLDER_05"), "name": "05 - Final", "password": os.getenv("PASSWORD_05"), "group": "Final"},
    }

# === UTILITAS ===
def require_folder_password(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        step = kwargs.get('step') or request.view_args.get('step') or request.args.get('folder') or request.form.get('folder')
        if not step:
            print("❌ Tidak ada step/folder_id terdeteksi!")
            abort(400)  # lebih baik abort daripada redirect ke URL rusak
        if step not in session:
            return redirect(url_for('access_file', folder_id=step, file_id=kwargs.get('file_id')))
        return func(*args, **kwargs)
    return wrapper

def extract_prefix(filename):
    match = re.search(r'\b(SR|MR|GR|SP|MP|GP)\b', filename.upper())
    return match.group(1) if match else ''

def find_keyword_position(pdf_path, keywords):
    doc = fitz.open(pdf_path)
    keywords = [k.lower() for k in keywords]
    last = None

    page = doc[0]  # hanya halaman pertama
    blocks = page.get_text("blocks")
    for block in blocks:
        text = block[4].lower()
        for keyword in keywords:
            # 🔥 PAKAI REGEX BOUNDARY (\b)
            if re.search(rf'\b{re.escape(keyword)}\b', text):
                print(f"🔍 Keyword ditemukan persis: {keyword} pada posisi {block[:4]}")
                if not last or block[1] > last[2]:
                    last = (0, block[0], block[1])
    doc.close()
    
    if not last:
        print(f"❌ Tidak ada keyword ditemukan dari: {keywords}")
    return last

def embed_signature(pdf_path, signature_data_url, output_path, keywords):
    header, encoded = signature_data_url.split(",", 1)
    image = Image.open(BytesIO(base64.b64decode(encoded)))
    img_byte_arr = BytesIO()
    image.save(img_byte_arr, format='PNG')
    img_data = img_byte_arr.getvalue()

    doc = fitz.open(pdf_path)
    found = find_keyword_position(pdf_path, keywords)
    if not found:
        print("⚠️ Keyword tidak ditemukan, tidak menambahkan tanda tangan.")
        doc.close()
        return False
    
    if found:
        page_num, x0, y0 = found
        scale = 1.5  # 150%
        sig_width = image.width * 0.3 * scale
        sig_height = image.height * 0.3 * scale

        # Geser ke atas 0.75 tinggi tanda tangan dari titik tengah keyword
        y_center = y0 - (sig_height * 0.75)
        y_top = y_center - sig_height / 2
        y_bottom = y_center + sig_height / 2

        # Geser sedikit ke kiri (0.2 dari lebar tanda tangan)
        x_left = x0 - (0.2 * sig_width)

        rect = fitz.Rect(
            x_left,
            y_top,
            x_left + sig_width,
            y_bottom
        )

        doc[page_num].insert_image(rect, stream=img_data)
    
    doc.save(output_path)
    doc.close()
    return True

def preprocess_drive_files():
    """Konversi semua file Word/ODT di folder 01 ke PDF sebelum Flask jalan."""
    folder_id = FOLDERS["01"]["id"]
    print(f"🔍 Mengecek file di folder 01 (ID: {folder_id})...")

    results = drive_service.files().list(
        q=f"'{folder_id}' in parents",
        fields="files(id, name, mimeType)"
    ).execute()

    files = results.get('files', [])
    for f in files:
        file_id = f['id']
        filename = f['name']
        ext = os.path.splitext(filename)[1].lower()

        if ext in ['.doc', '.docx', '.odt']:
            print(f"⚠️ {filename} bukan PDF → akan dikonversi sebelum server Flask jalan")
            ensure_pdf_on_drive(file_id)
        else:
            print(f"✅ {filename} sudah PDF")

def download_file_from_drive(file_id, dest_path):
    print(f"🔽 Mengunduh file {file_id} ke {dest_path}")
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    with open(dest_path, 'wb') as f:
        f.write(fh.getvalue())
    print(f"✅ File berhasil diunduh ke {dest_path}")

def ensure_pdf_on_drive(file_id):
    """
    Mengecek apakah file di Google Drive berformat .doc, .docx, atau .odt.
    Jika iya, konversi ke PDF lalu upload ulang untuk menggantikan file asli.
    """
    file_info = drive_service.files().get(fileId=file_id, fields="name, mimeType").execute()
    filename = file_info['name']
    ext = os.path.splitext(filename)[1].lower()

    # ✅ Hanya proses jika file Word/ODT
    if ext in ['.doc', '.docx', '.odt']:
        print(f"🔄 File {filename} (ID: {file_id}) bukan PDF, mengonversi...")

        # 1️⃣ Download file asli ke TEMP_DIR
        raw_path = os.path.join(TEMP_DIR, filename)
        request = drive_service.files().get_media(fileId=file_id)
        with open(raw_path, "wb") as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()

        # 2️⃣ Tentukan path PDF
        pdf_path = os.path.splitext(raw_path)[0] + ".pdf"

        try:
            # 3️⃣ Konversi sesuai tipe file
            docx2pdf_convert(raw_path, pdf_path)
            print(f"✅ {filename} berhasil dikonversi menjadi PDF (docx2pdf): {pdf_path}")
        except Exception as e:
            print(f"⚠️ docx2pdf gagal ({e}), fallback ke pypandoc")
            pypandoc.convert_file(raw_path, 'pdf', outputfile=pdf_path)
            print(f"✅ {filename} berhasil dikonversi menjadi PDF (pypandoc): {pdf_path}")

        # ✅ Bagian upload & rename HARUS keluar dari blok try/except
        media = MediaFileUpload(pdf_path, mimetype='application/pdf')
        drive_service.files().update(fileId=file_id, media_body=media).execute()
        print(f"📤 PDF berhasil diupload dan menggantikan file lama di Drive.")

        # 🔁 Rename file di Google Drive agar ekstensi jadi PDF
        new_name = os.path.splitext(filename)[0] + ".pdf"
        drive_service.files().update(fileId=file_id, body={"name": new_name}).execute()
        print(f"✏️ Nama file diubah menjadi: {new_name}")

        # 🧹 Bersihkan file sementara
        if os.path.exists(raw_path):
            os.remove(raw_path)

        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
                print(f"🗑 File PDF sementara dihapus: {pdf_path}")
            except PermissionError:
                print(f"⚠️ Tidak bisa hapus {pdf_path}, file sedang digunakan.")

    else:
        print(f"✅ File {filename} sudah PDF, tidak perlu konversi.")

def cleanup_temp_files(max_age_hours=24):
    """Hapus file di TEMP_DIR yang lebih tua dari 24 jam."""
    now = time.time()
    for filename in os.listdir(TEMP_DIR):
        file_path = os.path.join(TEMP_DIR, filename)
        if os.path.isfile(file_path):
            file_age = now - os.path.getmtime(file_path)
            if file_age > max_age_hours * 3600:
                os.remove(file_path)
                print(f"🗑 Hapus file lama: {file_path}")

# === ROUTING ===
@app.route('/')
def index():
    folder_files = {}
    for k, v in FOLDERS.items():
        files = drive_service.files().list(
            q=f"'{v['id']}' in parents and mimeType='application/pdf'",
            fields="files(id, name)"
        ).execute().get('files', [])

        folder_files[k] = {
            'files': files,
            'name': v['name'],
            'group': v['group'],
            'unlocked': k in session
        }
    return render_template('index.html', folders=folder_files, FOLDERS=FOLDERS)

@app.route("/view/<step>/<file_id>")
@require_folder_password
def view_pdf(step, file_id):
    try:
        final_path = os.path.join(TEMP_DIR, f"{file_id}.pdf")

        # ✅ Kalau file PDF sudah ada di TEMP, langsung pakai
        if os.path.exists(final_path):
            print(f"✅ File PDF sudah ada: {final_path}")
            ready_files[file_id] = final_path
            return redirect(url_for("sign_page", step=step, file_id=file_id))

        # ✅ Ambil metadata file
        file_info = drive_service.files().get(fileId=file_id, fields="name").execute()
        filename = file_info['name']
        print(f"⬇️ Mengunduh file {filename} dari Google Drive...")

        # ✅ Download langsung sebagai PDF
        request = drive_service.files().get_media(fileId=file_id)
        with open(final_path, "wb") as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    print(f"⬇️ Download progress: {int(status.progress() * 100)}%")

        print(f"✅ File berhasil diunduh: {final_path}")

        # ✅ Tandai file siap dipakai
        ready_files[file_id] = final_path
        print(f"📌 File {file_id} siap digunakan: {final_path}")

        return redirect(url_for("sign_page", step=step, file_id=file_id))

    except Exception as e:
        print(f"❌ Terjadi error saat download file: {e}")
        return redirect(url_for("index"))

@app.route('/file/<folder_id>/<file_id>', methods=['GET', 'POST'])
def access_file(folder_id, file_id):
    if folder_id not in FOLDERS:
        return "Folder tidak ditemukan", 404

    if request.method == 'POST':
        entered_password = request.form['password']
        if entered_password == FOLDERS[folder_id]['password']:
            session[folder_id] = True
            return redirect(url_for('loading_page', step=folder_id, file_id=file_id))
        else:
            return render_template('password.html', folder_id=folder_id, file_id=file_id, error="Password salah!")

    return render_template('password.html', folder_id=folder_id, file_id=file_id)

@app.route('/loading/<step>/<file_id>')
@require_folder_password
def loading_page(step, file_id):
    return render_template('loading.html', step=step, file_id=file_id)

@app.route("/check_ready/<file_id>")
def check_ready(file_id):
    file_ready = file_id in ready_files and os.path.exists(ready_files[file_id])
    print(f"✅ Checking if {file_id} is ready... {file_ready}")
    print(f"📂 Isi folder TEMP:", os.listdir(TEMP_DIR))
    return jsonify({"ready": file_ready})

@app.route('/pdf/<step>/<file_id>')
@require_folder_password
def serve_pdf(step, file_id):
    pdf_path = os.path.join(TEMP_DIR, f"{file_id}.pdf")

    # ✅ Kalau belum ada di TEMP, download dari Drive
    if not os.path.exists(pdf_path):
        print(f"⬇️ Mengunduh file PDF {file_id} dari Google Drive...")
        request = drive_service.files().get_media(fileId=file_id)
        with open(pdf_path, "wb") as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    print(f"⬇️ Download progress: {int(status.progress() * 100)}%")
        print(f"✅ File berhasil diunduh: {pdf_path}")

    return send_file(pdf_path, mimetype='application/pdf')

@app.route('/save_signature', methods=['POST'])
@require_folder_password
def save_signature():
    file_id = request.form['file_id']
    folder = request.form['folder']
    signature = request.form['signature']
    file_metadata = drive_service.files().get(fileId=file_id).execute()
    filename = file_metadata['name']

    print(f"📄 Menyimpan tanda tangan untuk file: {filename} (ID: {file_id})")
    print(f"📁 Dari folder: {folder}")

    # ✅ Pastikan file di Google Drive sudah PDF sebelum pindah folder
    ensure_pdf_on_drive(file_id)

    filepath = os.path.join(TEMP_DIR, f"{file_id}.pdf")
    signedpath = os.path.join(TEMP_DIR, f"signed_{file_id}.pdf")

    if folder == "01":
        keywords = ['GA', 'General Affair']
        prefix = request.form.get('prefix')
        bulan = request.form.get('bulan')
        tahun = request.form.get('tahun')

        if not prefix:
            return "❌ Anda harus memilih prefix (SR/MR/GR/SP/MP/GP)", 400

        new_filename = f"{prefix} - {tahun} - {bulan} - {filename}"
        drive_service.files().update(fileId=file_id, body={"name": new_filename}).execute()
        print(f"✏️ File di-rename menjadi: {new_filename}")

        if prefix in ['SR', 'MR', 'GR']:
            next_folder = "02A"
        elif prefix in ['SP', 'MP', 'GP']:
            next_folder = "02B"
        else:
            return "Prefix tidak valid", 400

    elif folder == "02A":
        keywords = ['HRGA']
        prefix = extract_prefix(filename)
        if not prefix:
            return "❌ Prefix tidak ditemukan di nama file", 400
        next_folder = {"SR": "03A", "MR": "03B", "GR": "03C"}.get(prefix)

    elif folder == "02B":
        keywords = ['PAMO']
        prefix = extract_prefix(filename)
        if not prefix:
            return "❌ Prefix tidak ditemukan di nama file", 400
        next_folder = {"SP": "04A", "MP": "04B", "GP": "04C"}.get(prefix)

    elif folder in ["03A", "04A"]:
        keywords = ['SPV']
        prefix = extract_prefix(filename)
        if not prefix:
            return "❌ Prefix tidak ditemukan di nama file", 400
        next_folder = "05"

    elif folder in ["03B", "04B"]:
        keywords = ['Finance Manager', 'Fin Manager', 'Nindy', 'Meri']
        prefix = extract_prefix(filename)
        if not prefix:
            return "❌ Prefix tidak ditemukan di nama file", 400
        next_folder = "05"

    elif folder in ["03C", "04C"]:
        keywords = ['GM', 'General Manager']
        prefix = extract_prefix(filename)
        if not prefix:
            return "❌ Prefix tidak ditemukan di nama file", 400
        next_folder = "05"

    else:
        return "Folder tidak dikenali", 400

    # 🔍 Cari dulu keyword sebelum embed tanda tangan
    found = find_keyword_position(filepath, keywords)
    if not found:
        print("❌ Keyword tanda tangan tidak ditemukan, proses dibatalkan.")
        return "❌ Keyword tanda tangan tidak ditemukan di dokumen", 400

    # Kalau keyword ketemu, baru embed tanda tangan
    embed_signature(filepath, signature, signedpath, keywords)

    # Upload file yang sudah ditandatangani
    media = MediaFileUpload(signedpath, mimetype='application/pdf')
    drive_service.files().update(fileId=file_id, media_body=media).execute()

    # Pindah ke folder berikutnya
    print(f"📂 Memindahkan file ke folder berikutnya: {FOLDERS[next_folder]['name']}")
    file_metadata = drive_service.files().get(fileId=file_id, fields='parents').execute()
    current_parents = ",".join(file_metadata.get('parents', []))
    result = drive_service.files().update(
        fileId=file_id,
        addParents=FOLDERS[next_folder]['id'],
        removeParents=current_parents,
        fields='id, parents'
    ).execute()
    print(f"✅ File dipindah ke folder baru. Parents sekarang: {result.get('parents')}")

    # ✅ Hapus file cache lokal
    try:
        os.remove(filepath)
        os.remove(signedpath)
        print(f"🧹 File cache dihapus: {filepath} & {signedpath}")
    except Exception as e:
        print(f"⚠️ Gagal hapus file sementara: {e}")

    # Redirect ke homepage
    return redirect(url_for('index'))

@app.route('/sign/<step>/<file_id>')
@require_folder_password
def sign_page(step, file_id):
    filepath = os.path.join(TEMP_DIR, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return "File belum siap", 404
    folder = step
    return render_template("view.html", file_id=file_id, folder=folder)

# === RUNNING APP ===
if __name__ == '__main__':
    cleanup_temp_files()
    preprocess_drive_files()
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
