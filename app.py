import os
import json
import io
import base64
from flask import Flask, render_template, request, redirect, url_for, session, jsonify
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter
import threading

# Muat variabel dari file .env
load_dotenv()
SECRET_KEY = os.getenv("SECRET_KEY")
GOOGLE_SERVICE_ACCOUNT_JSON = json.loads(os.getenv("GOOGLE_SERVICE_ACCOUNT"))
# Periksa apakah FOLDERS dan FOLDER_PASSWORDS sudah dimuat
FOLDERS = json.loads(os.getenv("FOLDERS"))
FOLDER_PASSWORDS = json.loads(os.getenv("FOLDER_PASSWORDS"))

app = Flask(__name__)
app.secret_key = SECRET_KEY

# Direktori untuk file sementara
TEMP_DIR = "temp"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Objek global untuk melacak status unduhan file
DOWNLOAD_STATUS = {}

def get_drive_service():
    """Menginisialisasi dan mengembalikan objek layanan Google Drive."""
    creds = service_account.Credentials.from_service_account_info(
        GOOGLE_SERVICE_ACCOUNT_JSON,
        scopes=["https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/drive.file"]
    )
    return build("drive", "v3", credentials=creds)

drive_service = get_drive_service()

def get_files(folder_id):
    """Mengambil daftar file di dalam folder Google Drive."""
    try:
        results = drive_service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            pageSize=100,
            fields="nextPageToken, files(id, name, parents)"
        ).execute()
        files = results.get("files", [])
        return files
    except Exception as e:
        print(f"Error saat mengambil file: {e}")
        return []

def get_file_by_id(file_id):
    """Mengambil metadata file berdasarkan ID-nya."""
    try:
        return drive_service.files().get(fileId=file_id, fields="id, name, parents").execute()
    except Exception as e:
        print(f"Error saat mengambil file dengan ID {file_id}: {e}")
        return None

def move_file(file_id, new_parent_id):
    """Memindahkan file dari satu folder ke folder lain."""
    try:
        file = drive_service.files().get(fileId=file_id, fields="parents").execute()
        previous_parents = ",".join(file.get("parents"))
        drive_service.files().update(
            fileId=file_id,
            addParents=new_parent_id,
            removeParents=previous_parents,
            fields="id, parents"
        ).execute()
        return True
    except Exception as e:
        print(f"Error saat memindahkan file: {e}")
        return False

def add_signature_to_pdf(input_pdf_path, signature_data_url, keyword):
    """Menambahkan tanda tangan ke PDF. Tanda tangan akan diletakkan di bawah keyword."""
    try:
        # Konversi data URL tanda tangan ke gambar
        header, encoded_data = signature_data_url.split(",", 1)
        signature_binary_data = base64.b64decode(encoded_data)

        # Buat PDF sementara untuk tanda tangan
        sig_pdf_path = os.path.join(TEMP_DIR, "signature.pdf")
        c = pdf_canvas.Canvas(sig_pdf_path, pagesize=A4)
        c.drawImage(
            io.BytesIO(signature_binary_data),
            x=220, y=100, width=150, height=50,
            mask='auto'
        )
        c.save()

        # Baca PDF input dan PDF tanda tangan
        input_pdf = PdfReader(open(input_pdf_path, "rb"))
        sig_pdf = PdfReader(open(sig_pdf_path, "rb"))
        
        output = PdfWriter()

        # Gabungkan tanda tangan ke PDF
        for i in range(len(input_pdf.pages)):
            page = input_pdf.pages[i]
            # Tambahkan tanda tangan ke halaman terakhir
            if i == len(input_pdf.pages) - 1:
                page.merge_page(sig_pdf.pages[0])
            output.add_page(page)

        signed_pdf_path = os.path.join(TEMP_DIR, f"signed_{os.path.basename(input_pdf_path)}")
        with open(signed_pdf_path, "wb") as f:
            output.write(f)

        return signed_pdf_path

    except Exception as e:
        print(f"Error saat menambahkan tanda tangan ke PDF: {e}")
        return None

# ==============================================================================
#                               ROUTING APLIKASI
# ==============================================================================

@app.route("/")
def index():
    """Halaman utama, menampilkan daftar folder berdasarkan grup."""
    # Definisikan grup folder
    folder_groups = {
        "Pengajuan Awal": ["01 - Pengajuan Awal"],
        "Rabat": ["02A - SPV HRGA", "03A - SPV", "03B - Manager", "03C - General"],
        "PRS": ["02B - PAMO", "04A - SPV", "04B - Manager", "04C - General"],
        "Final": ["05 - Final"]
    }
    
    group_data = {}
    for group_name, folders_in_group in folder_groups.items():
        group_data[group_name] = []
        for folder_name in folders_in_group:
            folder_id = FOLDERS.get(folder_name)
            if folder_id:
                files = get_files(folder_id)
                group_data[group_name].append({
                    "name": folder_name,
                    "id": folder_id,
                    "count": len(files)
                })

    return render_template("index.html", group_data=group_data)

@app.route("/folder/<folder_id>", methods=["GET", "POST"])
def view_folder(folder_id):
    """Menampilkan isi folder, dengan otentikasi kata sandi."""
    if request.method == "POST":
        password = request.form.get("password")
        if FOLDER_PASSWORDS.get(folder_id) == password:
            session[folder_id] = True
            return redirect(url_for("view_folder", folder_id=folder_id))
        
        error = "Password yang Anda masukkan salah."
        return render_template("password.html", folder_id=folder_id, error=error)

    if not session.get(folder_id):
        return render_template("password.html", folder_id=folder_id)

    files = get_files(folder_id)
    return render_template("folder.html", files=files, folder_id=folder_id)

@app.route("/load_file/<file_id>")
def load_file(file_id):
    """Halaman loading untuk memulai unduhan file."""
    file = get_file_by_id(file_id)
    if not file:
        return "File tidak ditemukan.", 404
    folder_name = get_folder_name_by_id(file.get('parents')[0])
    return render_template("loading.html", file_id=file_id, folder=folder_name)

@app.route("/start_download/<file_id>")
def start_download(file_id):
    """Mulai proses unduhan file dari Google Drive di latar belakang."""
    global DOWNLOAD_STATUS
    DOWNLOAD_STATUS[file_id] = "downloading"
    
    # Gunakan thread untuk menjalankan unduhan tanpa memblokir
    download_thread = threading.Thread(target=download_file_thread, args=(file_id,))
    download_thread.start()
    
    return jsonify({"status": "download_started"})

def download_file_thread(file_id):
    """Fungsi pembantu untuk mengunduh file dalam thread."""
    global DOWNLOAD_STATUS
    try:
        file_metadata = drive_service.files().get(fileId=file_id).execute()
        filename = file_metadata.get("name")
        temp_pdf_path = os.path.join(TEMP_DIR, filename)

        request_file = drive_service.files().get_media(fileId=file_id)
        with io.FileIO(temp_pdf_path, "wb") as f:
            downloader = MediaIoBaseDownload(f, request_file)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        
        DOWNLOAD_STATUS[file_id] = "ready"
        print(f"File {file_id} berhasil diunduh.")
    except Exception as e:
        print(f"Error saat mengunduh file {file_id}: {e}")
        DOWNLOAD_STATUS[file_id] = "error"

@app.route("/check_ready/<file_id>")
def check_ready(file_id):
    """Memeriksa status unduhan file."""
    global DOWNLOAD_STATUS
    status = DOWNLOAD_STATUS.get(file_id, "pending")
    return jsonify({"ready": status == "ready", "error": status == "error"})

@app.route("/preview_file/<file_id>")
def preview_file(file_id):
    """Menampilkan halaman pratinjau dan tanda tangan."""
    file = get_file_by_id(file_id)
    if not file:
        return "File tidak ditemukan.", 404
    
    folder_name = get_folder_name_by_id(file.get('parents')[0])
    
    return render_template("folder.html", file_id=file.get('id'), folder=folder_name)

def get_folder_name_by_id(folder_id):
    """Mencari nama folder berdasarkan ID."""
    for name, id in FOLDERS.items():
        if id == folder_id:
            return name
    return None

@app.route("/save_signature", methods=["POST"])
def save_signature():
    """Menerima tanda tangan, menambahkan ke PDF, mengunggah kembali, memindahkan, dan mengganti nama file."""
    try:
        data = request.json
        file_id = data.get("file_id")
        folder_name = data.get("folder")
        signature_data = data.get("signature")
        new_filename = data.get("new_filename")

        if not all([file_id, folder_name, signature_data]):
            return "Data tidak lengkap.", 400

        file_metadata = drive_service.files().get(fileId=file_id).execute()
        filename = file_metadata.get("name")
        
        temp_pdf_path = os.path.join(TEMP_DIR, filename)
        
        # Tambahkan tanda tangan
        signed_path = add_signature_to_pdf(temp_pdf_path, signature_data, folder_name)
        if not signed_path:
            return "Gagal menambahkan tanda tangan.", 500

        # Unggah kembali file yang sudah ditandatangani
        media = MediaFileUpload(signed_path, mimetype="application/pdf", resumable=True)
        drive_service.files().update(fileId=file_id, media_body=media).execute()

        # Ganti nama file jika `new_filename` ada
        if new_filename:
            drive_service.files().update(fileId=file_id, body={'name': new_filename}).execute()

        # Pindahkan file ke folder berikutnya
        folder_mapping = {
            "01 - Pengajuan Awal": {
                "SR": "02A - SPV HRGA", "MR": "02A - SPV HRGA", "GR": "02A - SPV HRGA",
                "SP": "02B - PAMO", "MP": "02B - PAMO", "GP": "02B - PAMO",
            },
            "02A - SPV HRGA": {"SR": "03A - SPV", "MR": "03B - Manager", "GR": "03C - General"},
            "02B - PAMO": {"SP": "04A - SPV", "MP": "04B - Manager", "GP": "04C - General"},
            "03A - SPV": {"SR": "05 - Final"},
            "03B - Manager": {"MR": "05 - Final"},
            "03C - General": {"GR": "05 - Final"},
            "04A - SPV": {"SP": "05 - Final"},
            "04B - Manager": {"MP": "05 - Final"},
            "04C - General": {"GP": "05 - Final"},
        }

        # Ambil prefix dari `new_filename` jika ada, atau dari `filename` yang lama
        prefix_to_map = new_filename.split()[0][:2].upper() if new_filename else filename.split()[0][:2].upper()

        target_folder_name = folder_mapping.get(folder_name, {}).get(prefix_to_map)
        if target_folder_name:
            target_id = FOLDERS.get(target_folder_name)
            if target_id:
                move_file(file_id, target_id)

        # Hapus file sementara
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(signed_path):
            os.remove(signed_path)

        return "OK", 200

    except Exception as e:
        print(f"Error dalam save_signature: {e}")
        return "Terjadi kesalahan server.", 500

if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
