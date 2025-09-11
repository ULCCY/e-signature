import os
import io
import base64
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash, send_from_directory
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaFileUpload, MediaIoBaseDownload
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter
import json
import threading
from werkzeug.utils import secure_filename
from datetime import datetime
import mimetypes

# Muat variabel dari file .env
load_dotenv()
SECRET_KEY = os.getenv("SECRET_KEY")

# Periksa dan muat variabel lingkungan sebagai JSON
google_service_account_str = os.getenv("GOOGLE_SERVICE_ACCOUNT")
if google_service_account_str:
    GOOGLE_SERVICE_ACCOUNT_JSON = json.loads(google_service_account_str)
else:
    raise ValueError("Variabel lingkungan GOOGLE_SERVICE_ACCOUNT tidak ditemukan.")

folders_str = os.getenv("FOLDERS")
if folders_str:
    FOLDERS = json.loads(folders_str)
else:
    raise ValueError("Variabel lingkungan FOLDERS tidak ditemukan.")

folder_passwords_str = os.getenv("FOLDER_PASSWORDS")
if folder_passwords_str:
    FOLDER_PASSWORDS = json.loads(folder_passwords_str)
else:
    raise ValueError("Variabel lingkungan FOLDER_PASSWORDS tidak ditemukan.")

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
        scopes=["https://www.googleapis.com/auth/drive"]
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
        header, encoded_data = signature_data_url.split(",", 1)
        signature_binary_data = base64.b64decode(encoded_data)

        sig_pdf_path = os.path.join(TEMP_DIR, "signature.pdf")
        c = pdf_canvas.Canvas(sig_pdf_path, pagesize=A4)
        c.drawImage(
            io.BytesIO(signature_binary_data),
            x=220, y=100, width=150, height=50,
            mask='auto'
        )
        c.save()

        input_pdf = PdfReader(open(input_pdf_path, "rb"))
        sig_pdf = PdfReader(open(sig_pdf_path, "rb"))
        
        output = PdfWriter()

        for i in range(len(input_pdf.pages)):
            page = input_pdf.pages[i]
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
#                                   ROUTING APLIKASI
# ==============================================================================
@app.route("/")
def index():
    """Halaman utama, menampilkan daftar folder berdasarkan grup."""
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

@app.route("/folder/<folder_id>", methods=["GET", "POST"]) # Perbaiki metode di sini
def view_folder(folder_id):
    """Menampilkan isi folder dengan otentikasi sesi."""
    folder_name = get_folder_name_by_id(folder_id)
    if not folder_name:
        return "Folder tidak ditemukan.", 404

    # Tangani permintaan POST (login)
    if request.method == "POST":
        password = request.form.get("password")
        if FOLDER_PASSWORDS.get(folder_id) == password:
            session["logged_in"] = True
            session["folder_id"] = folder_id
            flash("Login berhasil!", "success")
            return redirect(url_for("view_folder", folder_id=folder_id))
        else:
            flash("Password salah. Silakan coba lagi.", "error")
            return render_template("password.html", folder_id=folder_id, folder_name=folder_name)

    # Tangani permintaan GET (akses halaman)
    if not session.get("logged_in") or session.get("folder_id") != folder_id:
        return render_template("password.html", folder_id=folder_id, folder_name=folder_name)

    files = get_files(folder_id)
    is_pengajuan_awal = folder_name == "01 - Pengajuan Awal"
    return render_template("folder.html", files=files, folder_id=folder_id, is_pengajuan_awal=is_pengajuan_awal)

@app.route("/upload_file", methods=["POST"])
def upload_file():
    # Ambil folder ID dari form
    target_folder_id = request.form.get("folder_id")

    # Logika otentikasi
    if not session.get("logged_in") or session.get("folder_id") != target_folder_id:
        flash("Silakan login kembali untuk mengunggah file.", "error")
        return redirect(url_for("view_folder", folder_id=target_folder_id))
    
    try:
        uploaded_file = request.files.get("file")
        if not uploaded_file or uploaded_file.filename == "":
            flash("Tidak ada file yang dipilih.", "error")
            return redirect(url_for("view_folder", folder_id=target_folder_id))

        filename = secure_filename(uploaded_file.filename)
        mime_type = uploaded_file.content_type
        
        file_metadata = {
            "name": filename,
            "parents": [target_folder_id]
        }
        
        # Cek apakah file perlu dikonversi ke PDF
        is_conversion_needed = False
        if mime_type in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         "application/msword",
                         "application/vnd.oasis.opendocument.text"]:
            
            is_conversion_needed = True
            if "." in filename:
                name_without_ext = filename.rsplit(".", 1)[0]
                filename = f"{name_without_ext}.pdf"
            else:
                filename += ".pdf"

            file_metadata["name"] = filename
            file_metadata["mimeType"] = "application/pdf"
            
        media = MediaIoBaseUpload(uploaded_file.stream, mimetype=mime_type, resumable=True)

        drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()

        flash(f"File '{filename}' berhasil diunggah dan dikonversi.", "success")
        
    except Exception as e:
        print(f"Error saat mengunggah file: {e}")
        flash(f"Error: Gagal mengunggah file. {e}", "error")

    return redirect(url_for("view_folder", folder_id=target_folder_id))


@app.route("/delete_file/<file_id>", methods=["POST"])
def delete_file(file_id):
    """Menghapus file dari Google Drive."""
    folder_name = get_folder_name_by_id(request.form.get("folder_id"))
    
    # Pastikan penghapusan hanya bisa dilakukan di folder '01 - Pengajuan Awal'
    if folder_name != "01 - Pengajuan Awal":
        return "Akses Ditolak: Anda tidak memiliki izin untuk menghapus file di folder ini.", 403

    try:
        drive_service.files().delete(fileId=file_id).execute()
        flash("File berhasil dihapus.", "success")
    except Exception as e:
        print(f"Error saat menghapus file: {e}")
        flash(f"Gagal menghapus file: {e}", "error")

    return redirect(url_for("view_folder", folder_id=request.form.get("folder_id")))

@app.route("/load_file/<file_id>")
def load_file(file_id):
    """Halaman loading untuk memulai unduhan file."""
    file = get_file_by_id(file_id)
    if not file:
        return "File tidak ditemukan.", 404
    folder_id = file.get('parents')[0]
    folder_name = get_folder_name_by_id(folder_id)
    return render_template("loading.html", file_id=file_id, folder=folder_name, folder_id=folder_id)

@app.route("/start_download/<file_id>")
def start_download(file_id):
    """Mulai proses unduhan file dari Google Drive di latar belakang."""
    global DOWNLOAD_STATUS
    DOWNLOAD_STATUS[file_id] = "downloading"
    
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

@app.route("/download_pdf/<file_id>")
def download_pdf(file_id):
    """Mengirim file PDF yang sudah diunduh ke browser untuk pratinjau."""
    file_metadata = drive_service.files().get(fileId=file_id).execute()
    filename = file_metadata.get("name")
    return send_from_directory(TEMP_DIR, filename, mimetype=mimetypes.guess_type(filename)[0])

@app.route("/preview_file/<file_id>")
def preview_file(file_id):
    """Menampilkan halaman pratinjau dan tanda tangan."""
    file = get_file_by_id(file_id)
    if not file:
        return "File tidak ditemukan.", 404
    
    folder_id = file.get('parents')[0]
    folder_name = get_folder_name_by_id(folder_id)
    
    # Tentukan apakah ini folder '05 - Final' atau '01 - Pengajuan Awal'
    is_final_folder = folder_name == "05 - Final"
    is_pengajuan_awal = folder_name == "01 - Pengajuan Awal"
    
    return render_template(
        "preview.html", 
        file_id=file.get('id'), 
        folder=folder_name, 
        folder_id=folder_id,
        is_final_folder=is_final_folder,
        is_pengajuan_awal=is_pengajuan_awal
    )

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
        
        # Data tambahan dari form 'Pengajuan Awal'
        pengajuan_bulan = data.get("pengajuan_bulan")
        pengajuan_tahun = data.get("pengajuan_tahun")
        perusahaan = data.get("perusahaan")
        pengajuan_akhir = data.get("pengajuan_akhir")
        
        # Jika folder 'Pengajuan Awal', lakukan validasi
        if folder_name == "01 - Pengajuan Awal":
            if not all([signature_data, pengajuan_bulan, pengajuan_tahun, perusahaan, pengajuan_akhir]):
                return "Mohon lengkapi semua data dan tanda tangan.", 400

        # Hanya jalankan proses tanda tangan jika bukan folder 'Final'
        if folder_name != "05 - Final":
            file_metadata = drive_service.files().get(fileId=file_id).execute()
            filename = file_metadata.get("name")
            temp_pdf_path = os.path.join(TEMP_DIR, filename)
            signed_path = add_signature_to_pdf(temp_pdf_path, signature_data, folder_name)
            if not signed_path:
                return "Gagal menambahkan tanda tangan.", 500

            # Unggah kembali file yang sudah ditandatangani
            media = MediaFileUpload(signed_path, mimetype="application/pdf", resumable=True)
            drive_service.files().update(fileId=file_id, media_body=media).execute()
        
            # Ganti nama file jika data dari form 'Pengajuan Awal' ada
            if folder_name == "01 - Pengajuan Awal":
                now = datetime.now()
                month = now.strftime("%m")
                year = now.strftime("%y")
                
                kode_pengajuan = pengajuan_akhir.upper() + perusahaan.upper()
                original_filename = filename
                
                new_filename = f"{year}/{month} {kode_pengajuan} - {original_filename}"
                drive_service.files().update(fileId=file_id, body={'name': new_filename}).execute()
        
        # Pindahkan file ke folder berikutnya
        folder_mapping = {
            "01 - Pengajuan Awal": {
                "SR": "02A - SPV HRGA", "MR": "02B - PAMO", "GR": "02B - PAMO",
                "SP": "02A - SPV HRGA", "MP": "02A - SPV HRGA", "GP": "02A - SPV HRGA",
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
        
        # Ambil kode pengajuan dari form jika ada, atau dari nama file
        kode_pengajuan_to_map = kode_pengajuan if folder_name == "01 - Pengajuan Awal" else filename.split()[1][:2].upper()
        
        target_folder_name = folder_mapping.get(folder_name, {}).get(kode_pengajuan_to_map)
        if target_folder_name:
            target_id = FOLDERS.get(target_folder_name)
            if target_id:
                move_file(file_id, target_id)

        # Hapus file sementara
        if folder_name != "05 - Final":
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