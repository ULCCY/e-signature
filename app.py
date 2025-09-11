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

# Import necessary classes for the OAuth flow
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
import pickle

# Muat variabel dari file .env
load_dotenv()
SECRET_KEY = os.getenv("SECRET_KEY")
GOOGLE_SERVICE_ACCOUNT_JSON = json.loads(os.getenv("GOOGLE_SERVICE_ACCOUNT"))
# Periksa apakah FOLDERS dan FOLDER_PASSWORDS sudah dimuat
FOLDERS = json.loads(os.getenv("FOLDERS"))
FOLDER_PASSWORDS = json.loads(os.getenv("FOLDER_PASSWORDS"))

# Load client secrets from environment variable
GOOGLE_CLIENT_SECRETS_JSON = json.loads(os.getenv("GOOGLE_CLIENT_SECRETS"))

app = Flask(__name__)
app.secret_key = SECRET_KEY

# Direktori untuk file sementara
TEMP_DIR = "temp"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Objek global untuk melacak status unduhan file
DOWNLOAD_STATUS = {}

def get_drive_service():
    """
    Mendapatkan instance layanan Google Drive dengan otentikasi akun layanan.
    """
    try:
        credentials = service_account.Credentials.from_service_account_info(
            GOOGLE_SERVICE_ACCOUNT_JSON,
            scopes=['https://www.googleapis.com/auth/drive']
        )
        return build('drive', 'v3', credentials=credentials)
    except Exception as e:
        print(f"Error creating Google Drive service: {e}")
        return None

DRIVE_SERVICE = get_drive_service()

if DRIVE_SERVICE is None:
    print("Warning: Google Drive service could not be initialized. Check your environment variables.")

# Scopes yang dibutuhkan
OAUTH_SCOPES = [
    'https://www.googleapis.com/auth/drive.appdata',
    'https://www.googleapis.com/auth/drive.metadata.readonly',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid',
]

# URL redirect setelah otorisasi
GOOGLE_REDIRECT_URI = os.getenv("GOOGLE_REDIRECT_URI", "http://localhost:5000/callback")
# Pastikan URL redirect sama persis dengan yang terdaftar di Google Cloud Console
# Untuk Render, ini harus berupa URL publik Anda, misalnya: https://e-signature-d7er.onrender.com/callback

@app.route('/')
def home():
    return render_template('index.html', folders=FOLDERS)

@app.route('/login', methods=['POST'])
def login():
    selected_folder = request.form.get('folder')
    password = request.form.get('password')

    if selected_folder in FOLDER_PASSWORDS and FOLDER_PASSWORDS[selected_folder] == password:
        session['logged_in'] = True
        session['folder_name'] = selected_folder
        session['folder_id'] = FOLDERS[selected_folder]
        return redirect(url_for('folder_view', folder_id=session['folder_id']))
    else:
        flash('Password salah atau folder tidak ditemukan.')
        return redirect(url_for('home'))

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    session.pop('folder_id', None)
    session.pop('folder_name', None)
    return redirect(url_for('home'))

@app.route('/folder/<folder_id>')
def folder_view(folder_id):
    if not session.get('logged_in') or session.get('folder_id') != folder_id:
        return redirect(url_for('home'))

    return render_template('folder_view.html', folder_id=folder_id, folder_name=session.get('folder_name'))

@app.route('/authorize')
def authorize():
    # Menggunakan `from_client_config` untuk memuat rahasia klien dari memori
    # Alih-alih dari file, yang tidak tersedia di lingkungan Render
    flow = Flow.from_client_config(
        client_config=GOOGLE_CLIENT_SECRETS_JSON,
        scopes=OAUTH_SCOPES,
        redirect_uri=GOOGLE_REDIRECT_URI
    )

    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true'
    )

    session['oauth_state'] = state
    session['redirect_uri'] = GOOGLE_REDIRECT_URI

    return redirect(authorization_url)

@app.route('/callback')
def callback():
    state = session.get('oauth_state')
    redirect_uri = session.get('redirect_uri')

    flow = Flow.from_client_config(
        client_config=GOOGLE_CLIENT_SECRETS_JSON,
        scopes=OAUTH_SCOPES,
        state=state,
        redirect_uri=redirect_uri
    )

    try:
        flow.fetch_token(authorization_response=request.url)
    except Exception as e:
        print(f"Error fetching token: {e}")
        flash('Otentikasi Google gagal. Silakan coba lagi.')
        return redirect(url_for('home'))

    credentials = flow.credentials
    session['credentials'] = {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }

    flash('Otentikasi berhasil! Sekarang Anda dapat mengunggah file ke Google Drive.')
    return redirect(url_for('folder_view', folder_id=session['folder_id']))


def get_authenticated_service():
    if 'credentials' not in session:
        return None

    credentials_data = session['credentials']
    credentials = service_account.Credentials.from_authorized_user_info(
        credentials_data,
        scopes=OAUTH_SCOPES
    )

    # Memeriksa kredensial, jika hampir kedaluwarsa, segarkan
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        session['credentials'] = {
            'token': credentials.token,
            'refresh_token': credentials.refresh_token,
            'token_uri': credentials.token_uri,
            'client_id': credentials.client_id,
            'client_secret': credentials.client_secret,
            'scopes': credentials.scopes
        }

    return build('drive', 'v3', credentials=credentials)

@app.route('/upload_file', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('Tidak ada file yang dipilih.')
        return redirect(request.referrer or url_for('home'))

    file = request.files['file']
    if file.filename == '':
        flash('Tidak ada file yang dipilih.')
        return redirect(request.referrer or url_for('home'))

    folder_id = request.form.get('folder_id')
    user_name = request.form.get('user_name')

    if not all([folder_id, user_name, file]):
        flash('Data tidak lengkap. Unggahan dibatalkan.')
        return redirect(request.referrer or url_for('home'))

    drive_service = get_drive_service()
    if not drive_service:
        flash("Layanan Drive tidak dapat diakses.")
        return redirect(request.referrer or url_for('home'))

    try:
        file_metadata = {
            'name': f"{user_name}_{file.filename}",
            'parents': [folder_id]
        }
        media = MediaIoBaseUpload(file.stream, mimetype=file.mimetype, resumable=True)
        file_obj = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        
        flash(f'File "{file.filename}" berhasil diunggah dengan nama baru "{user_name}_{file.filename}".')
        
    except Exception as e:
        print(f"Error saat mengunggah file: {e}")
        # Tangani error khusus jika kuota terlampaui
        if "storageQuotaExceeded" in str(e):
            flash("Gagal mengunggah file: Kuota penyimpanan akun layanan Anda telah terlampaui.")
        else:
            flash(f"Error saat mengunggah file: {e}")
            
    return redirect(request.referrer or url_for('home'))


@app.route('/list_files/<folder_id>')
def list_files(folder_id):
    if not session.get('logged_in') or session.get('folder_id') != folder_id:
        return jsonify({'error': 'Unauthorized'}), 401
    
    drive_service = get_drive_service()
    if not drive_service:
        return jsonify({'error': 'Layanan Google Drive tidak dapat diakses'}), 500

    try:
        # Gunakan 'in parents' untuk mencari file di dalam folder tertentu
        query = f"'{folder_id}' in parents and trashed = false"
        results = drive_service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType, size, modifiedTime)",
            pageSize=1000
        ).execute()
        files = results.get('files', [])
        
        # Format ukuran file agar lebih mudah dibaca
        def format_size(size_bytes):
            if not size_bytes:
                return "0 B"
            size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
            i = 0
            while size_bytes >= 1024 and i < len(size_name) - 1:
                size_bytes /= 1024
                i += 1
            return f"{size_bytes:.2f} {size_name[i]}"

        # Tambahkan informasi tambahan ke setiap file
        file_list = []
        for file in files:
            file_info = {
                'id': file['id'],
                'name': file['name'],
                'mimeType': file['mimeType'],
                'modifiedTime': file['modifiedTime'],
                'size': format_size(int(file.get('size', 0)))
            }
            file_list.append(file_info)

        return jsonify({'files': file_list})

    except Exception as e:
        print(f"Error listing files: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/delete_file/<file_id>', methods=['POST'])
def delete_file(file_id):
    if not session.get('logged_in'):
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401

    drive_service = get_authenticated_service()
    if not drive_service:
        return jsonify({'success': False, 'error': 'Layanan Drive tidak dapat diakses'}), 500

    try:
        drive_service.files().delete(fileId=file_id).execute()
        return jsonify({'success': True, 'message': 'File berhasil dihapus.'})
    except Exception as e:
        print(f"Error deleting file: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/download_status/<task_id>')
def download_status(task_id):
    status = DOWNLOAD_STATUS.get(task_id)
    if not status:
        return jsonify({'error': 'Status tidak ditemukan.'}), 404
    return jsonify(status)

def download_file_async(drive_service, file_id, file_name, task_id):
    try:
        request = drive_service.files().get_media(fileId=file_id)
        file_path = os.path.join(TEMP_DIR, secure_filename(file_name))
        file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                progress = int(status.progress() * 100)
                DOWNLOAD_STATUS[task_id] = {'status': 'in_progress', 'progress': progress}
        
        file_stream.seek(0)
        with open(file_path, 'wb') as f:
            f.write(file_stream.read())
        
        DOWNLOAD_STATUS[task_id] = {'status': 'complete', 'file_path': file_path}

    except Exception as e:
        print(f"Error during async download: {e}")
        DOWNLOAD_STATUS[task_id] = {'status': 'error', 'error_message': str(e)}

@app.route('/download_file/<file_id>/<file_name>')
def download_file(file_id, file_name):
    if not session.get('logged_in'):
        flash("Anda harus login untuk mengunduh file.")
        return redirect(url_for('home'))

    drive_service = get_drive_service()
    if not drive_service:
        flash("Layanan Drive tidak dapat diakses.")
        return redirect(url_for('home'))

    # Buat ID tugas unik untuk melacak status
    task_id = str(os.urandom(16).hex())
    DOWNLOAD_STATUS[task_id] = {'status': 'pending', 'progress': 0}

    # Jalankan unduhan di thread terpisah
    thread = threading.Thread(target=download_file_async, args=(drive_service, file_id, file_name, task_id))
    thread.start()

    return jsonify({'task_id': task_id})

@app.route('/get_download_file/<task_id>')
def get_download_file(task_id):
    status = DOWNLOAD_STATUS.get(task_id)
    if not status or status.get('status') != 'complete':
        return jsonify({'error': 'File belum siap untuk diunduh.'}), 400
    
    file_path = status.get('file_path')
    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'File tidak ditemukan.'}), 404

    # Ambil nama file asli
    filename = os.path.basename(file_path)
    
    # Deteksi mimetype
    mimetype, _ = mimetypes.guess_type(filename)
    if mimetype is None:
        mimetype = 'application/octet-stream'

    # Kirim file
    response = send_from_directory(
        directory=TEMP_DIR,
        path=filename,
        as_attachment=True,
        mimetype=mimetype,
        download_name=filename
    )

    # Hapus file setelah dikirim
    os.remove(file_path)
    del DOWNLOAD_STATUS[task_id]

    return response

# Rute untuk pembuatan PDF dan penandatanganan
def create_pdf(text, file_name):
    path = os.path.join(TEMP_DIR, file_name)
    c = pdf_canvas.Canvas(path, pagesize=A4)
    c.drawString(100, 750, text)
    c.save()
    return path

def add_signature_to_pdf(pdf_path, signature_data):
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        # Dapatkan gambar dari data base64
        image_data = base64.b64decode(signature_data.split(',')[1])
        img = io.BytesIO(image_data)

        for page in reader.pages:
            writer.add_page(page)

        # Tambahkan tanda tangan ke halaman terakhir
        last_page = writer.pages[-1]
        
        # Simpan tanda tangan sementara
        sig_path = os.path.join(TEMP_DIR, "signature.png")
        with open(sig_path, "wb") as f:
            f.write(image_data)

        # Ubah ukuran tanda tangan
        # Menggunakan reportlab untuk menambahkan gambar
        pdf_signer = pdf_canvas.Canvas(os.path.join(TEMP_DIR, "signed_temp.pdf"), pagesize=A4)
        pdf_signer.drawInlineImage(sig_path, x=100, y=100, width=100, height=50) # Sesuaikan ukuran dan posisi
        pdf_signer.showPage()
        pdf_signer.save()
        
        # Gabungkan PDF
        signed_reader = PdfReader(os.path.join(TEMP_DIR, "signed_temp.pdf"))
        # Asumsi ini hanyalah contoh, dalam implementasi nyata Anda perlu menggabungkan halaman dengan benar
        
        # Gabungkan file
        output_pdf = PdfWriter()
        for p in reader.pages:
            output_pdf.add_page(p)

        # Tambahkan tanda tangan ke halaman
        # Ini adalah contoh sederhana, Anda mungkin perlu mengkombinasikan halaman secara lebih kompleks
        # atau menggunakan library lain untuk penandatanganan PDF
        
        # Simpan file yang telah ditandatangani
        signed_path = os.path.join(TEMP_DIR, f"signed_{os.path.basename(pdf_path)}")
        with open(signed_path, "wb") as f:
            output_pdf.write(f)

        os.remove(sig_path)
        # os.remove(os.path.join(TEMP_DIR, "signed_temp.pdf"))

        return signed_path
    
    except Exception as e:
        print(f"Error adding signature to PDF: {e}")
        return None

@app.route('/create_and_sign_pdf', methods=['POST'])
def create_and_sign_pdf():
    text_content = request.form.get('text_content')
    signature_data = request.form.get('signature_data')
    file_name = request.form.get('file_name')
    
    if not all([text_content, signature_data, file_name]):
        return jsonify({'error': 'Data tidak lengkap.'}), 400

    pdf_path = create_pdf(text_content, file_name)
    if not pdf_path:
        return jsonify({'error': 'Gagal membuat PDF.'}), 500

    signed_pdf_path = add_signature_to_pdf(pdf_path, signature_data)
    if not signed_pdf_path:
        os.remove(pdf_path)
        return jsonify({'error': 'Gagal menambahkan tanda tangan.'}), 500
    
    # Hapus file PDF sementara setelah penandatanganan
    os.remove(pdf_path)

    return send_from_directory(
        directory=TEMP_DIR,
        path=os.path.basename(signed_pdf_path),
        as_attachment=True,
        mimetype="application/pdf"
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
