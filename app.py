import os
import io
import base64
import json
import threading
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash, send_from_directory
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter
from werkzeug.utils import secure_filename
from datetime import datetime
import mimetypes
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request

# Muat variabel dari file .env
load_dotenv()

# Fungsi untuk memuat variabel lingkungan dan menangani kesalahan jika tidak ada
def load_env_variable(var_name, is_json=False):
    value = os.getenv(var_name)
    if value is None:
        print(f"ERROR: Variabel lingkungan '{var_name}' tidak ditemukan. Mohon setel di dasbor Render.")
        return None
    
    if is_json:
        try:
            return json.loads(value)
        except json.JSONDecodeError as e:
            print(f"ERROR: Variabel lingkungan '{var_name}' bukan JSON yang valid. {e}")
            return None
    return value

SECRET_KEY = load_env_variable("SECRET_KEY")
GOOGLE_SERVICE_ACCOUNT_JSON = load_env_variable("GOOGLE_SERVICE_ACCOUNT", is_json=True)
FOLDERS = load_env_variable("FOLDERS", is_json=True)
FOLDER_PASSWORDS = load_env_variable("FOLDER_PASSWORDS", is_json=True)
GOOGLE_CLIENT_SECRETS_JSON = load_env_variable("GOOGLE_CLIENT_SECRETS", is_json=True)

# Pastikan semua variabel penting dimuat
if not all([SECRET_KEY, GOOGLE_SERVICE_ACCOUNT_JSON, FOLDERS, FOLDER_PASSWORDS, GOOGLE_CLIENT_SECRETS_JSON]):
    print("FATAL: Satu atau lebih variabel lingkungan penting tidak dimuat. Aplikasi tidak dapat dijalankan.")
    exit(1)

app = Flask(__name__)
app.secret_key = SECRET_KEY

# Direktori untuk file sementara
TEMP_DIR = "temp"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Objek global untuk melacak status unduhan file
DOWNLOAD_STATUS = {}

# Scopes yang dibutuhkan untuk alur OAuth
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

def get_service_account_service():
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

def get_authenticated_service():
    """
    Mendapatkan instance layanan Google Drive yang diautentikasi oleh pengguna
    melalui alur OAuth.
    """
    if 'credentials' not in session:
        return None

    credentials_data = session['credentials']
    credentials = service_account.Credentials.from_authorized_user_info(
        credentials_data,
        scopes=OAUTH_SCOPES
    )

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

@app.route('/')
def home():
    group_data = {
        'Pengajuan Awal': [
            {'name': 'Proposal', 'count': 0, 'id': FOLDERS.get('Proposal', '')},
            {'name': 'Dokumen Pendukung', 'count': 0, 'id': FOLDERS.get('Dokumen Pendukung', '')},
            {'name': 'Legalitas', 'count': 0, 'id': FOLDERS.get('Legalitas', '')},
        ],
        'Rabat': [
            {'name': 'Rabat Bulan Ini', 'count': 0, 'id': FOLDERS.get('Rabat Bulan Ini', '')},
            {'name': 'Rabat Bulan Lalu', 'count': 0, 'id': FOLDERS.get('Rabat Bulan Lalu', '')},
            {'name': 'Riwayat Rabat', 'count': 0, 'id': FOLDERS.get('Riwayat Rabat', '')},
        ],
        'PRS': [
            {'name': 'PRS Tipe A', 'count': 0, 'id': FOLDERS.get('PRS Tipe A', '')},
            {'name': 'PRS Tipe B', 'count': 0, 'id': FOLDERS.get('PRS Tipe B', '')},
            {'name': 'PRS Tipe C', 'count': 0, 'id': FOLDERS.get('PRS Tipe C', '')},
        ],
        'Final': [
            {'name': 'Dokumen Final', 'count': 0, 'id': FOLDERS.get('Dokumen Final', '')},
            {'name': 'Arsip', 'count': 0, 'id': FOLDERS.get('Arsip', '')},
            {'name': 'Laporan Tahunan', 'count': 0, 'id': FOLDERS.get('Laporan Tahunan', '')},
        ],
    }

    # Anda bisa memperbarui 'count' di sini dengan data asli dari Google Drive jika diinginkan
    # Contoh (memerlukan otorisasi):
    # drive_service = get_service_account_service()
    # if drive_service:
    #     for group in group_data.values():
    #         for folder in group:
    #             if folder['id']:
    #                 query = f"'{folder['id']}' in parents and trashed = false"
    #                 results = drive_service.files().list(
    #                     q=query,
    #                     fields="files(id)",
    #                     pageSize=1000
    #                 ).execute()
    #                 folder['count'] = len(results.get('files', []))

    return render_template('index.html', group_data=group_data)

@app.route('/login', methods=['POST'])
def login():
    selected_folder = request.form.get('folder_id')
    password = request.form.get('password')

    folder_name_to_id = {v: k for k, v in FOLDERS.items()}
    selected_folder_name = folder_name_to_id.get(selected_folder)

    if selected_folder_name in FOLDER_PASSWORDS and FOLDER_PASSWORDS[selected_folder_name] == password:
        session['logged_in'] = True
        session['folder_name'] = selected_folder_name
        session['folder_id'] = selected_folder
        # Redirect ke alur otorisasi Google setelah berhasil login
        return redirect(url_for('authorize'))
    else:
        flash('Password salah atau folder tidak ditemukan.')
        return redirect(url_for('home'))

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    session.pop('folder_id', None)
    session.pop('folder_name', None)
    session.pop('credentials', None)
    flash('Berhasil keluar.')
    return redirect(url_for('home'))

@app.route('/folder/<folder_id>')
def folder_view(folder_id):
    if not session.get('logged_in') or session.get('folder_id') != folder_id:
        return redirect(url_for('home'))

    return render_template('folder_view.html', folder_id=folder_id, folder_name=session.get('folder_name'))

@app.route('/authorize')
def authorize():
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
    return redirect(authorization_url)

@app.route('/callback')
def callback():
    state = session.get('oauth_state')
    flow = Flow.from_client_config(
        client_config=GOOGLE_CLIENT_SECRETS_JSON,
        scopes=OAUTH_SCOPES,
        state=state,
        redirect_uri=GOOGLE_REDIRECT_URI
    )
    try:
        flow.fetch_token(authorization_response=request.url)
        credentials = flow.credentials
        session['credentials'] = {
            'token': credentials.token,
            'refresh_token': credentials.refresh_token,
            'token_uri': credentials.token_uri,
            'client_id': credentials.client_id,
            'client_secret': credentials.client_secret,
            'scopes': credentials.scopes
        }
        flash('Otentikasi Google berhasil!')
        return redirect(url_for('folder_view', folder_id=session.get('folder_id')))
    except Exception as e:
        print(f"Error fetching token: {e}")
        flash('Otentikasi Google gagal. Silakan coba lagi.')
        return redirect(url_for('home'))

@app.route('/upload_file', methods=['POST'])
def upload_file():
    if not session.get('logged_in'):
        flash('Anda harus masuk terlebih dahulu.')
        return redirect(url_for('home'))

    if 'file' not in request.files or request.files['file'].filename == '':
        flash('Tidak ada file yang dipilih.')
        return redirect(url_for('folder_view', folder_id=session['folder_id']))

    file = request.files['file']
    user_name = request.form.get('user_name', 'anonim')
    folder_id = session.get('folder_id')
    drive_service = get_authenticated_service()

    if not drive_service:
        flash("Layanan Drive tidak dapat diakses. Silakan masuk kembali.")
        return redirect(url_for('home'))

    try:
        original_filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        new_filename = f"{timestamp}_{user_name}_{original_filename}"
        
        file_mime_type = mimetypes.guess_type(file.filename)[0] or 'application/octet-stream'
        
        media_body = MediaIoBaseUpload(
            io.BytesIO(file.read()),
            mimetype=file_mime_type,
            resumable=True
        )
        
        file_metadata = {
            'name': new_filename,
            'parents': [folder_id]
        }
        
        drive_service.files().create(
            body=file_metadata,
            media_body=media_body,
            fields='id'
        ).execute()
        
        flash(f'Berkas "{new_filename}" berhasil diunggah.')
        
    except Exception as e:
        print(f"Error uploading file: {e}")
        flash(f'Gagal mengunggah berkas: {e}')
        
    return redirect(url_for('folder_view', folder_id=folder_id))

@app.route('/delete_file/<file_id>', methods=['POST'])
def delete_file(file_id):
    if not session.get('logged_in'):
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    drive_service = get_authenticated_service()
    if not drive_service:
        return jsonify({'success': False, 'error': 'Layanan Drive tidak dapat diakses. Silakan masuk kembali.'}), 401
    
    try:
        drive_service.files().delete(fileId=file_id).execute()
        return jsonify({'success': True, 'message': 'Berkas berhasil dihapus.'})
    except Exception as e:
        print(f"Error deleting file: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/list_files/<folder_id>')
def list_files(folder_id):
    if not session.get('logged_in') or session.get('folder_id') != folder_id:
        return jsonify({'error': 'Unauthorized'}), 401
    
    drive_service = get_authenticated_service()
    if not drive_service:
        return jsonify({'error': 'Layanan Google Drive tidak dapat diakses'}), 500

    try:
        query = f"'{folder_id}' in parents and trashed = false"
        results = drive_service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType, size, modifiedTime)",
            pageSize=1000
        ).execute()
        files = results.get('files', [])
        
        def format_size(size_bytes):
            if not size_bytes:
                return "0 B"
            size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
            i = 0
            while size_bytes >= 1024 and i < len(size_name) - 1:
                size_bytes /= 1024
                i += 1
            return f"{size_bytes:.2f} {size_name[i]}"

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

    drive_service = get_authenticated_service()
    if not drive_service:
        flash("Layanan Drive tidak dapat diakses.")
        return redirect(url_for('home'))

    task_id = str(os.urandom(16).hex())
    DOWNLOAD_STATUS[task_id] = {'status': 'pending', 'progress': 0}

    thread = threading.Thread(target=download_file_async, args=(drive_service, file_id, file_name, task_id))
    thread.start()

    return jsonify({'task_id': task_id})

@app.route('/download_status/<task_id>')
def download_status(task_id):
    status = DOWNLOAD_STATUS.get(task_id)
    if not status:
        return jsonify({'error': 'Status tidak ditemukan.'}), 404
    return jsonify(status)
    
@app.route('/get_download_file/<task_id>')
def get_download_file(task_id):
    status = DOWNLOAD_STATUS.get(task_id)
    if not status or status.get('status') != 'complete':
        return jsonify({'error': 'File belum siap untuk diunduh.'}), 400
    
    file_path = status.get('file_path')
    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'File tidak ditemukan.'}), 404

    filename = os.path.basename(file_path)
    mimetype, _ = mimetypes.guess_type(filename)
    if mimetype is None:
        mimetype = 'application/octet-stream'

    response = send_from_directory(
        directory=TEMP_DIR,
        path=filename,
        as_attachment=True,
        mimetype=mimetype,
        download_name=filename
    )
    
    os.remove(file_path)
    del DOWNLOAD_STATUS[task_id]

    return response

# Rute untuk pembuatan PDF dan penandatanganan
def create_pdf(text, file_name):
    path = os.path.join(TEMP_DIR, secure_filename(file_name))
    c = pdf_canvas.Canvas(path, pagesize=A4)
    c.drawString(40, 750, text)
    c.save()
    return path

def add_signature_to_pdf(pdf_path, signature_data):
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)
        
        image_data = base64.b64decode(signature_data.split(',')[1])
        sig_path = os.path.join(TEMP_DIR, f"signature_{os.path.basename(pdf_path)}.png")
        with open(sig_path, "wb") as f:
            f.write(image_data)

        overlay_buffer = io.BytesIO()
        overlay_c = pdf_canvas.Canvas(overlay_buffer, pagesize=A4)
        overlay_c.drawImage(sig_path, 400, 100, width=150, height=50) 
        overlay_c.save()
        overlay_buffer.seek(0)
        overlay_reader = PdfReader(overlay_buffer)
        
        first_page = writer.pages[0]
        first_page.merge_page(overlay_reader.pages[0])
        
        signed_path = os.path.join(TEMP_DIR, f"signed_{os.path.basename(pdf_path)}")
        with open(signed_path, "wb") as f:
            writer.write(f)

        os.remove(sig_path)
        os.remove(pdf_path)
        
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

    pdf_path = create_pdf(text_content, f"{file_name}.pdf")
    if not pdf_path:
        return jsonify({'error': 'Gagal membuat PDF.'}), 500

    signed_pdf_path = add_signature_to_pdf(pdf_path, signature_data)
    if not signed_pdf_path:
        os.remove(pdf_path)
        return jsonify({'error': 'Gagal menambahkan tanda tangan.'}), 500
    
    return send_from_directory(
        directory=TEMP_DIR,
        path=os.path.basename(signed_pdf_path),
        as_attachment=True,
        mimetype="application/pdf"
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)