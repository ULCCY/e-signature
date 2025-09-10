import os
import io
import base64
import json
import threading
import mimetypes
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash, send_from_directory
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaFileUpload, MediaIoBaseDownload
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter
from werkzeug.utils import secure_filename
from datetime import datetime
from oauthlib.oauth2.rfc6749.errors import AccessDeniedError

# Muat variabel dari file .env
load_dotenv()
SECRET_KEY = os.getenv("SECRET_KEY")
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

# Konfigurasi OAuth 2.0
SCOPES = ['https://www.googleapis.com/auth/drive']

@app.route('/')
def index():
    creds = session.get('creds')
    if not creds:
        return render_template('index.html')

    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)

    try:
        # Panggil API Drive v3
        response = service.files().list(
            q="'root' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="nextPageToken, files(id, name)").execute()
        
        main_folders = {file['name']: file['id'] for file in response.get('files', [])}
        
        folders_with_passwords = {name: {'id': main_folders.get(name), 'password_required': name in FOLDER_PASSWORDS} for name in FOLDERS.keys()}

        return render_template('index.html', folders=folders_with_passwords)
    except Exception as e:
        flash(f"Error accessing Google Drive: {e}", "error")
        return redirect(url_for('index'))

@app.route('/authorize')
def authorize():
    flow = InstalledAppFlow.from_client_config(
        json.loads(os.getenv("CLIENT_SECRETS")),
        SCOPES,
        redirect_uri=url_for('oauth2callback', _external=True)
    )
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true'
    )
    session['state'] = state
    return redirect(authorization_url)

@app.route('/oauth2callback')
def oauth2callback():
    state = session.get('state')
    if not state or state != request.args.get('state'):
        flash("State parameter mismatch.", "error")
        return redirect(url_for('index'))

    # Menggunakan try-except untuk menangani AccessDeniedError
    try:
        flow = InstalledAppFlow.from_client_config(
            json.loads(os.getenv("CLIENT_SECRETS")),
            SCOPES,
            redirect_uri=url_for('oauth2callback', _external=True)
        )
        authorization_response = request.url
        flow.fetch_token(authorization_response=authorization_response)
        
        creds = flow.credentials.to_json()
        session['creds'] = creds
        return redirect(url_for('index'))
    except AccessDeniedError:
        flash("Authorization failed. You denied access to the application.", "error")
        return redirect(url_for('index'))
    except Exception as e:
        flash(f"An unexpected error occurred: {e}", "error")
        return redirect(url_for('index'))

@app.route('/view_folder/<folder_id>')
def view_folder(folder_id):
    creds = session.get('creds')
    if not creds:
        return redirect(url_for('authorize'))
    
    # Periksa apakah folder_id ada di FOLDERS
    folder_name = next((name for name, id in FOLDERS.items() if id == folder_id), None)
    if folder_name is None:
        flash("Folder not found.", "error")
        return redirect(url_for('index'))

    # Periksa password jika diperlukan
    password = request.args.get('password')
    if folder_name in FOLDER_PASSWORDS:
        if not password or password != FOLDER_PASSWORDS[folder_name]:
            # Jika password salah, tampilkan pesan error dan kembali
            flash("Incorrect password for this folder.", "error")
            return redirect(url_for('index'))
    
    # Simpan folder saat ini di sesi
    session['current_folder_id'] = folder_id

    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)

    try:
        # Dapatkan file dan folder di dalam folder yang dipilih
        response = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name, mimeType, modifiedTime, size)").execute()
        
        items = response.get('files', [])
        
        files = []
        subfolders = []
        for item in items:
            if item['mimeType'] == 'application/vnd.google-apps.folder':
                subfolders.append({
                    'id': item['id'],
                    'name': item['name']
                })
            else:
                files.append({
                    'id': item['id'],
                    'name': item['name'],
                    'modifiedTime': item.get('modifiedTime'),
                    'size': item.get('size'),
                    'mimeType': item['mimeType']
                })

        return render_template('folder_content.html', folder_name=folder_name, subfolders=subfolders, files=files)
    except Exception as e:
        flash(f"Error accessing folder: {e}", "error")
        return redirect(url_for('index'))


@app.route('/upload_file', methods=['POST'])
def upload_file():
    creds = session.get('creds')
    if not creds:
        return jsonify({'error': 'Unauthorized'}), 401

    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)
    
    current_folder_id = session.get('current_folder_id')
    if not current_folder_id:
        return jsonify({'error': 'No active folder selected'}), 400

    try:
        filename = secure_filename(file.filename)
        mime_type, _ = mimetypes.guess_type(filename)
        
        # Baca konten file
        file_content = file.read()
        
        # Buat MediaIoBaseUpload
        media = MediaIoBaseUpload(io.BytesIO(file_content), mime_type, resumable=True)
        
        file_metadata = {
            'name': filename,
            'parents': [current_folder_id]
        }
        
        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        flash("File uploaded successfully!", "success")
        return jsonify({'message': 'File uploaded successfully', 'file_id': uploaded_file.get('id')}), 200
        
    except Exception as e:
        flash(f"Error uploading file: {e}", "error")
        return jsonify({'error': f"Error uploading file: {e}"}), 500


@app.route('/create_pdf', methods=['POST'])
def create_pdf():
    creds = session.get('creds')
    if not creds:
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.get_json()
    text = data.get('text', '')
    filename = data.get('filename', 'document.pdf')
    
    if not text:
        return jsonify({'error': 'No text provided'}), 400

    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)
    
    current_folder_id = session.get('current_folder_id')
    if not current_folder_id:
        return jsonify({'error': 'No active folder selected'}), 400

    try:
        # Buat file PDF di memori
        buffer = io.BytesIO()
        c = pdf_canvas.Canvas(buffer, pagesize=A4)
        c.drawString(100, 750, "Generated PDF from Text")
        c.drawString(100, 730, text)
        c.showPage()
        c.save()
        
        buffer.seek(0)
        
        media = MediaIoBaseUpload(buffer, 'application/pdf', resumable=True)
        
        file_metadata = {
            'name': filename,
            'parents': [current_folder_id]
        }
        
        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        flash("PDF created and uploaded successfully!", "success")
        return jsonify({'message': 'PDF created successfully', 'file_id': uploaded_file.get('id')}), 200
        
    except Exception as e:
        flash(f"Error creating PDF: {e}", "error")
        return jsonify({'error': f"Error creating PDF: {e}"}), 500


@app.route('/sign_document', methods=['POST'])
def sign_document():
    creds = session.get('creds')
    if not creds:
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.get_json()
    file_id = data.get('file_id')
    signature_base64 = data.get('signature')
    
    if not file_id or not signature_base64:
        return jsonify({'error': 'File ID or signature not provided'}), 400

    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)
    
    try:
        # Download file
        request_download = service.files().get_media(fileId=file_id)
        file_data = io.BytesIO()
        downloader = MediaIoBaseDownload(file_data, request_download)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        file_data.seek(0)
        
        # Tanda tangani dokumen
        signature_image = io.BytesIO(base64.b64decode(signature_base64.split(',')[1]))
        
        reader = PdfReader(file_data)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        # Buat PDF untuk tanda tangan
        sig_buffer = io.BytesIO()
        sig_canvas = pdf_canvas.Canvas(sig_buffer, pagesize=A4)
        sig_canvas.drawInlineImage(signature_image, 100, 100, width=50, height=50) # Tentukan posisi tanda tangan
        sig_canvas.showPage()
        sig_canvas.save()
        sig_buffer.seek(0)
        
        sig_reader = PdfReader(sig_buffer)
        
        # Gabungkan tanda tangan dengan dokumen asli
        writer.add_page(sig_reader.pages[0])
        
        output_buffer = io.BytesIO()
        writer.write(output_buffer)
        output_buffer.seek(0)

        # Upload kembali file yang ditandatangani
        media = MediaIoBaseUpload(output_buffer, 'application/pdf', resumable=True)
        
        updated_file = service.files().update(
            fileId=file_id,
            media_body=media,
            fields='id'
        ).execute()
        
        flash("Document signed successfully!", "success")
        return jsonify({'message': 'Document signed successfully', 'file_id': updated_file.get('id')}), 200
        
    except Exception as e:
        flash(f"Error signing document: {e}", "error")
        return jsonify({'error': f"Error signing document: {e}"}), 500

@app.route('/move_file', methods=['POST'])
def move_file():
    creds = session.get('creds')
    if not creds:
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.get_json()
    file_id = data.get('file_id')
    
    if not file_id:
        return jsonify({'error': 'File ID not provided'}), 400

    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)
    
    current_folder_id = session.get('current_folder_id')
    if not current_folder_id:
        return jsonify({'error': 'No active folder selected'}), 400

    try:
        file = service.files().get(fileId=file_id, fields='name, parents').execute()
        current_parents = file.get('parents', [])
        current_folder_name = next((name for name, id in FOLDERS.items() if id == current_folder_id), None)
        
        if not current_folder_name:
            return jsonify({'error': 'Current folder not in configured folders'}), 400

        # Logic for determining the next folder based on the file name
        filename = file.get('name')
        
        # Mendapatkan nama folder saat ini
        folder_name = next((name for name, id in FOLDERS.items() if id == current_folder_id), None)
        if not folder_name:
            return jsonify({'error': 'Current folder not recognized'}), 400
        
        # Aturan pemetaan folder
        folder_mapping = {
            "01 - Pengajuan Awal": {"SR": "02A - SPV HRGA", "MR": "02B - PAMO", "GP": "02B - PAMO",
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
        
        kode_pengajuan_to_map = kode_pengajuan if folder_name == "01 - Pengajuan Awal" else filename.split()[1][:2].upper()
        
        target_folder_name = folder_mapping.get(folder_name, {}).get(kode_pengajuan_to_map)
        if target_folder_name:
            target_id = FOLDERS.get(target_folder_name)
            if target_id:
                # Pindahkan file
                service.files().update(
                    fileId=file_id,
                    addParents=target_id,
                    removeParents=current_folder_id,
                    fields='id, parents'
                ).execute()
                flash(f"File moved to {target_folder_name} successfully!", "success")
                return jsonify({'message': f"File moved to {target_folder_name} successfully!", 'new_folder_id': target_id}), 200
            else:
                flash("Target folder ID not found.", "error")
                return jsonify({'error': 'Target folder ID not found'}), 404
        else:
            flash("No valid destination found for this file.", "error")
            return jsonify({'error': 'No valid destination found for this file'}), 400

    except Exception as e:
        flash(f"Error moving file: {e}", "error")
        return jsonify({'error': f"Error moving file: {e}"}), 500


@app.route('/download/<file_id>')
def download_file(file_id):
    creds = session.get('creds')
    if not creds:
        return jsonify({'error': 'Unauthorized'}), 401

    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)

    try:
        file = service.files().get(fileId=file_id, fields='name').execute()
        filename = file.get('name')
        
        # Inisialisasi status unduhan
        DOWNLOAD_STATUS[file_id] = {'progress': 0, 'done': False, 'error': None}

        # Gunakan threading untuk unduhan di latar belakang
        thread = threading.Thread(target=perform_download, args=(creds, file_id, filename))
        thread.daemon = True
        thread.start()
        
        return jsonify({'message': 'Download started', 'file_id': file_id}), 200
    except Exception as e:
        DOWNLOAD_STATUS.get(file_id, {})['error'] = str(e)
        return jsonify({'error': f"Error starting download: {e}"}), 500


def perform_download(creds, file_id, filename):
    try:
        creds = Credentials.from_authorized_user_info(json.loads(creds.to_json()), SCOPES)
        service = build('drive', 'v3', credentials=creds)

        file_path = os.path.join(TEMP_DIR, filename)
        
        request_download = service.files().get_media(fileId=file_id)
        with open(file_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request_download, chunksize=1024 * 1024)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                progress = int(status.progress() * 100)
                DOWNLOAD_STATUS[file_id]['progress'] = progress
        
        DOWNLOAD_STATUS[file_id]['done'] = True
    except Exception as e:
        DOWNLOAD_STATUS[file_id]['error'] = str(e)


@app.route('/download_status/<file_id>')
def download_status(file_id):
    status = DOWNLOAD_STATUS.get(file_id, {'progress': 0, 'done': False, 'error': 'File not found or download not started.'})
    return jsonify(status)

@app.route('/serve_file/<file_id>')
def serve_file(file_id):
    status = DOWNLOAD_STATUS.get(file_id)
    if not status or not status['done']:
        return jsonify({'error': 'File not ready for serving'}), 400
    
    file_name = next((fn for fn, fs in DOWNLOAD_STATUS.items() if fs.get('done') and fs.get('file_id') == file_id), None)
    if not file_name:
        return jsonify({'error': 'File not found'}), 404

    # Ambil nama file dari status unduhan
    # Ini mungkin perlu diperbaiki agar lebih dinamis, tapi untuk contoh ini cukup
    file_path = os.path.join(TEMP_DIR, file_name)
    if os.path.exists(file_path):
        return send_from_directory(TEMP_DIR, file_name, as_attachment=True)
    else:
        return jsonify({'error': 'File not found on server'}), 404

@app.route('/delete_file/<file_id>')
def delete_file(file_id):
    creds = session.get('creds')
    if not creds:
        return jsonify({'error': 'Unauthorized'}), 401

    creds = Credentials.from_authorized_user_info(json.loads(creds), SCOPES)
    service = build('drive', 'v3', credentials=creds)

    try:
        service.files().delete(fileId=file_id).execute()
        return jsonify({'message': 'File deleted successfully'}), 200
    except Exception as e:
        return jsonify({'error': f"Error deleting file: {e}"}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=os.environ.get('PORT', 5000))
