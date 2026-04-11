import os
import uuid
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, jsonify, send_file
)
import pandas as pd
import io
import time
import random
import logging
import sqlite3
import json

# ---------- configuration & logging ----------
app = Flask(__name__)
app.secret_key = 'admission-portal-secret-key-2026'

# Configure logging
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ---------- constants ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, 'admissions.db')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
# On Render, secret files are stored at /etc/secrets/
# Fall back to local path for development
CREDENTIALS_FILE = '/etc/secrets/credentials.json' if os.path.exists('/etc/secrets/credentials.json') else os.path.join(BASE_DIR, 'credentials.json')
SHEET_NAME = "SSEC_ADMISSION DATABASE_2026-27"  # User can change this name

os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------- SQLite Database Setup ----------
def init_db():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        # We store the entire row as a JSON string for simplicity in the fallback, 
        # or we can create columns. Let's create columns to match COLUMNS constant.
        columns_sql = ", ".join([f'"{col}" TEXT' for col in COLUMNS])
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS admissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columns_sql},
                synced INTEGER DEFAULT 0,
                last_error TEXT
            )
        ''')
        # Ensure ApplicationNumber has an index for fast lookups
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_app_num ON admissions (ApplicationNumber)')
        conn.commit()
        conn.close()
        logger.info("SQLite database initialized successfully.")
    except Exception as e:
        logger.error(f"Failed to initialize SQLite: {e}")

# ---------- Google Sheets Client Setup ----------
COLUMNS = [
    'ApplicationNumber', 'AdmissionDate', 'HostelDayscholar', 'Name', 'Degree', 
    'Regular / Lateral', 'DOB', 'Preference1', 'Preference2', 
    'Preference3', 'Community', 'Quota', 'Gender', 'FirstGraduate', 'AdmissionMode', 'Scholarship',
    'Phone', 'FatherName', 'FatherMobile', 'MotherName', 'MotherMobile', 
    'Aadhar', 'Address', 'District', 'State', 'Reference', 'StaffName',
    'PaymentStatus', 'Initial Amount', 'Tuition Fee',
    'Physics', 'Chemistry', 'Maths', 'Cutoff',
    'Certificate10th', 'Certificate11th', 'Certificate12th',
    'TransferCertificate', 'CommunityCertificate', 'IncomeCertificate',
    'NativityCertificate', 'AadharXerox',
    'BankPassbook', 'StudentPhoto',
    'CreatedAt', 'UpdatedAt'
]

init_db()

# ---------- Google Sheets Client Setup ----------
_cached_sheet = None

def get_col_letter(n):
    """Convert index to A-Z/AA-ZZ"""
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def get_worksheet():
    """Get the specific worksheet with global caching to improve performance."""
    global _cached_sheet
    if _cached_sheet is not None:
        return _cached_sheet
    
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(f"Credentials file not found at {CREDENTIALS_FILE}.")
    
    max_retries = 3
    retry_delay = 2 # initial delay in seconds
    
    for attempt in range(max_retries):
        try:
            scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
            client = gspread.authorize(creds)
            
            try:
                spreadsheet = client.open(SHEET_NAME)
                _cached_sheet = spreadsheet.sheet1
                
                # STRICTOR HEADER CHECK: Ensure headers match COLUMNS exactly
                headers = _cached_sheet.row_values(1)
                
                # If sheet is empty or headers are missing/misaligned/outdated
                if not headers or headers != COLUMNS:
                    print("DEBUG: Sheet headers missing, misaligned, or outdated. Syncing row 1...")
                    # Update row 1 with current COLUMNS
                    range_str = f'A1:{get_col_letter(len(COLUMNS))}1'
                    _cached_sheet.update(range_name=range_str, values=[COLUMNS])
                
            except gspread.exceptions.SpreadsheetNotFound:
                # Create the sheet if it doesn't exist
                spreadsheet = client.create(SHEET_NAME)
                _cached_sheet = spreadsheet.sheet1
                _cached_sheet.append_row(COLUMNS)
            
            return _cached_sheet
            
        except Exception as e:
            logger.error(f"Attempt {attempt + 1}/{max_retries} failed to initialize Google Sheets: {e}")
            if attempt < max_retries - 1:
                sleep_time = retry_delay * (2 ** attempt) + (random.randint(0, 1000) / 1000)
                logger.info(f"Retrying in {sleep_time:.2f} seconds...")
                time.sleep(sleep_time)
            else:
                logger.error("All connection attempts failed.")
                raise e

# ---------- Local Storage & Sync Helpers ----------

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def save_locally(row_data):
    """Saves a single record to SQLite."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Build SQL
        cols_str = ", ".join([f'"{col}"' for col in COLUMNS])
        placeholders = ", ".join(["?" for _ in COLUMNS])
        
        # Check if record exists
        app_num = row_data[0]
        cursor.execute('SELECT id FROM admissions WHERE ApplicationNumber = ?', (app_num,))
        existing = cursor.fetchone()
        
        if existing:
            # Update existing record
            set_clause = ", ".join([f'"{col}" = ?' for col in COLUMNS])
            cursor.execute(f'UPDATE admissions SET {set_clause}, synced = 0 WHERE ApplicationNumber = ?', (*row_data, app_num))
            logger.info(f"Updated record locally: {app_num}")
        else:
            # Insert new record
            cursor.execute(f'INSERT INTO admissions ({cols_str}, synced) VALUES ({placeholders}, 0)', (*row_data,))
            logger.info(f"Saved new record locally: {app_num}")
            
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        logger.error(f"Local save failed for {row_data[0]}: {e}")
        return False

def sync_to_sheets(app_number, row_data, is_new=True):
    """Attempts to sync a single record to Google Sheets."""
    try:
        sheet = get_worksheet()
        if is_new:
            sheet.append_row(row_data)
        else:
            # For edits, find and update
            cell = sheet.find(app_number)
            if cell:
                row_idx = cell.row
                update_range = f'A{row_idx}:{get_col_letter(len(COLUMNS))}{row_idx}'
                sheet.update(range_name=update_range, values=[row_data])
            else:
                # If not found in sheet, append it instead
                sheet.append_row(row_data)
        
        # Mark as synced in SQLite
        conn = get_db_connection()
        conn.execute('UPDATE admissions SET synced = 1, last_error = NULL WHERE ApplicationNumber = ?', (app_number,))
        conn.commit()
        conn.close()
        logger.info(f"Successfully synced {app_number} to Google Sheets.")
        return True
    except Exception as e:
        logger.error(f"Sync failed for {app_number}: {e}")
        conn = get_db_connection()
        conn.execute('UPDATE admissions SET last_error = ? WHERE ApplicationNumber = ?', (str(e), app_number))
        conn.commit()
        conn.close()
        return False

def sync_all_pending():
    """Retries syncing all unsynced records."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM admissions WHERE synced = 0')
        pending = cursor.fetchall()
        conn.close()
        
        if not pending:
            return 0
            
        logger.info(f"Found {len(pending)} pending records to sync.")
        count = 0
        for row in pending:
            # Reconstruct row data from sqlite Row
            row_data = [row[col] for col in COLUMNS]
            # Since we don't strictly know if it was an edit or new, 
            # we check if it exists in the sheet first in sync_to_sheets (handled there)
            if sync_to_sheets(row['ApplicationNumber'], row_data, is_new=False):
                count += 1
                # Small sleep to avoid rate limiting
                time.sleep(1)
        return count
    except Exception as e:
        logger.error(f"Error in sync_all_pending: {e}")
        return 0

# ===================== ROUTES =====================

@app.route('/')
def intro():
    return render_template('intro.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if username == 'admin' and password == 'Admission@123':
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('dashboard'))
        flash('Invalid username or password', 'error')
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    return render_template('dashboard.html', total=0, paid=0, unpaid=0)

@app.route('/api/stats')
def api_stats():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401
    
    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        
        total = 0
        paid = 0
        
        for r in records:
            app_n = str(r.get('ApplicationNumber', ''))
            if not app_n or app_n == 'ApplicationNumber':
                continue
            
            total += 1
            
            if str(r.get('PaymentStatus', '')).strip().lower() == 'paid':
                paid += 1
                
        unpaid = total - paid
        
        return jsonify({
            'success': True,
            'total': total,
            'paid': paid,
            'unpaid': unpaid
        })
    except Exception as e:
        print(f"API STATS ERROR: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/new-applicant', methods=['GET', 'POST'])
def new_applicant():
    if not session.get('logged_in'):
        if request.method == 'POST':
            return jsonify({'success': False, 'message': 'Session expired'}), 401
        return redirect(url_for('login'))

    if request.method == 'POST':
        try:
            app_number = 'APP-' + uuid.uuid4().hex[:8].upper()
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # --- Duplicate Prevention Logic ---
            name = request.form.get('name', '').strip()
            dob = request.form.get('dob', '').strip()
            phone = request.form.get('phone', '').strip()
            
            # Check SQLite for recent submission with same Name/DOB/Phone
            # (within last 2 minutes)
            try:
                conn = get_db_connection()
                # Use CreatedAt for time check
                two_mins_ago = (datetime.now() - timedelta(minutes=2)).strftime('%Y-%m-%d %H:%M:%S')
                dup = conn.execute('''
                    SELECT ApplicationNumber FROM admissions 
                    WHERE Name = ? AND DOB = ? AND Phone = ? 
                    AND CreatedAt > ?
                ''', (name, dob, phone, two_mins_ago)).fetchone()
                conn.close()
                
                if dup:
                    logger.warning(f"Duplicate submission blocked for {name} ({phone}). Existing App: {dup['ApplicationNumber']}")
                    return jsonify({
                        'success': False, 
                        'message': 'Duplicate submission detected. This application was already submitted a moment ago.'
                    }), 409
            except Exception as e:
                logger.error(f"Duplicate check failed: {e}")
                # We continue even if check fails to avoid blocking legitimate users
            # ----------------------------------

            row_data = [
                app_number,
                request.form.get('admission_date', ''),
                request.form.get('hostel_dayscholar', ''),
                request.form.get('name', ''),
                request.form.get('degree', ''),
                request.form.get('regulation', ''),
                request.form.get('dob', ''),
                request.form.get('preference1', ''),
                request.form.get('preference2', ''),
                request.form.get('preference3', ''),
                request.form.get('community', ''),
                request.form.get('quota', ''),
                request.form.get('gender', ''),
                request.form.get('firstGraduate', 'No'), # New field
                request.form.get('admission_mode', 'Select Mode'),
                request.form.get('scholarship', ''), 
                request.form.get('phone', ''),
                request.form.get('father_name', ''),
                request.form.get('father_mobile', ''),
                request.form.get('mother_name', ''),
                request.form.get('mother_mobile', ''),
                request.form.get('aadhar', ''),
                request.form.get('address', ''),
                request.form.get('district', ''),
                request.form.get('state', ''),
                request.form.get('reference', ''),
                request.form.get('staffName', ''), # New field
                request.form.get('payment_status', ''),
                request.form.get('initial_amount', ''),
                request.form.get('Tuition_fee', ''),
                request.form.get('physics', ''),
                request.form.get('chemistry', ''),
                request.form.get('maths', ''),
                request.form.get('cutoff', ''),
                request.form.get('certificate10th', 'No'),
                request.form.get('certificate11th', 'No'),
                request.form.get('certificate12th', 'No'),
                request.form.get('transferCertificate', 'No'),
                request.form.get('communityCertificate', 'No'),
                request.form.get('incomeCertificate', 'No'),
                request.form.get('nativityCertificate', 'No'),
                request.form.get('aadharXerox', 'No'),
                request.form.get('bankPassbook', 'No'),
                request.form.get('studentPhoto', 'No'),
                now,
                now
            ]

            logger.info(f"DEBUG: Processing submission for {app_number} - {request.form.get('name')}")
            
            # 1. Save locally FIRST (Critical for reliability)
            local_success = save_locally(row_data)
            
            # 2. Attempt Sync to Google Sheets
            sync_success = sync_to_sheets(app_number, row_data, is_new=True)
            
            if not local_success and not sync_success:
                logger.error(f"FATAL: Could not save data locally OR to Google Sheets for {app_number}")
                return jsonify({'success': False, 'message': 'Critical system error: Could not save data.'}), 500

            return jsonify({
                'success': True, 
                'application_number': app_number,
                'sync_status': 'success' if sync_success else 'pending'
            })
        except Exception as e:
            logger.error(f"ERROR in new_applicant POST: {e}")
            return jsonify({'success': False, 'message': str(e)}), 500

    return render_template('new_applicant.html')

@app.route('/edit/<app_number>', methods=['GET', 'POST'])
def edit_applicant(app_number):
    if not session.get('logged_in'):
        if request.method == 'POST':
            return jsonify({'success': False, 'message': 'Session expired'}), 401
        return redirect(url_for('login'))

    sheet = get_worksheet()
    
    if request.method == 'POST':
        try:
            # Find the row by ApplicationNumber
            cell = sheet.find(app_number)
            if not cell:
                return jsonify({'success': False, 'message': 'Record not found'}), 404
            
            row_idx = cell.row
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # We update all columns except ApplicationNumber (index 0) and CreatedAt (index 42)
            # Row values in gspread are 1-based, sheet row 1 is header
            # To update precisely, we get the existing row values to preserve what we don't change
            existing_row = sheet.row_values(row_idx)
            
            # Map form fields to COLUMNS
            # Note: scholarship mapping needs care
            scholarship_val = request.form.get('scholarship', '')
            
            updated_values = []
            for i, col_name in enumerate(COLUMNS):
                if col_name == 'ApplicationNumber':
                    updated_values.append(app_number)
                elif col_name == 'UpdatedAt':
                    updated_values.append(now)
                elif col_name == 'CreatedAt':
                    # Keep existing CreatedAt
                    val = existing_row[i] if i < len(existing_row) else now
                    updated_values.append(val)
                elif col_name == 'Scholarship':
                    updated_values.append(scholarship_val)
                else:
                    # Generic mapping from form
                    # Many form fields match COLUMNS exactly (case-insensitive or camelCase)
                    # We check common mappings
                    form_key = col_name.lower()
                    if col_name == 'HostelDayscholar': form_key = 'hostel_dayscholar'
                    elif col_name == 'AdmissionDate': form_key = 'admission_date'
                    elif col_name == 'Regular / Lateral': form_key = 'regulation'
                    elif col_name == 'FatherName': form_key = 'father_name'
                    elif col_name == 'FatherMobile': form_key = 'father_mobile'
                    elif col_name == 'MotherName': form_key = 'mother_name'
                    elif col_name == 'MotherMobile': form_key = 'mother_mobile'
                    elif col_name == 'Initial Amount': form_key = 'initial_amount'
                    elif col_name == 'Tuition Fee': form_key = 'Tuition_fee'
                    elif col_name == 'PaymentStatus': form_key = 'payment_status'
                    elif col_name == 'FirstGraduate': form_key = 'firstGraduate'
                    elif col_name == 'StaffName': form_key = 'staffName'
                    elif col_name == 'AdmissionMode': form_key = 'admission_mode'
                    
                    val = request.form.get(form_key, existing_row[i] if i < len(existing_row) else '')
                    updated_values.append(val)

            # Update the range
            # 1. Save locally FIRST
            local_success = save_locally(updated_values)
            
            # 2. Attempt Sync
            sync_success = sync_to_sheets(app_number, updated_values, is_new=False)
            
            return jsonify({
                'success': True, 
                'application_number': app_number,
                'sync_status': 'success' if sync_success else 'pending'
            })
        except Exception as e:
            logger.error(f"ERROR in edit_applicant POST: {e}")
            return jsonify({'success': False, 'message': str(e)}), 500

    # GET: Fetch record
    try:
        records = sheet.get_all_records()
        record = next((r for r in records if str(r.get('ApplicationNumber')) == app_number), None)
        if not record:
            flash("Applicant not found.")
            return redirect(url_for('existing_applicant'))
        
        return render_template('edit_applicant.html', record=record)
    except Exception as e:
        logger.error(f"ERROR in edit_applicant GET: {e}")
        flash("Error loading applicant data.")
        return redirect(url_for('existing_applicant'))

@app.route('/admin/sync', methods=['GET', 'POST'])
def admin_sync():
    """Manual sync trigger for administrators."""
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    if request.method == 'POST':
        synced_count = sync_all_pending()
        flash(f"Sync complete. {synced_count} records pushed to Google Sheets.", "success")
        return redirect(url_for('admin_sync'))
        
    # Get pending count
    try:
        conn = get_db_connection()
        pending = conn.execute('SELECT COUNT(*) as count FROM admissions WHERE synced = 0').fetchone()['count']
        conn.close()
    except:
        pending = 0
        
    return render_template('sync_status.html', pending_count=pending)

@app.route('/existing-applicant')
def existing_applicant():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    search_query = request.args.get('search', '').strip().lower()
    
    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        records = [r for r in records if str(r.get('ApplicationNumber')) != 'ApplicationNumber']
        
        if search_query:
            records = [
                r for r in records
                if search_query in str(r.get('ApplicationNumber', '')).lower() or
                   search_query in str(r.get('Name', '')).lower() or
                   search_query in str(r.get('AdmissionDate', '')).lower()
            ]
            
        return render_template('existing_applicant.html', records=records)
    except Exception as e:
        print(f"EXISTING APPLICANT ERROR: {e}")
        return render_template('existing_applicant.html', records=[])

@app.route('/api/all')
def api_all():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401
    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        records = [r for r in records if str(r.get('ApplicationNumber')) != 'ApplicationNumber']
        return jsonify({'records': records})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/search/<app_number>')
def api_search(app_number):
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401
    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        records = [r for r in records if str(r.get('ApplicationNumber')) != 'ApplicationNumber']
        match = next((r for r in records if str(r.get('ApplicationNumber')) == app_number), None)
        if not match:
            return jsonify({'found': False})
        return jsonify({'found': True, 'data': match})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/search_query', methods=['POST'])
def api_search_query():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401
    
    query = request.json.get('query', '').strip().lower()
    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        records = [r for r in records if str(r.get('ApplicationNumber')) != 'ApplicationNumber']
        results = [
            r for r in records 
            if query in str(r.get('ApplicationNumber', '')).lower() or 
               query in str(r.get('Name', '')).lower() or
               query in str(r.get('AdmissionDate', '')).lower()
        ]
        return jsonify({'records': results})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/toggle-status', methods=['POST'])
def api_toggle_status():
    if not session.get('logged_in'):
        return jsonify({'success': False, 'message': 'Unauthorized'}), 401
    
    payload = request.get_json()
    app_number = payload.get('ApplicationNumber')
    
    if not app_number:
        return jsonify({'success': False, 'message': 'Missing Application Number'}), 400
        
    try:
        sheet = get_worksheet()
        # Ensure we find the exact ApplicationNumber
        # Find all cells and filter to match the first column (ApplicationNumber)
        cell = sheet.find(app_number)
        if not cell or cell.col != 1:
            # Re-verify in case application number appears in other columns
            # (Though uuid should be unique, better to be safe)
            cells = sheet.findall(app_number)
            cell = next((c for c in cells if c.col == 1), None)
            
        if not cell:
            print(f"DEBUG: Toggle failed. Application {app_number} not found.")
            return jsonify({'success': False, 'message': 'Record not found'}), 404
            
        row_idx = cell.row
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # PaymentStatus index in list is 24, so column is 25
        status_col = COLUMNS.index('PaymentStatus') + 1
        updated_at_col = COLUMNS.index('UpdatedAt') + 1
        
        current_status = str(sheet.cell(row_idx, status_col).value).strip()
        print(f"DEBUG: Row {row_idx} matched for {app_number}. Current status: '{current_status}'")
        
        # Normalize and toggle
        is_paid = current_status.lower() == 'paid'
        new_status = 'Unpaid' if is_paid else 'Paid'
        
        # Update both status and UpdatedAt
        sheet.update_cell(row_idx, status_col, new_status)
        sheet.update_cell(row_idx, updated_at_col, now)
        
        print(f"DEBUG: Toggled {app_number} from {current_status} to {new_status}")
        
        return jsonify({
            'success': True, 
            'new_status': new_status,
            'message': f'Status toggled to {new_status}'
        })
    except Exception as e:
        print(f"TOGGLE ERROR: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/update', methods=['POST'])
def api_update():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401

    payload = request.get_json()
    app_number = payload.get('ApplicationNumber')
    if not app_number:
        return jsonify({'success': False, 'message': 'Missing Application Number'}), 400

    try:
        sheet = get_worksheet()
        cell = sheet.find(app_number)
        if not cell:
            return jsonify({'success': False, 'message': 'Record not found'}), 404
        
        row_idx = cell.row
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        row_values = sheet.row_values(row_idx)

        # Update columns from index 1 (AdmissionDate) to index 30 (UpdatedAt)
        # last_col_letter calculation
        num_cols = len(COLUMNS)
            
        update_range = f'B{row_idx}:{get_col_letter(num_cols)}{row_idx}'
        
        updated_row_section = []
        for i in range(1, num_cols):
            col_name = COLUMNS[i]
            if col_name == 'UpdatedAt':
                val = now
            else:
                val = payload.get(col_name, row_values[i] if i < len(row_values) else '')
            updated_row_section.append(val)

        sheet.update(range_name=update_range, values=[updated_row_section])

        return jsonify({'success': True, 'message': 'Record updated successfully'})
    except Exception as e:
        print(f"UPDATE ERROR: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/delete', methods=['POST'])
def api_delete():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401

    payload = request.get_json()
    app_number = payload.get('ApplicationNumber')
    if not app_number:
        return jsonify({'success': False, 'message': 'Missing Application Number'}), 400

    try:
        sheet = get_worksheet()
        cell = sheet.find(app_number)
        if not cell:
            return jsonify({'success': False, 'message': 'Record not found'}), 404
        
        row_idx = cell.row
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # No file deletion needed since certificates are now boolean indicators


        sheet.delete_rows(row_idx)
        return jsonify({'success': True, 'message': 'Record deleted successfully'})
    except Exception as e:
        print(f"DELETE ERROR: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/analytics', methods=['GET', 'POST'])
def analytics():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    # Check if a date is provided, otherwise default to None (for full data)
    selected_date = request.form.get('query_date')
    
    try:
        # We handle calculations in JavaScript via api/analytics-details
        # but keep selected_date for template pre-filling
        return render_template('analytics.html', selected_date=selected_date)
    except Exception as e:
        print(f"ANALYTICS ERROR: {e}")
        return render_template('analytics.html', selected_date=selected_date)

@app.route('/api/analytics-details')
def api_analytics_details():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401
    
    filter_date = request.args.get('date') # can be None or empty string

    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        records = [r for r in records if str(r.get('ApplicationNumber')) != 'ApplicationNumber']
        
        # Filter if date is provided
        filtered_records = [
            r for r in records 
            if not filter_date or str(r.get('AdmissionDate', '')).strip() == filter_date
        ]

        counts = {
            'total': len(filtered_records),
            'ug_regular': 0,
            'ug_lateral': 0,
            'pg_total': 0,
            'diploma_regular': 0,
            'diploma_lateral': 0
        }

        for r in filtered_records:
            deg = str(r.get('Degree', '')).strip()
            reg = str(r.get('Regular / Lateral', '')).strip()

            if deg == 'UG':
                if reg == 'Regular': counts['ug_regular'] += 1
                elif reg == 'Lateral': counts['ug_lateral'] += 1
            elif deg == 'PG':
                counts['pg_total'] += 1
            elif deg == 'Diploma':
                if reg == 'Regular': counts['diploma_regular'] += 1
                elif reg == 'Lateral': counts['diploma_lateral'] += 1

        return jsonify({
            'success': True,
            'date': filter_date or "Full Data",
            'counts': counts
        })
    except Exception as e:
        print(f"ANALYTICS DETAILS ERROR: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500
        print(f"ANALYTICS DETAILS ERROR: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/export-date-wise')
def export_date_wise():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    selected_date = request.args.get("date")
    if not selected_date:
        flash("Please select a date for export.", "warning")
        return redirect(url_for('analytics'))

    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        if not records:
             flash("No data available to export.", "warning")
             return redirect(url_for('analytics'))
             
        df = pd.DataFrame(records)
        df = df[df['AdmissionDate'].astype(str) == selected_date]
        
        if df.empty:
             flash(f"No records found for date: {selected_date}", "info")
             return redirect(url_for('analytics'))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Admissions')
        output.seek(0)

        filename = f"admission_data_{selected_date}.xlsx"
        return send_file(
            output, 
            as_attachment=True, 
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        print(f"DATE EXPORT ERROR: {e}")
        flash(f"Export failed: {e}", "error")
        return redirect(url_for('analytics'))

@app.route('/export-all')
def export_all():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    try:
        sheet = get_worksheet()
        records = sheet.get_all_records()
        if not records:
             flash("No data available to export.", "warning")
             return redirect(url_for('analytics'))
             
        df = pd.DataFrame(records)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Admissions Full')
        output.seek(0)

        filename = "admission_full_data.xlsx"
        return send_file(
            output, 
            as_attachment=True, 
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        print(f"FULL EXPORT ERROR: {e}")
        flash(f"Export failed: {e}", "error")
        return redirect(url_for('analytics'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    # Pre-heat the connection in a background thread
    # to ensure the first request is fast, without blocking app startup.
    import threading
    def pre_heat():
        try:
            print("Pre-heating Google Sheets connection...")
            get_worksheet()
            print("Google Sheets connected and ready.")
        except Exception as e:
            print(f"Warning: Could not pre-heat Google Sheets: {e}")
            
    threading.Thread(target=pre_heat, daemon=True).start()
    app.run(host="0.0.0.0", port=5000, debug=True)
