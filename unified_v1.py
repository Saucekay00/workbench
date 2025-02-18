from flask import Flask, request, render_template, send_file, jsonify
from io import BytesIO
import pandas as pd
import sqlite3
import os
import zipfile
import urllib.parse
import time
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Certificate generation settings
certificate_template = "templates/Cert_template.png"
font_path = "templates/EBGaramond-VariableFont_wght.ttf"
font_size = 100
name_position = (1400, 1160)

# Database initialization
def init_db():
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()

    cursor.execute('DROP TABLE IF EXISTS participants')

    cursor.execute('''

        CREATE TABLE IF NOT EXISTS participants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            participants_id TEXT UNIQUE NOT NULL,
            name TEXT, 
            email TEXT 
        )
    ''')
    connection.commit()
    return connection

# ID generation helper
def id_generate():
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()
    current_year = datetime.now().year
    cursor.execute('''
        SELECT participants_id
        FROM participants
        WHERE participants_id LIKE ?
        ORDER BY id DESC LIMIT 1
    ''', (f'MMT-{current_year}-%',))
    last_id = cursor.fetchone()
    new_number = int(last_id[0].split('-')[-1]) + 1 if last_id else 1
    new_id = f"MMT-{current_year}-{new_number:04d}"
    connection.close()
    return new_id

# Participant addition
def new_participants(name, email):
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()

    # Normalize email
    email = email.strip().lower()

    # Check if the participant already exists
    cursor.execute('''
            SELECT participants_id 
            FROM participants 
            WHERE email = ?
        ''', (email,))
    existing_id = cursor.fetchone()

    if existing_id:
        print(f"Participant already exists: {name} ({email}), ID: {existing_id[0]}")
        connection.close()
        return existing_id[0]

    # If the participant doesn't exist, generate a new ID
    participant_id = id_generate()

    try:
        # Insert new participant into the database
        cursor.execute('''
                INSERT INTO participants (participants_id, name, email)
                VALUES (?, ?, ?)
            ''', (participant_id, name, email))
        connection.commit()
        print(f"New participant added: {name} ({email}), ID: {participant_id}")
        return participant_id
    except sqlite3.IntegrityError:
        print("Error: Could not add participant.")
        return None
    finally:
        connection.close()

# Certificate generation route
@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    file = request.files.get('file')
    if not file or not file.filename.endswith('.xlsx'):
        return "Invalid file. Please upload an Excel file.", 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    try:
        df = pd.read_excel(file_path)
        if 'Full name (as per NRIC)' not in df.columns:
            return "Error: Required column not found in Excel file.", 400

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            font = ImageFont.truetype(font_path, font_size)
            for _, row in df.iterrows():
                name = row['Full name (as per NRIC)'].strip()
                img = Image.open(certificate_template)
                draw = ImageDraw.Draw(img)
                draw.text(name_position, name, font=font, fill="black")
                cert_filename = f"certificate_{name}.png"
                cert_path = os.path.join(app.config['UPLOAD_FOLDER'], cert_filename)
                img.save(cert_path)
                zipf.write(cert_path, cert_filename)
        zip_buffer.seek(0)
        return send_file(zip_buffer, as_attachment=True, download_name="Certificates.zip", mimetype="application/zip")
    except Exception as e:
        return str(e), 500

# ID generation route
@app.route('/generate_ids', methods=['POST'])
def generate_ids():
    file = request.files.get('file')
    if not file or not file.filename.endswith('.xlsx'):
        return "Invalid file. Please upload an Excel file.", 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    try:
        df = pd.read_excel(file_path)
        if not {'name', 'email'}.issubset(df.columns):
            return "Error: Required columns not found in Excel file.", 400

        results = []
        for _, row in df.iterrows():
            name = row['name']
            email = row['email']

            # Add participant and get their ID (existing or new)
            participant_id = new_participants(name, email)

            # Append the result (existing or new)
            results.append({'name': name, 'email': email, 'participant_id': participant_id})

        output_file = os.path.join(app.config['UPLOAD_FOLDER'], "processed_participants.xlsx")
        pd.DataFrame(results).to_excel(output_file, index=False)
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return str(e), 500

# LinkedIn link generation route
@app.route('/generate_links', methods=['POST'])
def generate_links():
    file = request.files.get('file')
    issue_year = request.form.get('issueYear')
    issue_month = request.form.get('issueMonth')
    if not file or not file.filename.endswith('.xlsx') or not issue_year or not issue_month:
        return "Invalid input. Ensure file and required form data are provided.", 400

    try:
        df = pd.read_excel(file)
        if not {'name', 'certId'}.issubset(df.columns):
            return "Error: Required columns not found in Excel file.", 400

        df['LinkedIn Link'] = df.apply(lambda row: (
            f"https://www.linkedin.com/profile/add?"
            f"name=Certification&issuingOrganization=MMT UNIVERSAL ACADEMY&issueYear={issue_year}"
            f"&issueMonth={issue_month}&certId={row['certId']}&certUrl="
            f"https://certificates.com/{row['name'].replace(' ', '_')}_{issue_year}_{issue_month}.pdf"
        ), axis=1)

        output_file = os.path.join(app.config['UPLOAD_FOLDER'], "participants_with_links.xlsx")
        df.to_excel(output_file, index=False)
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return str(e), 500

# Home route
@app.route('/')
def index():
    return render_template('unified_ui.html')

if __name__ == '__main__':
    app.run(debug=True, port=8000)
