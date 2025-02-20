from flask import Flask, request, render_template, send_file, jsonify,redirect, url_for, session, g
from werkzeug.security import generate_password_hash, check_password_hash
from io import BytesIO
import pandas as pd
import sqlite3
import os
import zipfile
import urllib.parse
import time
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
from flask_cors import CORS
from email.mime.base import MIMEBase
from email import encoders
import tempfile
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import img2pdf
from waitress import serve

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("Skipping Windows-specific modules (pythoncom, win32com.client) on Linux.")
    

app = Flask(__name__)
CORS(app)

app.secret_key = "mmtacademy"

app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

workflow_data = {
    "participants": [],  # Stores participant data with IDs and certificates
    "certificates": []   # Stores certificates with file paths
}

certificate_template = "templates/Cert_template.png"
font_path = "templates/EBGaramond-VariableFont_wght.ttf"
font_size = 100
name_position = (1400, 1160)

def is_valid_email(email):
    regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(regex, email) is not None

def clear_sessions():
    session.clear()  # Clears session when server restarts
    print("All sessions cleared on server start.")

def init_db():
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()

    # Ensure Participants Table Exists
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS participants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            participants_id TEXT UNIQUE NOT NULL,
            name TEXT,
            email TEXT
        )
    ''')

    # Ensure Users Table Exists
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            full_name TEXT NOT NULL
        )
    ''')

    # Ensure Task History Table Exists (Create if missing)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS task_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_type TEXT NOT NULL,
            participant_id TEXT,
            status TEXT NOT NULL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Check if `username` column exists in `task_history`
    cursor.execute("PRAGMA table_info(task_history)")
    columns = [col[1] for col in cursor.fetchall()]

    if "username" not in columns:
        print("Adding missing column: username in task_history table")
        cursor.execute("ALTER TABLE task_history ADD COLUMN username TEXT;")

    connection.commit()
    connection.close()

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        print(f"Login Attempt: Username = {username}, Password = {password}")

        connection = sqlite3.connect("participants.db")
        cursor = connection.cursor()
        cursor.execute('SELECT username, password, full_name FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        connection.close()

        if user:
            print(f"User Found in DB: {user[0]}")
            print(f"Stored Hash: {user[1]}")
            if check_password_hash(user[1], password):
                print("Password Matched!")
                session['username'] = user[0]
                session['full_name'] = user[2]
                return redirect(url_for('index'))  # Redirect to main page
            else:
                print("Password Did Not Match!")
        else:
            print("User Not Found in Database!")

        return "Invalid login. Please try again.", 401

    return render_template('login.html')



def add_user(username, password, full_name):
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()
    hashed_password = generate_password_hash(password)  # Hash password for security
    try:
        cursor.execute('INSERT INTO users (username, password, full_name) VALUES (?, ?, ?)',
                       (username, hashed_password, full_name))
        connection.commit()
    except sqlite3.IntegrityError:
        return "Username already exists"
    finally:
        connection.close()

@app.route('/getUser', methods=['GET'])
def get_user():
    if 'username' in session:
        return jsonify({"full_name": session['full_name']})
    return jsonify({"error": "Not logged in"}), 401

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        full_name = request.form['full_name']
        username = request.form['username']
        password = request.form['password']

        if not (full_name and username and password):
            return "All fields are required!", 400

        # Check if username already exists
        connection = sqlite3.connect("participants.db")
        cursor = connection.cursor()
        cursor.execute('SELECT username FROM users WHERE username = ?', (username,))
        existing_user = cursor.fetchone()

        if existing_user:
            connection.close()
            return "Username already exists. Please choose another one.", 400

        # Hash password and create new user
        hashed_password = generate_password_hash(password)
        cursor.execute('INSERT INTO users (username, password, full_name) VALUES (?, ?, ?)',
                       (username, hashed_password, full_name))
        connection.commit()
        connection.close()

        return redirect(url_for('login'))  # Redirect to login page after successful registration

    return render_template('register.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file uploaded'}), 400
    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'Invalid file format. Only Excel files are allowed.'}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    if os.path.exists(file_path):
        base, ext = os.path.splitext(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{base}_{int(time.time())}{ext}")

    file.save(file_path)
    workflow_data["shared_file_path"] = file_path
    return jsonify({'message': 'File uploaded successfully'}), 200

@app.route('/certificates/<filename>')
def serve_certificate(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename))

@app.route('/getParticipants', methods=['GET'])
def get_participants():
    try:
        participants = [
            {"name": p["name"], "participant_id": p["participant_id"]}
            for p in workflow_data.get("participants", [])
        ]
        if not participants:
            return jsonify({"participants": []}), 200
        return jsonify({"participants": participants}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

class idGen:
    @staticmethod
    @app.route('/idGen', methods=['POST'])
    def id_gen():
        shared_file_path = workflow_data.get("shared_file_path")
        if not shared_file_path or not os.path.exists(shared_file_path):
            return jsonify({"error": "No uploaded file found or the file does not exist. Please upload again."}), 400

        result = idGen.process_uploaded_file(shared_file_path)
        if isinstance(result, str) and result.endswith('.xlsx'):
            workflow_data["participants"] = pd.read_excel(result).to_dict(orient="records")
            print("Participants generated:", workflow_data["participants"])  # Debug log
            return send_file(result, as_attachment=True, download_name='Generated_Participants.xlsx')
        else:
            return jsonify({"error": f"Error: {result}"}), 500

    @staticmethod
    def process_uploaded_file(file_path):
        try:
            data = pd.read_excel(file_path)
            data.columns = data.columns.str.strip().str.lower()
            print("Columns in uploaded file:", data.columns)  # Debug log

            if not {'full name (as per nric)', 'email id'}.issubset(data.columns):
                return "Error: Excel file must contain 'Full name (as per NRIC)' and 'Email ID' columns."

            results = []
            for _, row in data.iterrows():
                name = row['full name (as per nric)']
                email = row['email id']
                participant_id = idGen.new_participants(name, email)
                results.append({'name': name, 'email': email, 'participant_id': participant_id})

            # Store participants in workflow_data for use in certificate generation & email sending
            workflow_data["participants"] = results

            output_file = os.path.join(app.config['UPLOAD_FOLDER'], "processed_participants.xlsx")
            pd.DataFrame(results).to_excel(output_file, index=False)

            return output_file
        except Exception as e:
            return f"Error: Unable to process the file. ({e})"

    @staticmethod
    def new_participants(name, email):
        connection = sqlite3.connect("participants.db")
        cursor = connection.cursor()
        email = email.strip().lower()
        cursor.execute('SELECT participants_id FROM participants WHERE email = ?', (email,))
        existing_id = cursor.fetchone()
        if existing_id:
            connection.close()
            return existing_id[0]

        participant_id = idGen.id_generate()
        try:
            cursor.execute('INSERT INTO participants (participants_id, name, email) VALUES (?, ?, ?)',
                           (participant_id, name, email))
            connection.commit()
            log_task("ID Generated", participant_id)  # Log Task History
            return participant_id
        except sqlite3.IntegrityError:
            return None
        finally:
            connection.close()

    @staticmethod
    def id_generate():
        connection = sqlite3.connect("participants.db")
        cursor = connection.cursor()
        current_year = datetime.now().year
        cursor.execute('SELECT participants_id FROM participants WHERE participants_id LIKE ? ORDER BY id DESC LIMIT 1',
                       (f'MMT-{current_year}-%',))
        last_id = cursor.fetchone()
        new_number = int(last_id[0].split('-')[-1]) + 1 if last_id else 1
        return f"MMT-{current_year}-{new_number:04d}"

class certGen:
    @staticmethod
    @app.route('/certGen', methods=['POST'])
    def cert_gen_route():
        if not workflow_data["participants"]:
            return "No participants found. Please generate IDs first.", 400

        result = certGen.generate_certificates(workflow_data["participants"])

        if isinstance(result, tuple):  # ✅ Ensure we received two values
            email_cert_map, zip_buffer = result
            return send_file(zip_buffer, as_attachment=True, download_name="Certificates.zip",
                             mimetype="application/zip")
        else:
            return jsonify({"error": result}), 500  # ✅ Handle error properly

    @staticmethod
    def generate_certificates(participants):
        try:
            font = ImageFont.truetype(font_path, font_size)
            zip_buffer = BytesIO()
            email_cert_map = {}  # Store participant email -> certificate path

            with zipfile.ZipFile(zip_buffer, mode='w') as zipf:
                for participant in participants:
                    img = Image.open(certificate_template)
                    draw = ImageDraw.Draw(img)
                    draw.text(name_position, participant["name"], font=font, fill="black")

                    # Save as PNG first
                    cert_png = f"certificate_{participant['participant_id']}.png"
                    cert_png_path = os.path.join(app.config['UPLOAD_FOLDER'], cert_png)
                    img.save(cert_png_path)

                    # Convert PNG to PDF
                    cert_pdf = f"certificate_{participant['participant_id']}.pdf"
                    cert_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], cert_pdf)
                    with open(cert_pdf_path, "wb") as pdf_file:
                        pdf_file.write(img2pdf.convert(cert_png_path))

                    # Map email to PDF path
                    email_cert_map[participant["email"]] = cert_pdf_path

                    # Add PDF to ZIP file
                    zipf.write(cert_pdf_path, cert_pdf)

                    # Remove temporary PNG file
                    os.remove(cert_png_path)

            zip_buffer.seek(0)

            # Return both the dictionary and the ZIP buffer
            return email_cert_map, zip_buffer
        except Exception as e:
            return f"Error: Unable to generate certificates. ({e})"


class linkGen:
    @staticmethod
    @app.route('/linkGen', methods=['POST'])
    def link_gen_post():
        participants = workflow_data.get("participants", [])
        if not participants:
            return "No participants found. Please generate IDs first.", 400

        selected_names = request.form.getlist('participants[]')
        if not selected_names:
            return "No participants selected.", 400

        print("Selected names from form:", selected_names)  # Debug log
        print("Participants in workflow_data:", participants)  # Debug log

        selected_participants = [
            p for p in participants if p["name"].strip().lower() in [name.strip().lower() for name in selected_names]
        ]

        if not selected_participants:
            return "No matching participants found.", 400

        response = linkGen.generate_links(selected_participants, request.form.get('issueYear'),
                                          request.form.get('issueMonth'))
        if isinstance(response, str) and response.endswith('.xlsx'):
            return send_file(response, as_attachment=True, download_name=os.path.basename(response))
        return response

    @staticmethod
    def generate_links(participants, issue_year, issue_month):
        try:
            links = []
            for participant in participants:
                cert_path = participant.get("cert_path", "")
                cert_url = f"http://127.0.0.1:8000/certificates/{os.path.basename(cert_path)}" if cert_path else ""
                params = {
                    'name': 'Certification',
                    'issuingOrganization': 'MMT UNIVERSAL ACADEMY SDN BHD',
                    'issueYear': issue_year,
                    'issueMonth': issue_month,
                    'certUrl': cert_url,
                    'certId': participant['participant_id']
                }
                link = f"https://www.linkedin.com/profile/add?{urllib.parse.urlencode(params)}"
                links.append({'Name': participant['name'], 'LinkedIn Link': link})
            output_file = os.path.join(app.config['UPLOAD_FOLDER'], f'LinkedIn_Links_{int(time.time())}.xlsx')
            pd.DataFrame(links).to_excel(output_file, index=False)
            return output_file
        except Exception as e:
            return f"Error: Unable to generate links. ({e})"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GMAIL_USER = "mmtlumi@gmail.com"  # Replace with your Gmail
GMAIL_PASSWORD = "xieu gldv zypm kdlg"

@app.route('/sendCertificates', methods=['POST'])
def send_certs():
    print("send_certs function called")  # Debugging
    data = request.get_json()
    event_name = data.get("event_name", "MMT Universal Academy")

    if not workflow_data.get("participants"):
        return jsonify({"error": "No participants found. Please generate certificates first."}), 400

    # Deduplicate participants
    unique_participants = {participant["email"]: participant for participant in workflow_data["participants"]}.values()
    workflow_data["participants"] = list(unique_participants)

    # Unpack the tuple returned by generate_certificates()
    email_cert_map, _ = certGen.generate_certificates(workflow_data["participants"])

    if isinstance(email_cert_map, str) and "Error" in email_cert_map:
        return jsonify({"error": email_cert_map}), 500

    success_count = 0
    failure_count = 0

    # Connect to Gmail SMTP
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # Upgrade connection to secure
        server.login(GMAIL_USER, GMAIL_PASSWORD)  # Login
    except Exception as e:
        return jsonify({"error": f"Failed to connect to Gmail: {e}"}), 500

    for participant in workflow_data["participants"]:
        email = participant["email"]
        cert_path = email_cert_map.get(email)  # ✅ Now works correctly

        if not cert_path or not os.path.exists(cert_path):
            print(f"Warning: Certificate not found for {email} at {cert_path}")
            failure_count += 1
            continue

        try:
            # Create a temporary file for the certificate
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"cert_{email}_{os.path.basename(cert_path)}")
            shutil.copy(cert_path, temp_path)  # Copy certificate to temp location

            print(f"📎 Attaching certificate: {temp_path}")

            if not os.path.exists(temp_path):
                print(f"Error: Temporary file {temp_path} does not exist.")
                failure_count += 1
                continue

            # Create the email
            msg = MIMEMultipart()
            msg["From"] = GMAIL_USER
            msg["To"] = email
            msg["Subject"] = f"Your Certificate & Participant ID for {event_name}"

            participant_id = participant.get("participant_id", "N/A")

            # Email Body
            cert_path = email_cert_map.get(email, "")
            cert_url = f"http://127.0.0.1:8000/certificates/{os.path.basename(cert_path)}" if cert_path else ""

            params = {
                'name': event_name,
                'issuingOrganization': 'MMT UNIVERSAL ACADEMY SDN BHD',
                'issueYear': datetime.now().year,
                'issueMonth': datetime.now().month,
                'certUrl': cert_url,
                'certId': participant_id
            }

            linkedin_badge_link = f"https://www.linkedin.com/profile/add?{urllib.parse.urlencode(params)}"
            body = f"""Dear {participant['name']},

            Congratulations on successfully completing {event_name}!

            We’re delighted to recognize your achievement. Below are your completion details:

            Participant ID: {participant_id}
            
            LinkedIn Badge Link: [Click here to add your certification to LinkedIn]({linkedin_badge_link}) 

            Please find your attached certificate. Keep this for your records.

            Best regards,  
            MMT Universal Academy  
            """
            msg.attach(MIMEText(body, "plain"))

            # Attach Certificate
            with open(temp_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(cert_path)}")
                msg.attach(part)

            # Send Email
            server.sendmail(GMAIL_USER, email, msg.as_string())
            print(f"Email sent successfully to {email}")
            success_count += 1

            log_task("Email Sent", participant["participant_id"])

        except Exception as e:
            print(f"Error sending email to {email}: {e}")
            failure_count += 1
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)
                print(f"Temporary file {temp_path} deleted.")

    server.quit()  # Close the SMTP server connection

    return jsonify({
        "message": f"Certificates emailed successfully from {GMAIL_USER}.",
        "success_count": success_count,
        "failure_count": failure_count
    }), 200


def log_task(task_type, participant_id=None, status="Completed"):
    username = session.get('username', 'Unknown')  # Get logged-in user safely
    print(f"Logging Task: {task_type}, Participant ID: {participant_id}, User: {username}, Status: {status}")  # Debugging

    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()
    cursor.execute('''
        INSERT INTO task_history (task_type, participant_id, username, status)
        VALUES (?, ?, ?, ?)
    ''', (task_type, participant_id, username, status))
    connection.commit()
    connection.close()




@app.route('/taskHistory', methods=['GET'])
def get_task_history():
    if 'username' not in session:
        return jsonify({"error": "Unauthorized"}), 401

    username = session['username']
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()
    cursor.execute('''
        SELECT task_type, participant_id, status, timestamp FROM task_history
        WHERE username = ?
        ORDER BY timestamp DESC
    ''', (username,))

    history = [
        {"task_type": row[0], "participant_id": row[1], "status": row[2], "timestamp": row[3]}
        for row in cursor.fetchall()
    ]
    connection.close()

    return jsonify({"task_history": history or []})  # ✅ Return empty list if no data


@app.route('/previewCertificate/<participant_id>', methods=['GET'])
def preview_certificate(participant_id):
    # Find the certificate path
    cert_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"certificate_{participant_id}.pdf")

    if not os.path.exists(cert_pdf_path):
        return "Certificate not found", 404

    return send_file(cert_pdf_path, mimetype='application/pdf')

@app.route('/logout')
def logout():
    session.pop('username', None)
    session.pop('full_name', None)
    return redirect(url_for('login'))


@app.route('/')
def index():
    print(f"Session Data: {session}")  # Print session data
    if 'username' not in session:
        print("No active session, redirecting to login.")
        return redirect(url_for('login'))  # Force login

    print(f"Logged in as: {session['username']}")
    return render_template('unifiedv2_ui.html')


if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=8080)
    #serve(app, host="0.0.0.0", port=8080)
