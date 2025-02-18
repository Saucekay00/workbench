from flask import Flask, request, render_template, send_file, jsonify, g
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
import pythoncom
import re
import win32com.client
from email.mime.base import MIMEBase
from email import encoders
import tempfile
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import img2pdf

app = Flask(__name__)
CORS(app)

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

def init_db():
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS participants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            participants_id TEXT UNIQUE NOT NULL,
            name TEXT, 
            email TEXT 
        )
    ''')
    connection.commit()
    connection.close()

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

        response = certGen.generate_certificates(workflow_data["participants"])
        if isinstance(response, BytesIO):
            return send_file(response, as_attachment=True, download_name="Certificates.zip", mimetype="application/zip")
        else:
            return response

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
            return zip_buffer  # Return ZIP containing all PDFs
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

SMTP_SERVER = "smtp.office365.com"  # Outlook SMTP server
SMTP_PORT = 587
OUTLOOK_USER = "lumi@mmt.my"  # Replace with your Outlook email
OUTLOOK_PASSWORD = "Paranskanda@33"  # Replace with your Outlook password

def send_email_smtp(sender, recipient, subject, body, attachment_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename={os.path.basename(attachment_path)}'
                )
                msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(OUTLOOK_USER, OUTLOOK_PASSWORD)
            server.send_message(msg)
            print(f"Email sent successfully to {recipient}")
            return True
    except Exception as e:
        print(f"Error sending email to {recipient}: {e}")
        return False


@app.route('/sendCertificates', methods=['POST'])
def send_certs():
    pythoncom.CoInitialize()  # Ensure COM is initialized
    print("/sendCertificates endpoint triggered!")  # Log every request

    if not workflow_data.get("participants"):
        return jsonify({"error": "No participants found. Please generate certificates first."}), 400

    email_cert_map = certGen.generate_certificates(workflow_data["participants"])

    if isinstance(email_cert_map, str) and "Error" in email_cert_map:
        return jsonify({"error": email_cert_map}), 500

    try:
        olApp = win32com.client.Dispatch("Outlook.Application")
        olNS = olApp.GetNamespace("MAPI")
    except Exception as e:
        return jsonify({"error": f"Failed to initialize Outlook: {e}"}), 500

    sender = "lumi@mmt.my"  # Replace with your Outlook email
    success_count = 0
    failure_count = 0
    seen_emails = set()  # Track emails to prevent duplicates

    for participant in workflow_data["participants"]:
        email = participant["email"]

        # Skip duplicate emails
        if email in seen_emails:
            print(f"Skipping duplicate email for: {email}")
            continue
        seen_emails.add(email)

        cert_path = email_cert_map.get(email)

        if not cert_path or not os.path.exists(cert_path):
            print(f"Warning: Certificate not found for {email} at {cert_path}")
            failure_count += 1
            continue

        try:
            # Create a temporary file for the certificate
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.abspath(os.path.join(temp_dir, os.path.basename(cert_path)))
            shutil.copy(cert_path, temp_path)  # Copy certificate to temp location

            print(f"Attaching certificate: {temp_path}")

            if not os.path.exists(temp_path):
                print(f"Error: Temporary file {temp_path} doesnot exist.")
                failure_count += 1
                continue

            mailItem = olApp.CreateItem(0)

            # Ensure correct sender account is used
            sender_account = None
            for account in olNS.Accounts:
                if hasattr(account, "SmtpAddress") and account.SmtpAddress.lower() == sender.lower():
                    sender_account = account
                    break

            if not sender_account:
                print(f"Warning: Sender account {sender} not found. Skipping {email}.")
                failure_count += 1
                continue

            mailItem.SendUsingAccount = sender_account
            mailItem.To = email
            mailItem.Subject = "Your Certificate is Ready"
            mailItem.BodyFormat = 1
            mailItem.Body = f"Dear {participant['name']},\n\nAttached is your certificate.\n\nBest regards,\nMMT Academy"

            # Attach the certificate
            print(f"Attaching: {temp_path}")
            mailItem.Attachments.Add(temp_path)

            # Add delay to avoid Outlook rate limiting
            time.sleep(3)

            retry_attempts = 3
            for attempt in range(retry_attempts):
                try:
                    mailItem.Send()
                    print(f"Email sent successfully to {email}")
                    success_count += 1
                    break
                except Exception as send_error:
                    print(f"Attempt {attempt+1} failed for {email}: {send_error}")
                    time.sleep(2)
            else:
                print(f"Failed to send email to {email} after {retry_attempts} attempts.")
                failure_count += 1

        except Exception as e:
            print(f"Error sending email to {email}: {e}")
            failure_count += 1
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)
                print(f"Temporary file {temp_path} deleted.")

        subject = "Your Certificate is Ready"
        body = f"Dear {participant['name']},\n\nAttached is your certificate.\n\nBest regards,\nMMT Academy"

        if send_email_smtp(sender, email, subject, body, cert_path):
            success_count += 1
        else:
            failure_count += 1

        return jsonify({
            "message": f"Certificates emailed successfully from {sender}.",
            "success_count": success_count,
            "failure_count": failure_count
        }), 200


@app.route('/')
def index():
    return render_template('unifiedv2_ui.html')

if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=8000)