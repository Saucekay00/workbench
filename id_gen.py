import sqlite3
from datetime import datetime
import pandas as pd
from flask import request, render_template, send_file, Flask, send_from_directory
import os


app = Flask(__name__)

# Ensure upload folder exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('id_gen_frontend.html')

def process_uploaded_file(file_path):
    try:
        # Read the Excel file
        data = pd.read_excel(file_path)

        # Standardize column names
        data.columns = data.columns.str.strip().str.lower()

        # Check for required columns
        if not {'name', 'email'}.issubset(data.columns):
            return "Error: Excel file must contain 'name' and 'email' columns."

        # Normalize email addresses in the uploaded data
        data['email'] = data['email'].str.strip().str.lower()

        # Process each participant
        results = []
        for _, row in data.iterrows():
            name = row['name']
            email = row['email']

            # Add participant and get their ID (existing or new)
            participant_id = new_participants(name, email)

            # Append the result (existing or new)
            results.append({'name': name, 'email': email, 'participant_id': participant_id})

        # Save results to a new Excel file
        output_file = "processed_mixed_participants.xlsx"
        output_path = os.path.join(UPLOAD_FOLDER, output_file)
        pd.DataFrame(results).to_excel(output_path, index=False)

        return output_file

    except Exception as e:
        print(f"Error processing file: {e}")
        return f"Error: Unable to process the file. Ensure it is a valid Excel file. ({e})"



@app.route('/upload', methods=['GET'])
def upload():
    if 'file' not in request.files:
        return "No file part in the request", 400

    file = request.files['file']
    if file.filename == '':
        return "No file selected for uploading", 400

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    processed_file_path = process_uploaded_file(file_path)

    return f'File processed successfully. <a href="/uploads/{processed_file_path}" download>Download Processed File</a>'


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

    if last_id:
        last_number = int(last_id[0].split('-')[-1])
        new_number = last_number + 1

    else:
        new_number = 1

    new_id = f"MMT-{current_year}-{new_number:04d}"
    return new_id

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

def export_data():
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()

    cursor.execute('SELECT participants_id, name, email FROM participants')
    rows = cursor.fetchall()
    connection.close()

    df = pd.DataFrame(rows, columns=['participants_id', 'name', 'email'])
    output_file = "all_participants.xlsx"
    df.to_excel(output_file, index=False)
    print(f"All participants data exported to: {output_file}")

init_db()

def show_table_schema():
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()
    cursor.execute("PRAGMA table_info(participants)")
    schema = cursor.fetchall()
    connection.close()
    for column in schema:
        print(column)

def bulk_uplaod(file_path):
    try:
        data = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading excel file: {e}")
        return

    if not {'name', 'email'}.issubset(data.columns):
        print("Error: excel file must contain 'name' and 'email' columns")
        return

    results = []
    for _, row in data.iterrows():
        name = row['name']
        email = row['email']

        participants_id = new_participants(name, email)
        results.append({'name': name, 'email': email, 'participant_id': participants_id})

    output_file = "processed_participants.xlsx"
    pd.DataFrame(results).to_excel(output_file, index=False)
    print(f"Processed participants data saved to {output_file}")



if __name__ == "__main__":
    app.run(debug=True, port=5000)