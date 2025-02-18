from flask import Flask, request, send_file, jsonify
import pandas as pd
import urllib.parse
import os
import time
from flask import request, render_template, send_file, Flask, send_from_directory

app = Flask(__name__)

# Function to simulate certificate generation
def generate_certificate(participant):
    # Replace this with your actual certificate generation logic
    cert_url = f"https://your-certificate-generator.com/certificates/{participant['name'].replace(' ', '_')}_{participant['issueYear']}_{participant['issueMonth']}.pdf"
    return cert_url

@app.route('/')
def index():
    return render_template('link_gen.html')

# Function to generate LinkedIn links
def generate_link(participant):
    cert_url = generate_certificate(participant)
    encoded_cert_url = urllib.parse.quote(cert_url)

    params = {
        'name': urllib.parse.quote("Test Certification"),
        'issuingOrganization': 'MMT UNIVERSAL ACADEMY SDN BHD',
        'issueYear': str(participant['issueYear']),
        'issueMonth': str(participant['issueMonth']),
        'certUrl': encoded_cert_url,
        'certId': str(participant['certId']),
        'organizationName': 'MMT UNIVERSAL ACADEMY SDN BHD'
    }

    encoded_params = '&'.join([f"{key}={urllib.parse.quote(str(value))}" for key, value in params.items()])
    full_url = f"https://www.linkedin.com/profile/add?startTask=CERTIFICATION_NAME&{encoded_params}"
    return full_url

# Flask endpoint to handle file uploads and process links
@app.route('/generate', methods=['GET'])
def generate_links():
    try:
        # Check if file is uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected for uploading'}), 400

        # Read issueYear and issueMonth from the form data
        issue_year = request.form.get('issueYear')
        issue_month = request.form.get('issueMonth')

        if not issue_year or not issue_month:
            return jsonify({'error': 'Issue year and month are required'}), 400

        # Read the Excel file
        df = pd.read_excel(file)

        # Validate required columns
        required_columns = ['name', 'certId']
        for col in required_columns:
            if col not in df.columns:
                return jsonify({'error': f'Missing required column: {col}'}), 400

        # Generate LinkedIn Links
        def generate_link(row):
            cert_url = f"https://certificates.com/{row['name'].replace(' ', '_')}_{issue_year}_{issue_month}.pdf"
            params = {
                'name': 'Test Certification',
                'issuingOrganization': 'MMT UNIVERSAL ACADEMY SDN BHD',
                'issueYear': issue_year,
                'issueMonth': issue_month,
                'certUrl': cert_url,
                'certId': row['certId']
            }
            return f"https://www.linkedin.com/profile/add?{urllib.parse.urlencode(params)}"

        df['LinkedIn Link'] = df.apply(generate_link, axis=1)

        # Save the updated file
        output_folder = 'output_files'
        os.makedirs(output_folder, exist_ok=True)
        output_file = os.path.join(output_folder, f'participants_with_links_{int(time.time())}.xlsx')
        df.to_excel(output_file, index=False)

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'error': 'An unexpected error occurred while processing the file.'}), 500


# Main entry point for running the Flask app
if __name__ == '__main__':
    app.run(debug=True, port=5002)
