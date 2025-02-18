from PIL.ImageFont import ImageFont
from flask import Flask, request, render_template, send_file
import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import zipfile
from io import BytesIO

app = Flask(__name__, template_folder=r"S:\MMT Work\Certificate generator\templates")
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Certificate settings
certificate_template = r"S:\MMT Work\Certificate generator\templates\Cert_template.png"
font_path = r"S:\MMT Work\Certificate generator\templates\EBGaramond-VariableFont_wght.ttf"
font_size = 100
name_position = (1400, 1160)

@app.route('/')
def index():
    return render_template('testform.html')

@app.route('/generate', methods=['GET'])
def generate_certificates():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    if not file.filename.endswith('.xlsx'):
        return "Invalid file type. Please re-upload the excel file and try again.", 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    try:
        # Verify template file
        if not os.path.exists(certificate_template):
            return "Error: Certificate template file not found!", 500

        # Verify font file
        if not os.path.exists(font_path):
            return "Error: Font file not found!", 500

        # Test loading font
        try:
            font = ImageFont.truetype(font_path, font_size)
        except Exception as e:
            return f"Error loading font: {e}", 500

        # Process Excel file
        df = pd.read_excel(file_path)

        if 'Full name (as per NRIC)' not in df.columns:
            return "Error: 'Full name (as per NRIC)' column not found in the Excel file.", 400

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            for _, row in df.iterrows():
                participant_name = row.get('Full name (as per NRIC)', '').strip()
                if not participant_name:
                    continue  # Skip rows with empty names

                # Generate certificate
                img = Image.open(certificate_template)
                draw = ImageDraw.Draw(img)
                draw.text(name_position, participant_name, font=font, fill="black")

                # Save certificate
                certificate_filename = f"certificate_{participant_name}.png"
                certificate_path = os.path.join(app.config['UPLOAD_FOLDER'], certificate_filename)
                img.save(certificate_path)

                # Add to ZIP archive
                zipf.write(certificate_path, certificate_filename)

        # Send ZIP file
        zip_buffer.seek(0)
        return send_file(zip_buffer, as_attachment=True, download_name="Certificates.zip", mimetype="application/zip")

    except Exception as e:
        return str(e), 500


if __name__ == '__main__':
    app.run(debug=True, port=5001)
