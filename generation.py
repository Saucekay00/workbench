import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import zipfile
import tkinter as tk
from tkinter import *
from tkcalendar import DateEntry

def generate_certificate(participant_name, event_name, event_date, event_venue, org_name, department):
    excel_path = r"S:\\MMT Work\\ALI_Online_Attendance_2024-10-21_23_44_00.xlsx"
    certificate_template = r"S:\MMT Work\Cert template.png"
    font_path_regular = r"C:\Users\Shanjif.C\Downloads\EB_Garamond\EBGaramond-VariableFont_wght.ttf"
    font_path_bolditalic = r"C:\Users\Shanjif.C\Downloads\EB-Garamond-Italic\ebgaramond\EBGaramond-BoldItalic.ttf"
    font_path_italic = r"C:\Users\Shanjif.C\Downloads\EB-Garamond-Italic\ebgaramond\EBGaramond-Italic.ttf"

    # Font sizes and positions
    font_size_name = 100
    font_size_event = 50
    name_position = (1400, 1160)
    event_position = (1370, 900)

    # Event details (adjust as needed)
    event_name = "Design Thinking"
    event_date = "September 18th & 19th, 2024"
    event_venue = "Hotel Eastin, Penang"
    event_intro = 'For attending the "'
    event_outro_part1 = f'" held in {event_venue}'
    event_outro_part2 = f'on {event_date}.'

    # Load participant data from Excel
    df = pd.read_excel(excel_path)
    df = df.head(1)  # Limit rows for testing, remove this for full dataset

    # Create a directory to store generated certificates
    temp_dir = "S:\\MMT Work\\TemporaryCertificates"
    os.makedirs(temp_dir, exist_ok=True)

    # Load fonts
    font_regular = ImageFont.truetype(font_path_regular, font_size_name)
    font_bolditalic = ImageFont.truetype(font_path_bolditalic, font_size_event)
    font_italic = ImageFont.truetype(font_path_italic, font_size_event)

    # Generate certificates and save them as PNGs
    for index, row in df.iterrows():
        participant_name = row['Full name (as per NRIC)']
        sanitized_name = "".join(char for char in participant_name if char.isalnum() or char in " _-")
        certificate_filename = f"certificate_{sanitized_name}.png"
        certificate_path = os.path.join(temp_dir, certificate_filename)

        print(f"Generating certificate for {participant_name} ({index + 1}/{len(df)})")

        # Open the certificate template and set up drawing context
        img = Image.open(certificate_template)
        draw = ImageDraw.Draw(img)

        # Draw the participant's name in bold
        draw.text(name_position, participant_name, font=font_regular, fill="black")

        # Draw event details with mixed styles
        # Draw the intro part in italic
        draw.text(event_position, event_intro, font=font_italic, fill="gray")
        intro_bbox = draw.textbbox((0, 0), event_intro, font=font_italic)
        intro_width = intro_bbox[2] - intro_bbox[0]

        # Draw the event name in bold
        event_name_position = (event_position[0] + intro_width, event_position[1])
        draw.text(event_name_position, event_name, font=font_bolditalic, fill="gray")
        event_name_bbox = draw.textbbox((0, 0), event_name, font=font_bolditalic)
        event_name_width = event_name_bbox[2] - event_name_bbox[0]

        # Draw the outro part in italic
        event_outro_position_part1 = (event_name_position[0] + event_name_width, event_position[1])
        draw.text(event_outro_position_part1, event_outro_part1, font=font_italic, fill="gray")
        event_outro_bbox_part1 = draw.textbbox((0, 0), event_outro_part1, font=font_italic)
        event_outro_width_part1 = event_outro_bbox_part1[2] - event_outro_bbox_part1[0]

        event_outro_position_part2 = (
        event_position[0], event_position[1] + 60)  # Adjust '60' for line spacing as needed
        draw.text(event_outro_position_part2, event_outro_part2, font=font_italic, fill="gray")

        img.save(certificate_path)
        img.close()  # Close image to free up memory

        # Store the PNG path in the DataFrame for reference
        df.loc[index, 'Certificate File Path'] = certificate_path

    # Create a .zip file of all PNGs in the temporary directory
    zip_filename = "Certificates_PNGs.zip"
    zip_path = os.path.join(temp_dir, zip_filename)

    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root_dir, _, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                zipf.write(file_path, os.path.basename(file_path))  # Add file to zip with just the filename

        # List contents of the ZIP file for verification
        print("Contents of ZIP file:")
        for item in zipf.namelist():
            print(item)

    # Prompt the user to save the .zip file
    save_as_path = filedialog.asksaveasfilename(
        initialfile=zip_filename,
        title="Save Certificates Zip File As",
        defaultextension=".zip",
        filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
    )

    if save_as_path:
        # Move the temporary .zip file to the chosen location
        os.replace(zip_path, save_as_path)
        print("Certificates saved to:", save_as_path)
    else:
        print("Save canceled. Certificates zip file not saved.")

    # Optionally save the updated Excel file with file paths for reference
    output_excel_path = filedialog.asksaveasfilename(
        initialfile="Updated_Attendance_With_Certificates.xlsx",
        title="Save Updated Excel File As",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    if output_excel_path:
        df.to_excel(output_excel_path, index=False)
        print("Excel file with certificate paths saved to:", output_excel_path)
    else:
        print("Excel file save was canceled.")

    print("Certificate generation process completed.")

    # Clean up temporary directory after use
    for file in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, file)
        os.remove(file_path)
    os.rmdir(temp_dir)
