import shutil

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import zipfile
import tkinter as tk
from tkinter import filedialog
import tempfile
from tkcalendar import DateEntry


class CertGenerator:

    def __init__(self):
        self.temp_dir = tempfile.mkdtemp(prefix="Certificates_")
        print(f"Temporary directory created at: {self.temp_dir}")

        self.root = tk.Tk()
        self.root.withdraw()

    def excel_path(self, excel_path):
        self.excel_path = excel_path

    def load_participants(self):
        if not self.excel_path:
            raise ValueError("Excel file path is not set.")
        self.df = pd.read_excel(self.excel_path)
        self.df.columns = self.df.columns.str.strip()

        print("Excel data loaded successfully.")
        print("DataFrame Columns:", self.df.columns)  # Print the column names to verify
        print(self.df.head())  # Print the first few rows to check data

    def paths(self):
        self.certificate_template = r"S:\MMT Work\Cert template.png"
        self.font_path_regular = r"C:\Users\Shanjif.C\Downloads\EB_Garamond\EBGaramond-VariableFont_wght.ttf"
        self.font_path_bolditalic = r"C:\Users\Shanjif.C\Downloads\EB-Garamond-Italic\ebgaramond\EBGaramond-BoldItalic.ttf"
        self.font_path_italic = r"C:\Users\Shanjif.C\Downloads\EB-Garamond-Italic\ebgaramond\EBGaramond-Italic.ttf"

    def fonts(self):
        self.font_size_name = 100
        self.font_size_event = 50
        self.name_position = (1400, 1160)
        self.event_position = (1370, 900)

    def event_details(self, participant_name, event_name, event_date, event_venue):
        self.participant_name = participant_name
        self.event_name = event_name
        self.event_date = event_date
        self.event_venue = event_venue
        self.event_intro = 'For attending the "'
        self.event_outro_part1 = f'" held in {self.event_venue}'
        self.event_outro_part2 = f'on {self.event_date}.'

    def certgenerator(self, participant_name):
        print(f"Entered certgenerator with participant_name: {participant_name}")
        if not hasattr(self, 'df'):
            raise ValueError("Participant data is not loaded. Please load the Excel file first.")

        font_regular = ImageFont.truetype(self.font_path_regular, self.font_size_name)
        font_bolditalic = ImageFont.truetype(self.font_path_bolditalic, self.font_size_event)
        font_italic = ImageFont.truetype(self.font_path_italic, self.font_size_event)

        img = Image.open(self.certificate_template)
        draw = ImageDraw.Draw(img)

        draw.text(self.name_position, participant_name, font=font_regular, fill="black")
        name_bbox = draw.textbbox((0, 0), participant_name, font=font_regular)
        name_width = name_bbox[2] - name_bbox[0]

        draw.text(self.event_position, self.event_intro, font=font_italic, fill="gray")
        intro_bbox = draw.textbbox((0, 0), self.event_intro, font=font_italic)
        intro_width = intro_bbox[2] - intro_bbox[0]

        event_name_position = (self.event_position[0] + intro_width, self.name_position[1])
        draw.text(event_name_position, self.event_name, font=font_bolditalic, fill="gray")
        event_name_bbox = draw.textbbox((0, 0), self.event_name, font=font_bolditalic)
        event_name_width = event_name_bbox[2] - event_name_bbox[0]

        event_outro_position_part1 = (event_name_position[0] + event_name_width, self.name_position[1])
        draw.text(event_outro_position_part1, self.event_outro_part1, font=font_italic, fill="gray")
        event_outro_bbox_part1 = draw.textbbox((0, 0), self.event_outro_part1, font=font_italic)
        event_outro_width_part1 = event_outro_bbox_part1[2] - event_outro_bbox_part1[0]

        event_outro_position_part2 = (
        self.event_position[0], self.event_position[1] + 60)  # Adjust '60' for line spacing as needed
        draw.text(event_outro_position_part2, self.event_outro_part2, font=font_italic, fill="gray")

        certificate_filename = f"certificate_{participant_name}.png"
        certificate_path = os.path.join(self.temp_dir, certificate_filename)
        img.save(certificate_path)
        img.close()
        print(f"Certificate generated for {participant_name} at {certificate_path}")


    def filesaving(self):
        if not hasattr(self, 'df'):
            raise ValueError("Participant data is not loaded. Please load the Excel file first.")

        zip_filename = "Certificates_PNGs.zip"
        zip_path = os.path.join(self.temp_dir, zip_filename)

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root_dir, _, files in os.walk(self.temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    zipf.write(file_path, os.path.basename(file_path))

            print("Contents of ZIP file:")
            for item in zipf.namelist():
                print(item)

        save_as_path = filedialog.asksaveasfilename(
            initialfile=zip_filename,
            title="Save Certificates Zip File As",
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )

        if save_as_path:
            # Move the temporary .zip file to the chosen location
            shutil.move(zip_path, save_as_path)
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
            self.df.to_excel(output_excel_path, index=False)
            print("Excel file with certificate paths saved to:", output_excel_path)
        else:
            print("Excel file save was canceled.")

        print("Certificate generation process completed.")

        # Clean up temporary directory after use
        for file in os.listdir(self.temp_dir):
            file_path = os.path.join(self.temp_dir, file)
            os.remove(file_path)
        os.rmdir(self.temp_dir)

    def main(self):
        def main(self):
            self.paths()
            self.fonts()
            if not hasattr(self, 'df'):
                raise ValueError("Participant data is not loaded. Please load the Excel file first.")

            for index, row in self.df.iterrows():
                participant_name = row.get('Full name (as per NRIC)', '').strip()

                if participant_name:
                    print(f"Calling certgenerator for: {participant_name}")
                    self.certgenerator(participant_name)
                else:
                    print(f"Missing name at row {index}, skipping.")

        self.filesaving()

if __name__ == "__main__":
    generator = CertGenerator()
    generator.main()


