import tkinter as tk
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
from cert_gen import CertGenerator

# Initialize the certificate generator instance
cert_gen = CertGenerator()


def select_file():
    # Open a file dialog to select the Excel file
    excel_path = filedialog.askopenfilename(
        title="Select Participant Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if excel_path:
        # Set the selected file path in the cert_gen instance and load data
        cert_gen.excel_path = excel_path
        file_label.config(text=f"Selected File: {excel_path}")
        cert_gen.load_participants()  # Load participant data from Excel
        messagebox.showinfo("File Loaded", "Excel data loaded successfully.")
    else:
        file_label.config(text="No file selected")
        messagebox.showwarning("No File", "Please select a valid Excel file.")


def generate_cert():
    try:
        # Retrieve input values from the UI
        participant_name = name_entry.get()
        event_venue = venue_entry.get()
        event_name = event_name_entry.get()
        event_date = event_date_picker.get_date().strftime("%B %d, %Y")

        # Configure cert_gen instance with event details
        cert_gen.paths()
        cert_gen.fonts()
        cert_gen.event_details(participant_name, event_name, event_date, event_venue)

        # Generate certificates and save them
        cert_gen.certgenerator(participant_name)
        certificate_path = cert_gen.filesaving()

        messagebox.showinfo("Success", f"Certificates saved to: {certificate_path}")
    except ValueError as e:
        messagebox.showerror("Error", str(e))


# Set up the Tkinter UI
root = tk.Tk()
root.title("Certificate Generator")

# Excel file selection section
tk.Button(root, text="Upload Excel File", command=select_file).grid(row=0, column=0, sticky="w")
file_label = tk.Label(root, text="No file selected")
file_label.grid(row=0, column=1, sticky="w")

# Participant name entry
tk.Label(root, text="Participant Name:").grid(row=1, column=0, sticky="w")
name_entry = tk.Entry(root)
name_entry.grid(row=1, column=1, sticky="w")

# Venue entry
tk.Label(root, text="Venue:").grid(row=2, column=0, sticky="w")
venue_entry = tk.Entry(root)
venue_entry.grid(row=2, column=1, sticky="w")

# Event name entry
tk.Label(root, text="Event Name:").grid(row=3, column=0, sticky="w")
event_name_entry = tk.Entry(root)
event_name_entry.grid(row=3, column=1, sticky="w")

# Event date entry with date picker
tk.Label(root, text="Event Date:").grid(row=4, column=0, sticky="w")
event_date_picker = DateEntry(root, width=12, background="darkblue", foreground="white", borderwidth=2)
event_date_picker.grid(row=4, column=1, sticky="w")

# Button to generate the certificate
generate_button = tk.Button(root, text="Generate Certificate", command=generate_cert)
generate_button.grid(row=5, columnspan=2, pady=10)

# Run the UI main loop
root.mainloop()
