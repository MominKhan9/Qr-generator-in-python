import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog, messagebox
import qrcode
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import os
import threading
import subprocess



# Google Drive
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/drive.file']

last_modified_pdf = None


def upload_to_drive(filepath):
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('drive', 'v3', credentials=creds)

    file_metadata = {'name': os.path.basename(filepath)}
    media = MediaFileUpload(filepath, mimetype='application/pdf')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()

    # Make file public
    service.permissions().create(fileId=file['id'], body={'role': 'reader', 'type': 'anyone'}).execute()

    return f"https://drive.google.com/file/d/{file['id']}/view?usp=sharing"


def create_qr_pdf(pdf_path):
    # Generate a QR code from the PDF path (could be a URL or a hash of the PDF)
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L,
                       box_size=10, border=4)
    qr.add_data(pdf_path)  # You can add the file path or any URL
    qr.make(fit=True)

    # Create an image of the QR code
    img = qr.make_image(fill='black', back_color='white')

    # Specify a temporary directory on Windows (you can choose a directory of your choice)
    temp_dir = os.path.join(os.getcwd(), 'temp')  # Current working directory + 'temp' folder
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)  # Create the directory if it doesn't exist
    
    # Temporary file path to save the QR code image
    img_path = os.path.join(temp_dir, 'qr_code.png')

    # Save the QR code image to the temporary directory
    img.save(img_path)

    # Create a temporary PDF to place the QR code
    packet = BytesIO()
    c = canvas.Canvas(packet)
    c.drawImage(img_path, 510, 780, 60, 60)  # Position QR code (adjust coordinates as needed)
    c.save()

    # Move to the beginning of the BytesIO buffer
    packet.seek(0)

    return packet


def add_qr_to_pdf(pdf_path, qr_pdf_packet):
    pdf_reader = PdfReader(pdf_path)
    pdf_writer = PdfWriter()
    qr_pdf = PdfReader(qr_pdf_packet)

    for i in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[i]
        page.merge_page(qr_pdf.pages[0])
        pdf_writer.add_page(page)

     # Extract original filename without extension
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    # Create new filename with _modified suffix
    modified_pdf_path = f"{base_name}_modified.pdf"

    # modified_pdf_path = "modified_document.pdf"
    with open(modified_pdf_path, "wb") as f:
        pdf_writer.write(f)

    return modified_pdf_path



def print_pdf(filepath):
    try:
        print(f"Sending {filepath} to printer...")
        os.startfile(filepath, "printto")  # Windows-only print command
    except Exception as e:
        messagebox.showerror("Print Error", f"Failed to print file: {e}")


def open_pdf(filepath):
    try:
        subprocess.Popen(['start', '', filepath], shell=True)
    except Exception as e:
        messagebox.showerror("Open Error", f"Failed to open file: {e}")


def handle_pdf(pdf_path):
    print(f"Uploading to Google Drive: {pdf_path}")
    link = upload_to_drive(pdf_path)
    print(f"Public link: {link}")

    qr_pdf_packet = create_qr_pdf(link)
    modified_pdf = add_qr_to_pdf(pdf_path, qr_pdf_packet)

    global last_modified_pdf
    last_modified_pdf = modified_pdf  # Save it for printing

    print(f"Modified PDF saved as: {modified_pdf}")
    popup_success("Success", f"File '{modified_pdf}' has been uploaded and modified successfully!")
    print_button.config(state='normal')  # Enable print button
    animate_frame()



def on_drop(event):
    pdf_path = event.data.strip('{}')  # Clean Windows curly braces
    print(f"File dropped: {pdf_path}")
    handle_pdf(pdf_path)


def open_file_dialog():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        print(f"File selected: {pdf_path}")
        handle_pdf(pdf_path)



def popup_success(title, message):
    messagebox.showinfo(title, message)

def animate_frame():
    """Simple animation: change frame color briefly."""
    original_color = '#2e2e42'
    highlight_color = '#505072'

    def _animate():
        frame.configure(bg=highlight_color)
        label.configure(bg=highlight_color)
        insert_button.configure(bg=highlight_color)
        root.update()
        # After 300 milliseconds, revert to original color
        root.after(300, lambda: (
            frame.configure(bg=original_color),
            label.configure(bg=original_color),
            insert_button.configure(bg=original_color)
        ))

    threading.Thread(target=_animate).start()



# GUI Setup
root = TkinterDnD.Tk()
root.title("PDF QR Code Inserter")

root.geometry("600x400")
root.configure(bg='#1e1e2f')

frame = tk.Frame(root, bg='#2e2e42', bd=2, relief="solid")
frame.place(relx=0.5, rely=0.5, anchor="center", width=400, height=250)

label = tk.Label(
    frame,
    text="Drag and Drop PDF Here",
    fg="white",
    bg="#2e2e42",
    font=("Helvetica", 16, "bold")
)
label.pack(pady=40)

insert_button = tk.Button(
    frame,
    text="Browse Files",
    command=open_file_dialog,
    bg="#4e4e6a",
    fg="white",
    activebackground="#6e6e8a",
    activeforeground="white",
    font=("Helvetica", 12, "bold"),
    relief="flat",
    padx=10,
    pady=5
)
insert_button.pack(pady=20)

print_button = tk.Button(
    frame,
    text="Print PDF",
    command=lambda: print_pdf(last_modified_pdf),
    bg="#4e4e6a",
    fg="white",
    activebackground="#6e6e8a",
    activeforeground="white",
    font=("Helvetica", 12, "bold"),
    relief="flat",
    padx=10,
    pady=5,
    state='disabled'  # Initially disabled
)
print_button.pack(pady=10)


# DnD binding
frame.drop_target_register(DND_FILES)
frame.dnd_bind('<<Drop>>', on_drop)

root.mainloop()