import os
import time
import pandas as pd
from tkinter import Tk, filedialog, Button, Label
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PIL import Image

# Avery 5160 layout constants
LABEL_WIDTH = 2.625 * inch
LABEL_HEIGHT = 1.0 * inch
PAGE_WIDTH, PAGE_HEIGHT = letter
COLUMNS = 3
ROWS = 10
H_MARGIN = 0.19 * inch
V_MARGIN = 0.5 * inch
H_GAP = 0.125 * inch
V_GAP = 0.0 * inch  # No vertical gap

def generate_pdf(csv_path):
    try:
        df = pd.read_csv(csv_path)

        if "code" not in df.columns:
            label.config(text="❌ CSV must contain a 'code' column.")
            return

        output_pdf = "avery_labels.pdf"
        c = canvas.Canvas(output_pdf, pagesize=letter)

        for idx, row in df.iterrows():
            barcode_data = str(row["code"])
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}.png"

            # Generate barcode and save to file (no extension in save call)
            barcode_obj = Code128(barcode_data, writer=ImageWriter())
            barcode_obj.save(barcode_base)

            # Ensure file is fully saved
            time.sleep(0.1)
            img = Image.open(barcode_filename)
            img.save(barcode_filename)
            img.close()

            # Calculate label position
            col = idx % COLUMNS
            row_pos = (idx // COLUMNS) % ROWS
            x = H_MARGIN + col * (LABEL_WIDTH + H_GAP)
            y = PAGE_HEIGHT - V_MARGIN - (row_pos + 1) * LABEL_HEIGHT

            if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                c.showPage()

            # Draw barcode image and text
            c.drawImage(barcode_filename, x + 4, y + 10, width=LABEL_WIDTH - 8, height=LABEL_HEIGHT - 25)
            c.setFont("Helvetica", 8)
            c.drawCentredString(x + LABEL_WIDTH / 2, y + 2, barcode_data)

            # Clean up barcode image file
            os.remove(barcode_filename)

        c.save()
        label.config(text="✅ PDF created: avery_labels.pdf")

    except Exception as e:
        label.config(text=f"❌ Error: {str(e)}")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        label.config(text="Processing...")
        generate_pdf(file_path)

# === GUI Setup ===
root = Tk()
root.title("Avery 5160 Barcode Label Generator")
root.geometry("320x160")

label = Label(root, text="Upload a CSV with a 'code' column", wraplength=300)
label.pack(pady=10)

btn = Button(root, text="Choose CSV and Generate PDF", command=select_file)
btn.pack(pady=20)

root.mainloop()





