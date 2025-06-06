import os
import time
import pandas as pd
import webbrowser
from tkinter import Tk, filedialog, Button, Label, Frame, StringVar, BooleanVar, Checkbutton
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PIL import Image

# Avery 5160 label constants
LABEL_WIDTH = 2.625 * inch
LABEL_HEIGHT = 1.0 * inch
PAGE_WIDTH, PAGE_HEIGHT = letter
COLUMNS = 3
ROWS = 10
H_MARGIN = 0.19 * inch
V_MARGIN = 0.5 * inch
H_GAP = 0.125 * inch

output_pdf = "avery_labels.pdf"

def generate_pdf(csv_path):
    try:
        df = pd.read_csv(csv_path)

        if "code" not in df.columns:
            status_var.set("❌ CSV must contain a 'code' column.")
            return

        total_labels = len(df)
        c = canvas.Canvas(output_pdf, pagesize=letter)

        for idx, row in df.iterrows():
            barcode_data = str(row["code"])
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}.png"

            barcode_obj = Code128(barcode_data, writer=ImageWriter())
            barcode_obj.save(barcode_base)

            time.sleep(0.1)
            img = Image.open(barcode_filename)
            img.save(barcode_filename)
            img.close()

            col = idx % COLUMNS
            row_pos = (idx // COLUMNS) % ROWS
            x = H_MARGIN + col * (LABEL_WIDTH + H_GAP)
            y = PAGE_HEIGHT - V_MARGIN - (row_pos + 1) * LABEL_HEIGHT

            if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                c.showPage()

            c.drawImage(barcode_filename, x + 4, y + 5, width=LABEL_WIDTH - 8, height=LABEL_HEIGHT - 10)
            os.remove(barcode_filename)

        c.save()
        status_var.set(f"✅ {total_labels} labels generated.")
        link_label.config(text="📂 Open PDF", fg="#2196f3")
        link_label.bind("<Button-1>", lambda e: webbrowser.open(output_pdf))

    except Exception as e:
        status_var.set(f"❌ Error: {str(e)}")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        status_var.set("Processing...")
        generate_pdf(file_path)

def toggle_theme():
    dark = theme_var.get()
    bg = "#2e2e2e" if dark else "#f4f4f4"
    fg = "#ffffff" if dark else "#000000"
    hint_fg = "#cccccc" if dark else "#666666"

    root.configure(bg=bg)

    # Apply background only to frame (no text color)
    frame.configure(bg=bg)

    # Apply full styling to widgets with text
    for widget in [header, subheader, status, link_label]:
        widget.configure(bg=bg, fg=fg)

    subheader.configure(fg=hint_fg)
    checkbox.configure(bg=bg, fg=fg)
    status.configure(fg=fg)


# === GUI Setup ===
root = Tk()
root.title("🧾 Avery 5160 Barcode Label Generator")
root.geometry("420x270")
root.configure(bg="#f4f4f4")

# Theme toggle
theme_var = BooleanVar()
theme_var.set(False)

status_var = StringVar()
status_var.set("Please upload a CSV with a 'code' column.")

header = Label(root, text="Barcode Label Generator", font=("Helvetica", 16, "bold"), bg="#f4f4f4")
header.pack(pady=(20, 5))

subheader = Label(root, text="CSV must include a 'code' column", font=("Helvetica", 10), bg="#f4f4f4", fg="#666")
subheader.pack()

frame = Frame(root, bg="#f4f4f4")
frame.pack(pady=(10, 5))

btn = Button(frame, text="📂 Choose CSV", font=("Helvetica", 11), command=select_file, bg="#4caf50", fg="white", padx=12, pady=6)
btn.grid(row=0, column=0, padx=5)

checkbox = Checkbutton(frame, text="🌙 Dark Mode", variable=theme_var, command=toggle_theme, bg="#f4f4f4")
checkbox.grid(row=0, column=1, padx=5)

link_label = Label(root, text="", font=("Helvetica", 10, "underline"), bg="#f4f4f4", cursor="hand2")
link_label.pack(pady=(5, 0))

status = Label(root, textvariable=status_var, wraplength=360, bg="#f4f4f4", font=("Helvetica", 9), fg="#333")
status.pack(pady=15)

root.mainloop()




