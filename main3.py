# Re-posting the full working script with simulated full-label preview (Option 2) implemented.

import os
import time
import pandas as pd
import webbrowser
from tkinter import (
    Tk,
    filedialog,
    Button,
    Label,
    Frame,
    StringVar,
    BooleanVar,
    Checkbutton,
    Canvas,
)
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PIL import Image, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
    dragdrop_enabled = True
except ImportError:
    BaseTk = Tk
    dragdrop_enabled = False

# Avery 5160 layout
LABEL_WIDTH = 2.625 * inch
LABEL_HEIGHT = 1.0 * inch
PAGE_WIDTH, PAGE_HEIGHT = letter
COLUMNS = 3
ROWS = 10
H_MARGIN = 0.19 * inch
V_MARGIN = 0.5 * inch
H_GAP = 0.125 * inch
output_pdf = "avery_labels.pdf"

preview_image = None  # Global reference to prevent garbage collection


def generate_pdf(csv_path, preview_only=False):
    global preview_image

    try:
        df = pd.read_csv(csv_path)
        if "code" not in df.columns:
            status_var.set("❌ CSV must contain a 'code' column.")
            return

        if preview_only:
            code = str(df.iloc[0]["code"])
            filename = "preview_barcode"
            Code128(code, writer=ImageWriter()).save(filename)
            img = Image.open(f"{filename}.png")
            img = img.resize((260, 60))  # Resize for label preview
            preview_image = ImageTk.PhotoImage(img)
            label_canvas.delete("all")
            label_canvas.create_rectangle(0, 0, 300, 100, fill="white", outline="gray")
            label_canvas.create_image(20, 20, anchor="nw", image=preview_image)
            label_canvas.create_text(150, 90, text=code, font=("Helvetica", 8))
            img.close()
            os.remove(f"{filename}.png")
            return

        total_labels = len(df)
        c = pdf_canvas.Canvas(output_pdf, pagesize=letter)

        for idx, row in df.iterrows():
            barcode_data = str(row["code"])
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}.png"

            Code128(barcode_data, writer=ImageWriter()).save(barcode_base)
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

            c.drawImage(
                barcode_filename,
                x + 4,
                y + 5,
                width=LABEL_WIDTH - 8,
                height=LABEL_HEIGHT - 10,
            )
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
        generate_pdf(file_path, preview_only=True)


def toggle_theme():
    dark = theme_var.get()
    bg = "#2e2e2e" if dark else "#f4f4f4"
    fg = "#ffffff" if dark else "#000000"
    hint_fg = "#cccccc" if dark else "#666666"

    root.configure(bg=bg)
    frame.configure(bg=bg)
    preview_frame.configure(bg=bg)
    label_canvas.configure(bg="white")

    for widget in [header, subheader, status, link_label, drop_label]:
        widget.configure(bg=bg, fg=fg)

    subheader.configure(fg=hint_fg)
    checkbox.configure(bg=bg, fg=fg)
    drop_label.configure(bg="#444" if dark else "#e0e0e0")


def handle_drop(event):
    file_path = event.data.strip("{}")
    if file_path.lower().endswith(".csv"):
        status_var.set("Processing dropped file...")
        generate_pdf(file_path)
        generate_pdf(file_path, preview_only=True)
    else:
        status_var.set("❌ Only CSV files are supported.")


def on_drag_enter(event):
    drop_label.config(bg="#d0ffd0")


def on_drag_leave(event):
    drop_label.config(bg="#e0e0e0")


# === GUI Setup ===
root = BaseTk()
root.title("🧾 Avery 5160 Barcode Label Generator")
root.geometry("480x500")
root.configure(bg="#f4f4f4")

theme_var = BooleanVar()
theme_var.set(False)
status_var = StringVar()
status_var.set("Upload or drag a CSV with a 'code' column.")

header = Label(
    root, text="Barcode Label Generator", font=("Helvetica", 16, "bold"), bg="#f4f4f4"
)
header.pack(pady=(20, 5))

subheader = Label(
    root,
    text="CSV must include a 'code' column",
    font=("Helvetica", 10),
    bg="#f4f4f4",
    fg="#666",
)
subheader.pack()

frame = Frame(root, bg="#f4f4f4")
frame.pack(pady=(10, 5))

btn = Button(
    frame,
    text="📂 Choose CSV",
    font=("Helvetica", 11),
    command=select_file,
    bg="#4caf50",
    fg="white",
    padx=12,
    pady=6,
)
btn.grid(row=0, column=0, padx=5)

checkbox = Checkbutton(
    frame,
    text="🌙 Dark Mode",
    variable=theme_var,
    command=toggle_theme,
    bg="#f4f4f4",
)
checkbox.grid(row=0, column=1, padx=5)

link_label = Label(
    root,
    text="",
    font=("Helvetica", 10, "underline"),
    bg="#f4f4f4",
    cursor="hand2",
)
link_label.pack(pady=(5, 0))

status = Label(
    root,
    textvariable=status_var,
    wraplength=400,
    bg="#f4f4f4",
    font=("Helvetica", 9),
    fg="#333",
)
status.pack(pady=10)

# Drop zone
drop_label = Label(
    root,
    text="⬇️ Drop CSV file here",
    font=("Helvetica", 10, "italic"),
    bg="#e0e0e0",
    fg="#444",
    relief="groove",
    width=40,
    height=3,
    bd=2,
)
drop_label.pack(pady=(0, 10))

if dragdrop_enabled:
    drop_label.drop_target_register(DND_FILES)
    drop_label.dnd_bind("<<Drop>>", handle_drop)
    drop_label.dnd_bind("<<DragEnter>>", on_drag_enter)
    drop_label.dnd_bind("<<DragLeave>>", on_drag_leave)

# Preview panel
preview_frame = Frame(root, bg="#f4f4f4")
preview_frame.pack()

preview_label = Label(
    preview_frame, text="🔍 Label Preview", font=("Helvetica", 10, "bold"), bg="#f4f4f4"
)
preview_label.pack()

label_canvas = Canvas(
    preview_frame, width=300, height=100, bg="white", bd=1, relief="sunken"
)
label_canvas.pack(pady=(5, 20))

root.mainloop()



