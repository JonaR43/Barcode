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
    OptionMenu,
)
from tkinter import ttk
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

# Label definitions
label_types = {
    "Avery 5160": (2.625 * inch, 1.0 * inch),
    "Avery 5163": (4.0 * inch, 2.0 * inch)
}
PAGE_WIDTH, PAGE_HEIGHT = letter
COLUMNS = 3
ROWS = 10
H_MARGIN = 0.15 * inch  # Reduced from 0.19
V_MARGIN = 0.3 * inch   # Keep the same
H_GAP = 0.20 * inch     # Reduced from 0.125
#V_GAP = 0.0 * inch      # Adjust the row gap
output_pdf = "avery_labels.pdf"

preview_image = None

def generate_pdf(csv_path, preview_only=False):
    global preview_imagecs
    try:
        df = pd.read_csv(csv_path)
        if "code" not in df.columns:
            status_var.set("❌ CSV must contain a 'code' column.")
            return

        label_w, label_h = label_types[label_type.get()]
        show_text = show_text_var.get()

        if preview_only:
            code = str(df.iloc[0]["code"])
            filename = "preview_barcode"
            Code128(code, writer=ImageWriter()).save(filename, options={
                "font_size": int(barcode_font_size_var.get()),
                "font_path": "Calibri.ttf",
                "module_height": 20
            })
            img = Image.open(f"{filename}.png")

            # --- Scale barcode to fit within the preview canvas, maintaining aspect ratio ---
            canvas_w, canvas_h = 300, 100
            img_w, img_h = img.size
            scale = min((canvas_w - 10) / img_w, (canvas_h - 10) / img_h)  # 5px margin
            new_w = int(img_w * scale)
            new_h = int(img_h * scale)
            img = img.resize((new_w, new_h), Image.LANCZOS)

            # Center the image in the canvas
            x_offset = (canvas_w - new_w) // 2
            y_offset = (canvas_h - new_h) // 2

            preview_image = ImageTk.PhotoImage(img)
            label_canvas.delete("all")
            label_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill="white", outline="gray")
            label_canvas.create_image(x_offset, y_offset, anchor="nw", image=preview_image)

            img.close()
            os.remove(f"{filename}.png")
            return

        total_labels = len(df)
        c = pdf_canvas.Canvas(output_pdf, pagesize=letter)


        # Calculate vertical spacing
        usable_height = PAGE_HEIGHT - (2 * V_MARGIN)
        total_label_height = ROWS * label_h
        v_gap = (usable_height - total_label_height) / (ROWS - 1)

        for idx, row in df.iterrows():
            col = idx % COLUMNS
            row_pos = (idx // COLUMNS) % ROWS
            
            # Generate the barcode image first
            barcode_data = str(row["code"])
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}.png"
            
            # Generate barcode image with specified font size
            Code128(barcode_data, writer=ImageWriter()).save(barcode_base, options={
                "font_size": int(barcode_font_size_var.get()),
                "font_path": "Calibri.ttf",
                "module_height": 20
            })
            
            # Calculate label position
            label_w, label_h = label_types[label_type.get()]
            x = H_MARGIN + col * (label_w + H_GAP)
            y = PAGE_HEIGHT - V_MARGIN - (row_pos * (label_h + v_gap)) - label_h

            if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                c.showPage()

            # Calculate centered position with more padding
            padding = 0.08 * inch  # Add padding around each barcode
            barcode_width = label_w - (2 * padding)  # Reduce width by padding on both sides
            barcode_height = label_h - (2 * padding)  # Reduce height by padding on both sides
            x_centered = x + padding  # Add padding to left
            y_centered = y + padding  # Add padding to bottom

            # Draw the barcode with centered positioning and padding
            c.drawImage(
                barcode_filename,
                x_centered,
                y_centered,
                width=barcode_width,
                height=barcode_height
            )
            
            # Draw label outline (grid)
            c.saveState()
            c.setStrokeColorRGB(0.8, 0.8, 0.8)  # Light gray color
            c.setLineWidth(0.25)  # Thin line
            c.rect(x, y, label_w, label_h)  # Draw rectangle around label area
            c.restoreState()

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
        abs_path = os.path.abspath(file_path)  # 👈 Ensure absolute path
        status_var.set("Processing...")
        generate_pdf(abs_path)
        generate_pdf(abs_path, preview_only=True)


def toggle_theme():
    dark = theme_var.get()
    bg = "#2e2e2e" if dark else "#f4f4f4"
    fg = "#ffffff" if dark else "#000000"
    hint_fg = "#cccccc" if dark else "#666666"

    root.configure(bg=bg)
    for widget in [status, link_label, drop_label, preview_label, header]:
        widget.configure(bg=bg, fg=fg)
    subheader.configure(bg=bg, fg=hint_fg)
    settings_tab.configure(bg=bg)
    generator_tab.configure(bg=bg)
    label_canvas.configure(bg="white")

def handle_drop(event):
    file_path = event.data.strip("{}")
    if file_path.lower().endswith(".csv"):
        abs_path = os.path.abspath(file_path)  # 👈 Ensure absolute path
        status_var.set("Processing dropped file...")
        generate_pdf(abs_path)
        generate_pdf(abs_path, preview_only=True)
    else:
        status_var.set("❌ Only CSV files are supported.")


def on_drag_enter(event):
    drop_label.config(bg="#d0ffd0")

def on_drag_leave(event):
    drop_label.config(bg="#e0e0e0")

# === GUI Setup ===
root = BaseTk()
root.title("🧾 Avery Barcode Label Generator")
root.geometry("520x640")
root.configure(bg="#f4f4f4")

theme_var = BooleanVar()
theme_var.set(False)
show_text_var = BooleanVar(value=True)
barcode_font_size_var = StringVar(value="14")  # default barcode font size
label_type = StringVar(value="Avery 5160")
status_var = StringVar()
status_var.set("Upload or drag a CSV with a 'code' column.")

# === Notebook ===
notebook = ttk.Notebook(root)
notebook.pack(expand=1, fill="both", pady=5)

generator_tab = Frame(notebook, bg="#f4f4f4")
settings_tab = Frame(notebook, bg="#f4f4f4")
notebook.add(generator_tab, text="🧾 Generator")
notebook.add(settings_tab, text="⚙️ Settings")

# === Generator Tab ===
header = Label(generator_tab, text="Barcode Label Generator", font=("Helvetica", 16, "bold"), bg="#f4f4f4")
header.pack(pady=(15, 3))

subheader = Label(generator_tab, text="CSV must include a 'code' column", font=("Helvetica", 10), bg="#f4f4f4", fg="#666")
subheader.pack()

frame = Frame(generator_tab, bg="#f4f4f4")
frame.pack(pady=(10, 5))

btn = Button(frame, text="📂 Choose CSV", font=("Helvetica", 11), command=select_file, bg="#4caf50", fg="white", padx=12, pady=6)
btn.grid(row=0, column=0, padx=5)

link_label = Label(generator_tab, text="", font=("Helvetica", 10, "underline"), bg="#f4f4f4", cursor="hand2")
link_label.pack(pady=(5, 0))

status = Label(generator_tab, textvariable=status_var, wraplength=400, bg="#f4f4f4", font=("Helvetica", 9), fg="#333")
status.pack(pady=10)

drop_label = Label(generator_tab, text="⬇️ Drop CSV file here", font=("Helvetica", 10, "italic"), bg="#e0e0e0", fg="#444", relief="groove", width=40, height=3, bd=2)
drop_label.pack(pady=(0, 10))

if dragdrop_enabled:
    drop_label.drop_target_register(DND_FILES)
    drop_label.dnd_bind("<<Drop>>", handle_drop)
    drop_label.dnd_bind("<<DragEnter>>", on_drag_enter)
    drop_label.dnd_bind("<<DragLeave>>", on_drag_leave)

preview_label = Label(generator_tab, text="🔍 Label Preview", font=("Helvetica", 10, "bold"), bg="#f4f4f4")
preview_label.pack()

label_canvas = Canvas(generator_tab, width=300, height=100, bg="white", bd=1, relief="sunken")
label_canvas.pack(pady=(5, 20))

# === Settings Tab ===
Label(settings_tab, text="⚙️ Settings", font=("Helvetica", 14, "bold"), bg="#f4f4f4").pack(pady=(20, 5))
Checkbutton(settings_tab, text="Show Text Below Barcode", variable=show_text_var, bg="#f4f4f4").pack(pady=5)
Label(settings_tab, text="Label Type:", bg="#f4f4f4").pack()
OptionMenu(settings_tab, label_type, *label_types.keys()).pack(pady=5)
Checkbutton(settings_tab, text="🌙 Dark Mode", variable=theme_var, command=toggle_theme, bg="#f4f4f4").pack(pady=(10, 20))
Label(settings_tab, text="Barcode Font Size (under barcode)", bg="#f4f4f4").pack()
OptionMenu(settings_tab, barcode_font_size_var, *[str(i) for i in range(6, 30)]).pack(pady=(0, 10))


root.mainloop()



