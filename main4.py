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
    Scale,
    HORIZONTAL,
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
H_MARGIN = 0.19 * inch
V_MARGIN = 0.5 * inch
H_GAP = 0.125 * inch
output_pdf = "avery_labels.pdf"

preview_image = None
current_csv_path = None
barcode_width = 260
barcode_height = 60

# === GUI Setup ===
root = BaseTk()
root.title("🧾 Avery Barcode Label Generator")
root.geometry("520x760")
root.configure(bg="#f4f4f4")

text_font_size_var = StringVar(value="8")
barcode_font_size_var = StringVar(value="14")  # default font size under barcode
theme_var = BooleanVar()
theme_var.set(False)
show_text_var = BooleanVar(value=True)
label_type = StringVar(value="Avery 5160")
status_var = StringVar()
status_var.set("Upload or drag a CSV with a 'code' column.")

def generate_pdf(csv_path, preview_only=False):
    global preview_image
    try:
        df = pd.read_csv(csv_path)
        if "code" not in df.columns:
            status_var.set("❌ CSV must contain a 'code' column.")
            return

        label_w, label_h = label_types[label_type.get()]
        show_text = show_text_var.get()
        font_size = int(text_font_size_var.get())

        if preview_only:
            code = str(df.iloc[0]["code"])
            filename = "preview_barcode"
            Code128(code, writer=ImageWriter()).save(filename, options={"font_size": int(barcode_font_size_var.get())})
            img = Image.open(f"{filename}.png")
            img = img.resize((barcode_width, barcode_height))
            preview_image = ImageTk.PhotoImage(img)
            label_canvas.delete("all")
            label_canvas.create_rectangle(0, 0, 300, 100, fill="white", outline="gray")
            label_canvas.create_image(20, 20, anchor="nw", image=preview_image)
            if show_text:
                label_canvas.create_text(150, 90, text=code, font=("Helvetica", font_size))
            img.close()
            os.remove(f"{filename}.png")
            return

        total_labels = len(df)
        c = pdf_canvas.Canvas(output_pdf, pagesize=letter)

        for idx, row in df.iterrows():
            barcode_data = str(row["code"])
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}.png"

            Code128(barcode_data, writer=ImageWriter()).save(barcode_base, options={"font_size": int(barcode_font_size_var.get())})
            time.sleep(0.1)
            img = Image.open(barcode_filename)
            img.save(barcode_filename)
            img.close()

            col = idx % COLUMNS
            row_pos = (idx // COLUMNS) % ROWS
            x = H_MARGIN + col * (label_w + H_GAP)
            y = PAGE_HEIGHT - V_MARGIN - (row_pos + 1) * label_h

            if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                c.showPage()

            c.drawImage(barcode_filename, x + 4, y + 5, width=label_w - 8, height=label_h - 10)
            if show_text:
                c.setFont("Helvetica", font_size)
                c.drawCentredString(x + label_w / 2, y + 2, barcode_data)

            os.remove(barcode_filename)

        c.save()
        status_var.set(f"✅ {total_labels} labels generated.")
        link_label.config(text="📂 Open PDF", fg="#2196f3")
        link_label.bind("<Button-1>", lambda e: webbrowser.open(output_pdf))

    except Exception as e:
        status_var.set(f"❌ Error: {str(e)}")

def select_file():
    global current_csv_path
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        current_csv_path = file_path
        status_var.set("Processing...")
        generate_pdf(file_path)
        generate_pdf(file_path, preview_only=True)

def on_slider_change(event=None):
    if current_csv_path:
        generate_pdf(current_csv_path, preview_only=True)
    else:
        status_var.set("❌ Load a CSV before adjusting preview settings.")

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
    global current_csv_path
    file_path = event.data.strip("{}")
    if file_path.lower().endswith(".csv"):
        current_csv_path = file_path
        status_var.set("Processing dropped file...")
        generate_pdf(file_path)
        generate_pdf(file_path, preview_only=True)
    else:
        status_var.set("❌ Only CSV files are supported.")

def on_drag_enter(event):
    drop_label.config(bg="#d0ffd0")

def on_drag_leave(event):
    drop_label.config(bg="#e0e0e0")

def update_slider_value(dimension, value):
    global barcode_width, barcode_height
    if dimension == "width":
        barcode_width = int(value)
    elif dimension == "height":
        barcode_height = int(value)
    on_slider_change()

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
Checkbutton(settings_tab, text="🌙 Dark Mode", variable=theme_var, command=toggle_theme, bg="#f4f4f4").pack(pady=(10, 5))

Label(settings_tab, text="Barcode Width", bg="#f4f4f4").pack()
Scale(settings_tab, from_=100, to=400, orient=HORIZONTAL, resolution=10, command=lambda v: update_slider_value("width", v), bg="#f4f4f4").pack(pady=(0, 10))
Label(settings_tab, text="Barcode Height", bg="#f4f4f4").pack()
Scale(settings_tab, from_=30, to=120, orient=HORIZONTAL, resolution=5, command=lambda v: update_slider_value("height", v), bg="#f4f4f4").pack(pady=(0, 10))

Label(settings_tab, text="Text Font Size", bg="#f4f4f4").pack()
Label(settings_tab, text="Barcode Font Size", bg="#f4f4f4").pack()
OptionMenu(settings_tab, barcode_font_size_var, *[str(i) for i in range(6, 30)], command=lambda v: on_slider_change()).pack(pady=(0, 10))
OptionMenu(settings_tab, text_font_size_var, *[str(i) for i in range(6, 20)], command=lambda v: on_slider_change()).pack(pady=(0, 10))

root.mainloop()





