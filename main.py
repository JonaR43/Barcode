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

# Attempt to import tkinterdnd2 for drag-and-drop functionality.
# If not available, set dragdrop_enabled to False and use standard Tk.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
    dragdrop_enabled = True
except ImportError:
    BaseTk = Tk
    dragdrop_enabled = False

# --- Configuration Constants ---
# Define dimensions for different Avery label types (width, height)
label_types = {
    "Avery 5160": (2.625 * inch, 1.0 * inch),
    "Avery 5163": (4.0 * inch, 2.0 * inch)
}
PAGE_WIDTH, PAGE_HEIGHT = letter  # Standard letter page size
COLUMNS = 3  # Number of label columns per page
ROWS = 10  # Number of label rows per page
H_MARGIN = 0.19 * inch  # Horizontal margin on the page
V_MARGIN = 0.5 * inch  # Vertical margin on the page
H_GAP = 0.125 * inch  # Horizontal gap between labels
output_pdf = "avery_labels.pdf"  # Default output PDF filename

preview_image = None  # Global variable to hold the barcode preview image

# --- Core Functionality ---
def generate_pdf(csv_path, preview_only=False):
    """
    Generates a PDF of Avery labels with barcodes from a CSV file.

    Args:
        csv_path (str): The absolute path to the input CSV file.
        preview_only (bool): If True, only generates and displays a single
                             barcode preview on the canvas. Defaults to False.
    """
    global preview_image
    try:
        df = pd.read_csv(csv_path)
        # Validate CSV: ensure it contains a 'code' column
        if "code" not in df.columns:
            status_var.set("❌ CSV must contain a 'code' column.")
            return

        # Get selected label dimensions and text visibility setting
        label_w, label_h = label_types[label_type.get()]
        show_text = show_text_var.get() # This variable is currently not used in the barcode generation options.

        if preview_only:
            # Generate and display a preview of the first barcode
            if not df.empty:
                code = str(df.iloc[0]["code"])
            else:
                status_var.set("No data in CSV for preview.")
                return

            filename = "preview_barcode"
            # Generate barcode image with specified font size for the text below the barcode
            Code128(code, writer=ImageWriter()).save(filename, options={"font_path" : "Calibri.ttf",
                                                                        "font_size": int(barcode_font_size_var.get()),
                                                                        "module_width" : 0.2,
                                                                        "module_height": 8,
                                                                        "quiet_zone": 2.0,
                                                                        "text_distance" : 1})
            img = Image.open(f"{filename}.png")
            # Resize image for better display in the preview canvas
            # img = img.resize((260, 60))
            # Maintain aspect ratio
            w_percent = (260 / float(img.size[0]))
            h_size = int((float(img.size[1]) * float(w_percent)))
            img = img.resize((260, h_size), Image.LANCZOS)
            preview_image = ImageTk.PhotoImage(img) # Store as ImageTk.PhotoImage to prevent garbage collection
            label_canvas.delete("all") # Clear previous canvas content
            # Draw a gray rectangle as a background for the preview
            label_canvas.create_rectangle(0, 0, 300, 100, fill="white", outline="gray")
            label_canvas.create_image(20, 20, anchor="nw", image=preview_image) # Place image on canvas
            img.close()
            os.remove(f"{filename}.png") # Clean up temporary barcode image
            return

        # Initialize PDF canvas for full generation
        total_labels = len(df)
        c = pdf_canvas.Canvas(output_pdf, pagesize=letter)

        # Iterate through each row (barcode data) in the DataFrame
        for idx, row in df.iterrows():
            barcode_data = str(row["code"])
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}"

            # Generate barcode image with specified font size
            Code128(barcode_data, writer=ImageWriter()).save(barcode_filename, options={"font_path" : "Calibri.ttf",
                                                                        "font_size": int(barcode_font_size_var.get()),
                                                                        "module_width" : 0.2,
                                                                        "module_height": 8,
                                                                        "quiet_zone": 2.0,
                                                                        "text_distance": 1})
            time.sleep(0.1)
            img = Image.open(f"{barcode_filename}.png")
            img.save(f"{barcode_filename}.png")
            img.close()

            # Calculate position for the current label on the page
            col = idx % COLUMNS
            row_pos = (idx // COLUMNS) % ROWS
            x = H_MARGIN + col * (label_w + H_GAP)
            y = PAGE_HEIGHT - V_MARGIN - (row_pos + 1) * label_h

            # Add a new page if the current label exceeds the number of labels per page
            if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                c.showPage()

            # Draw the barcode image onto the PDF canvas
            # Adjust coordinates slightly for better centering within the label area
            c.drawImage(f"{barcode_filename}.png", x + 4, y + 5, width=label_w - 8, height=label_h - 10)
            os.remove(f"{barcode_filename}.png")

        c.save() # Save the generated PDF
        status_var.set(f"✅ {total_labels} labels generated.") # Update status
        link_label.config(text="📂 Open PDF", fg="#2196f3") # Change link text and color
        link_label.bind("<Button-1>", lambda e: webbrowser.open(output_pdf)) # Bind click to open PDF

    except Exception as e:
        status_var.set(f"❌ Error: {str(e)}") # Display error message

def select_file():
    """
    Opens a file dialog for the user to select a CSV file.
    Triggers PDF generation and then a preview generation.
    """
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        abs_path = os.path.abspath(file_path)  # Ensure absolute path for consistency
        status_var.set("Processing...")
        generate_pdf(abs_path) # Generate full PDF
        generate_pdf(abs_path, preview_only=True) # Generate preview after full generation

def toggle_theme():
    """
    Toggles between light and dark mode for the application's GUI.
    Updates background and foreground colors of relevant widgets.
    """
    dark = theme_var.get()
    bg = "#2e2e2e" if dark else "#f4f4f4" # Dark or light background
    fg = "#ffffff" if dark else "#000000" # Dark or light foreground
    hint_fg = "#cccccc" if dark else "#666666" # Hint text color

    root.configure(bg=bg) # Update root window background
    # Update colors for various labels
    for widget in [status, link_label, drop_label, preview_label, header]:
        widget.configure(bg=bg, fg=fg)
    subheader.configure(bg=bg, fg=hint_fg) # Subheader has a specific hint color
    settings_tab.configure(bg=bg) # Update tab backgrounds
    generator_tab.configure(bg=bg)
    label_canvas.configure(bg="white") # Canvas always remains white for barcode display

def handle_drop(event):
    """
    Handles file drop events for drag-and-drop functionality.
    Processes the dropped file if it's a CSV.
    """
    file_path = event.data.strip("{}") # Get file path from event data, remove braces
    if file_path.lower().endswith(".csv"):
        abs_path = os.path.abspath(file_path)  # Ensure absolute path
        status_var.set("Processing dropped file...")
        generate_pdf(abs_path) # Generate full PDF
        generate_pdf(abs_path, preview_only=True) # Generate preview
    else:
        status_var.set("❌ Only CSV files are supported.") # Error for unsupported file types

def on_drag_enter(event):
    """
    Callback for drag enter event. Changes drop label background for visual feedback.
    """
    drop_label.config(bg="#d0ffd0") # Light green when dragging over

def on_drag_leave(event):
    """
    Callback for drag leave event. Restores drop label background.
    """
    drop_label.config(bg="#e0e0e0") # Restore original gray when drag leaves

# --- GUI Setup ---
root = BaseTk() # Use TkinterDnD.Tk if available, otherwise standard Tk
root.title("🧾 Avery Barcode Label Generator")
root.geometry("520x640") # Set initial window size
root.configure(bg="#f4f4f4") # Default light background

# --- Tkinter Variables ---
theme_var = BooleanVar()
theme_var.set(False) # Default to light mode
show_text_var = BooleanVar(value=True) # Default to showing text under barcode
barcode_font_size_var = StringVar(value="10")  # Default barcode text font size
label_type = StringVar(value="Avery 5160") # Default label type
status_var = StringVar()
status_var.set("Upload or drag a CSV with a 'code' column.") # Initial status message

# --- Notebook (Tabs) ---
notebook = ttk.Notebook(root)
notebook.pack(expand=1, fill="both", pady=5)

# Create two tabs: Generator and Settings
generator_tab = Frame(notebook, bg="#f4f4f4")
settings_tab = Frame(notebook, bg="#f4f4f4")
notebook.add(generator_tab, text="🧾 Generator")
notebook.add(settings_tab, text="⚙️ Settings")

# --- Generator Tab Layout ---
header = Label(generator_tab, text="Barcode Label Generator", font=("Helvetica", 16, "bold"), bg="#f4f4f4")
header.pack(pady=(15, 3))

subheader = Label(generator_tab, text="CSV must include a 'code' column", font=("Helvetica", 10), bg="#f4f4f4", fg="#666")
subheader.pack()

frame = Frame(generator_tab, bg="#f4f4f4")
frame.pack(pady=(10, 5))

# Button to choose CSV file
btn = Button(frame, text="📂 Choose CSV", font=("Helvetica", 11), command=select_file, bg="#4caf50", fg="white", padx=12, pady=6)
btn.grid(row=0, column=0, padx=5)

# Label to display link to open generated PDF
link_label = Label(generator_tab, text="", font=("Helvetica", 10, "underline"), bg="#f4f4f4", cursor="hand2")
link_label.pack(pady=(5, 0))

# Label to display application status messages
status = Label(generator_tab, textvariable=status_var, wraplength=400, bg="#f4f4f4", font=("Helvetica", 9), fg="#333")
status.pack(pady=10)

# Label for drag-and-drop area
drop_label = Label(generator_tab, text="⬇️ Drop CSV file here", font=("Helvetica", 10, "italic"), bg="#e0e0e0", fg="#444", relief="groove", width=40, height=3, bd=2)
drop_label.pack(pady=(0, 10))

# Enable drag-and-drop if tkinterdnd2 was successfully imported
if dragdrop_enabled:
    drop_label.drop_target_register(DND_FILES) # Register label as a drop target for files
    drop_label.dnd_bind("<<Drop>>", handle_drop) # Bind drop event
    drop_label.dnd_bind("<<DragEnter>>", on_drag_enter) # Bind drag enter for visual feedback
    drop_label.dnd_bind("<<DragLeave>>", on_drag_leave) # Bind drag leave for visual feedback

# Label for barcode preview section
preview_label = Label(generator_tab, text="🔍 Label Preview", font=("Calibri", 10, "bold"), bg="#f4f4f4")
preview_label.pack()

# Canvas to display the barcode preview
label_canvas = Canvas(generator_tab, width=300, height=100, bg="white", bd=1, relief="sunken")
label_canvas.pack(pady=(5, 20))

# --- Settings Tab Layout ---
Label(settings_tab, text="⚙️ Settings", font=("Calibri", 14, "bold"), bg="#f4f4f4").pack(pady=(20, 5))

# Checkbox to toggle text visibility under barcode (currently not implemented in barcode generation options)
Checkbutton(settings_tab, text="Show Text Below Barcode", variable=show_text_var, bg="#f4f4f4").pack(pady=5)

# Option menu for selecting label type
Label(settings_tab, text="Label Type:", bg="#f4f4f4").pack()
OptionMenu(settings_tab, label_type, *label_types.keys()).pack(pady=5)

# Checkbox to toggle dark mode theme
Checkbutton(settings_tab, text="🌙 Dark Mode", variable=theme_var, command=toggle_theme, bg="#f4f4f4").pack(pady=(10, 20))

# Option menu for selecting barcode font size
Label(settings_tab, text="Barcode Font Size (under barcode)", bg="#f4f4f4").pack()
OptionMenu(settings_tab, barcode_font_size_var, *[str(i) for i in range(6, 30)]).pack(pady=(0, 10))

# Start the Tkinter event loop
root.mainloop()