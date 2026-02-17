import os
import sys
import time
import sqlite3
import json
from datetime import datetime
import pandas as pd
import webbrowser
import customtkinter as ctk
from tkinter import filedialog, Canvas, StringVar, BooleanVar, simpledialog, messagebox
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PIL import Image, ImageTk

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

# Font path for barcode generation - look in exe's folder
if getattr(sys, 'frozen', False):
    # Running as exe - look in same folder as exe
    BARCODE_FONT = os.path.join(os.path.dirname(sys.executable), "calibri.ttf")
else:
    # Running as script - look in script's folder
    BARCODE_FONT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "calibri.ttf")

# Set appearance and theme
ctk.set_appearance_mode("dark")  # "dark", "light", or "system"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

# Label definitions
# Avery 5160 spec:
#   Label:          2.625" W x 1.0" H
#   Columns/Rows:   3 x 10  (30 labels per sheet)
#   Left/Right margin: 0.1875"
#   Top/Bottom margin: 0.5"
#   Horizontal pitch:  2.75"  (center-to-center)
#   Vertical pitch:    1.0"   (center-to-center, no gap)
#   Horizontal gap:    0.125" (2.75 pitch - 2.625 label width)
label_types = {
    "Avery 5160": (2.625 * inch, 1.0 * inch),
    "Avery 5163": (4.0 * inch, 2.0 * inch)
}
PAGE_WIDTH, PAGE_HEIGHT = letter
COLUMNS = 3
ROWS = 10
H_PITCH = 2.75 * inch      # Avery 5160 horizontal pitch (center to center)
H_GAP = 0.125 * inch       # Avery 5160 horizontal gap between labels (2.75 - 2.625)
output_pdf = "avery_labels.pdf"

# Database setup
# Database path - save next to exe when frozen, next to script when developing
if getattr(sys, 'frozen', False):
    DB_PATH = os.path.join(os.path.dirname(sys.executable), "barcodes.db")
else:
    DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "barcodes.db")

def init_database():
    """Initialize the SQLite database with required tables."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Create categories table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Create barcode_batches table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS barcode_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER,
            name TEXT NOT NULL,
            prefix TEXT,
            number TEXT,
            start_letter TEXT,
            amount INTEGER,
            display_text TEXT,
            codes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (category_id) REFERENCES categories (id)
        )
    ''')

    conn.commit()
    conn.close()

def get_categories():
    """Get all categories from database."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT id, name FROM categories ORDER BY name')
    categories = cursor.fetchall()
    conn.close()
    return categories

def add_category(name):
    """Add a new category."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('INSERT INTO categories (name) VALUES (?)', (name,))
        conn.commit()
        category_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        category_id = None
    conn.close()
    return category_id

def delete_category(category_id):
    """Delete a category and its batches."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM barcode_batches WHERE category_id = ?', (category_id,))
    cursor.execute('DELETE FROM categories WHERE id = ?', (category_id,))
    conn.commit()
    conn.close()

def save_barcode_batch(category_id, name, prefix, number, start_letter, amount, display_text, codes):
    """Save a barcode batch to the database."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO barcode_batches
        (category_id, name, prefix, number, start_letter, amount, display_text, codes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (category_id, name, prefix, number, start_letter, amount, display_text, json.dumps(codes)))
    conn.commit()
    batch_id = cursor.lastrowid
    conn.close()
    return batch_id

def get_batches_by_category(category_id):
    """Get all barcode batches for a category."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, name, prefix, number, start_letter, amount, display_text, codes, created_at
        FROM barcode_batches WHERE category_id = ? ORDER BY created_at DESC
    ''', (category_id,))
    batches = cursor.fetchall()
    conn.close()
    return batches

def get_batch_by_id(batch_id):
    """Get a specific barcode batch."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, name, prefix, number, start_letter, amount, display_text, codes
        FROM barcode_batches WHERE id = ?
    ''', (batch_id,))
    batch = cursor.fetchone()
    conn.close()
    return batch

def delete_batch(batch_id):
    """Delete a barcode batch."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM barcode_batches WHERE id = ?', (batch_id,))
    conn.commit()
    conn.close()

# Initialize database on startup
init_database()

preview_image = None

def generate_pdf(csv_path, preview_only=False):
    global preview_image
    try:
        df = pd.read_csv(csv_path)
        if "code" not in df.columns:
            status_var.set("CSV must contain a 'code' column.")
            return

        label_w, label_h = label_types[label_type_var.get()]
        show_text = show_text_var.get()

        if preview_only:
            code = str(df.iloc[0]["code"])
            filename = "preview_barcode"
            Code128(code, writer=ImageWriter()).save(filename, options={
                "font_path": BARCODE_FONT,
                "font_size": 10,
                "module_width": 0.2,
                "module_height": 6,
                "quiet_zone": 2.0,
                "text_distance": 3,
                "write_text": True
            })
            img = Image.open(f"{filename}.png")
            img.load()  # Force load the image data
            img = img.resize((260, 60))
            preview_image = ImageTk.PhotoImage(img)
            csv_canvas.delete("all")
            csv_canvas.create_rectangle(0, 0, 300, 100, fill="white", outline="#cccccc")
            csv_canvas.create_image(20, 20, anchor="nw", image=preview_image)
            if show_text:
                csv_canvas.create_text(150, 90, text=code, font=("Helvetica", 8), fill="black")
            img.close()
            time.sleep(0.05)
            try:
                os.remove(f"{filename}.png")
            except PermissionError:
                pass  # Skip if file is locked
            return

        total_labels = len(df)
        c = pdf_canvas.Canvas(output_pdf, pagesize=letter)

        for idx, row in df.iterrows():
            barcode_data = str(row["code"])
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}.png"

            try:
                Code128(barcode_data, writer=ImageWriter()).save(barcode_base, options={
                    "font_path": BARCODE_FONT,
                    "font_size": 10,
                    "module_width": 0.2,
                    "module_height": 6,
                    "quiet_zone": 2.0,
                    "text_distance": 5,
                    "write_text": False
                })
                time.sleep(0.1)
                img = Image.open(barcode_filename)
                img.load()  # Force load the image data
                img.save(barcode_filename)
                img.close()
                time.sleep(0.05)  # Brief wait to ensure file is closed

                col = idx % COLUMNS
                row_pos = (idx // COLUMNS) % ROWS
                h_margin = float(h_margin_var.get()) * inch
                v_margin = float(v_margin_var.get()) * inch
                x = h_margin + col * H_PITCH
                y = PAGE_HEIGHT - v_margin - (row_pos + 1) * label_h

                if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                    c.showPage()

                if show_label_outlines_var.get():
                    c.setStrokeColorRGB(0.8, 0.2, 0.2)
                    c.setLineWidth(0.5)
                    c.rect(x, y, label_w, label_h)
                    c.setStrokeColorRGB(0, 0, 0)

                code_fs = int(code_font_size_var.get())
                c.drawImage(barcode_filename, x + 4, y + 23, width=label_w - 8, height=label_h - 31)
                c.setFont("Helvetica", code_fs)
                c.drawCentredString(x + label_w / 2, y + 14, barcode_data)

                # Ensure file is closed before deleting
                try:
                    os.remove(barcode_filename)
                except PermissionError:
                    time.sleep(0.1)  # Wait a bit longer
                    try:
                        os.remove(barcode_filename)
                    except:
                        pass  # Skip if still can't delete
            except Exception as e:
                status_var.set(f"Error generating barcode {idx}: {str(e)}")
                continue

        c.save()
        status_var.set(f"✓ {total_labels} labels generated successfully!")
        csv_link_btn.configure(state="normal")

    except Exception as e:
        status_var.set(f"Error: {str(e)}")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        status_var.set("Processing...")
        generate_pdf(file_path)
        generate_pdf(file_path, preview_only=True)

def open_pdf():
    webbrowser.open(output_pdf)

def generate_from_codes(codes, preview_only=False):
    """Generates barcodes from a list of codes."""
    global preview_image
    try:
        if not codes:
            status_var.set("No codes to generate.")
            return

        label_w, label_h = label_types[label_type_var.get()]
        show_text = show_text_var.get()
        display_text = display_text_var.get().strip()  # Get custom label text

        if preview_only:
            code = str(codes[0])
            filename = "preview_barcode"
            Code128(code, writer=ImageWriter()).save(filename, options={
                "font_path": BARCODE_FONT,
                "font_size": 10,
                "module_width": 0.2,
                "module_height": 6,
                "quiet_zone": 2.0,
                "text_distance": 3,
                "write_text": True
            })
            img = Image.open(f"{filename}.png")
            img.load()  # Force load the image data
            img = img.resize((260, 60))
            preview_image = ImageTk.PhotoImage(img)
            manual_canvas.delete("all")
            manual_canvas.create_rectangle(0, 0, 300, 100, fill="white", outline="#cccccc")
            manual_canvas.create_image(20, 20, anchor="nw", image=preview_image)
            if show_text and display_text:
                manual_canvas.create_text(150, 90, text=display_text, font=("Helvetica", 8), fill="black")
            img.close()
            time.sleep(0.05)
            try:
                os.remove(f"{filename}.png")
            except PermissionError:
                pass  # Skip if file is locked
            return

        total_labels = len(codes)
        c = pdf_canvas.Canvas(output_pdf, pagesize=letter)

        # Reset progress bar
        progress_bar.set(0)
        root.update()

        for idx, barcode_data in enumerate(codes):
            barcode_base = f"barcode_{idx}"
            barcode_filename = f"{barcode_base}.png"

            try:
                Code128(str(barcode_data), writer=ImageWriter()).save(barcode_base, options={
                    "font_path": BARCODE_FONT,
                    "font_size": 10,
                    "module_width": 0.2,
                    "module_height": 6,
                    "quiet_zone": 2.0,
                    "text_distance": 5,
                    "write_text": False
                })
                time.sleep(0.05)  # Reduced sleep time
                img = Image.open(barcode_filename)
                img.load()  # Force load the image data
                img.save(barcode_filename)
                img.close()
                time.sleep(0.05)  # Brief wait to ensure file is closed

                col = idx % COLUMNS
                row_pos = (idx // COLUMNS) % ROWS
                h_margin = float(h_margin_var.get()) * inch
                v_margin = float(v_margin_var.get()) * inch
                x = h_margin + col * H_PITCH
                y = PAGE_HEIGHT - v_margin - (row_pos + 1) * label_h

                if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                    c.showPage()

                if show_label_outlines_var.get():
                    c.setStrokeColorRGB(0.8, 0.2, 0.2)
                    c.setLineWidth(0.5)
                    c.rect(x, y, label_w, label_h)
                    c.setStrokeColorRGB(0, 0, 0)

                code_fs = int(code_font_size_var.get())
                c.drawImage(barcode_filename, x + 4, y + 23, width=label_w - 8, height=label_h - 31)
                c.setFont("Helvetica", code_fs)
                c.drawCentredString(x + label_w / 2, y + 14, barcode_data)
                if display_text:
                    c.drawCentredString(x + label_w / 2, y + 5, display_text)

                # Ensure file is closed before deleting
                try:
                    os.remove(barcode_filename)
                except PermissionError:
                    time.sleep(0.1)  # Wait a bit longer
                    try:
                        os.remove(barcode_filename)
                    except:
                        pass  # Skip if still can't delete
            except Exception as e:
                status_var.set(f"Error generating barcode {idx}: {str(e)}")
                continue

            # Update progress bar
            progress = (idx + 1) / total_labels
            progress_bar.set(progress)
            status_var.set(f"Generating... {idx + 1}/{total_labels}")
            root.update()

        c.save()
        progress_bar.set(1)
        status_var.set(f"✓ {total_labels} labels generated successfully!")
        manual_link_btn.configure(state="normal")

    except Exception as e:
        status_var.set(f"Error: {str(e)}")

def index_to_letters(n):
    """Converts a number to Excel-style letter sequence (0=A, 25=Z, 26=AA, 27=AB, etc.)"""
    result = ""
    while True:
        result = chr(ord('A') + (n % 26)) + result
        n = n // 26 - 1
        if n < 0:
            break
    return result

def letters_to_index(letters):
    """Converts Excel-style letter sequence to number (A=0, Z=25, AA=26, AB=27, etc.)"""
    letters = letters.upper().strip()
    result = 0
    for char in letters:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def get_codes_list():
    """Generate list of codes based on current inputs without creating barcodes."""
    prefix = prefix_var.get().strip()
    number = number_var.get()
    start_letter = start_letter_var.get().strip().upper()

    try:
        amount = int(amount_var.get())
    except ValueError:
        return None, "Amount must be a number."

    if not prefix:
        return None, "Please enter a prefix."

    if amount < 1:
        return None, "Amount must be at least 1."

    if not start_letter or not start_letter.isalpha():
        return None, "Starting letter must be letters only (A-Z)."

    # Calculate starting index from letter
    start_index = letters_to_index(start_letter)

    # Generate codes list
    codes = []
    for i in range(amount):
        letter = index_to_letters(start_index + i)
        code = f"{prefix}D{number}{letter}"
        codes.append(code)

    return codes, None

def preview_codes():
    """Show preview of codes that will be generated."""
    codes, error = get_codes_list()
    if error:
        status_var.set(error)
        return

    # Update preview textbox
    preview_textbox.configure(state="normal")
    preview_textbox.delete("1.0", "end")
    preview_textbox.insert("1.0", "\n".join(codes))
    preview_textbox.configure(state="disabled")
    status_var.set(f"Preview: {len(codes)} codes")

def generate_manual():
    """Generates barcodes from manual entry inputs."""
    codes, error = get_codes_list()
    if error:
        status_var.set(error)
        return

    status_var.set("Processing...")
    generate_from_codes(codes)
    generate_from_codes(codes, preview_only=True)

def toggle_multi_batch_ui():
    if multi_batch_mode_var.get():
        multi_batch_frame.pack(fill="x", pady=(0, 5), before=codes_preview_label)
    else:
        multi_batch_frame.pack_forget()

def refresh_batch_queue():
    batch_queue_box.configure(state="normal")
    batch_queue_box.delete("1.0", "end")
    total = sum(len(b["codes"]) for b in multi_batches)
    for i, b in enumerate(multi_batches):
        batch_queue_box.insert("end", f"{i+1}. {b['summary']}\n")
    batch_queue_box.configure(state="disabled")
    if multi_batches:
        status_var.set(f"Queue: {len(multi_batches)} batch(es) — {total} total labels")

def add_to_batch_queue():
    codes, error = get_codes_list()
    if error:
        status_var.set(error)
        return
    prefix = prefix_var.get().strip()
    number = number_var.get()
    start = start_letter_var.get().strip().upper()
    display_text = display_text_var.get().strip()
    end_letter = index_to_letters(letters_to_index(start) + len(codes) - 1)
    summary = f"{prefix}D{number}  {start} → {end_letter}  ({len(codes)} labels)"
    if display_text:
        summary += f'  [{display_text}]'
    multi_batches.append({"codes": codes, "display_text": display_text, "summary": summary})
    refresh_batch_queue()

def clear_batch_queue():
    multi_batches.clear()
    refresh_batch_queue()
    status_var.set("Queue cleared.")

def generate_all_batches():
    if not multi_batches:
        status_var.set("Queue is empty — add batches first.")
        return

    all_pairs = [(code, b["display_text"]) for b in multi_batches for code in b["codes"]]
    total_labels = len(all_pairs)
    label_w, label_h = label_types[label_type_var.get()]
    c = pdf_canvas.Canvas(output_pdf, pagesize=letter)
    progress_bar.set(0)
    root.update()

    for idx, (barcode_data, display_text) in enumerate(all_pairs):
        barcode_base = f"barcode_{idx}"
        barcode_filename = f"{barcode_base}.png"
        try:
            Code128(str(barcode_data), writer=ImageWriter()).save(barcode_base, options={
                "font_path": BARCODE_FONT,
                "font_size": 10,
                "module_width": 0.2,
                "module_height": 6,
                "quiet_zone": 2.0,
                "text_distance": 5,
                "write_text": False
            })
            time.sleep(0.05)
            img = Image.open(barcode_filename)
            img.load()
            img.save(barcode_filename)
            img.close()
            time.sleep(0.05)

            col = idx % COLUMNS
            row_pos = (idx // COLUMNS) % ROWS
            h_margin = float(h_margin_var.get()) * inch
            v_margin = float(v_margin_var.get()) * inch
            x = h_margin + col * H_PITCH
            y = PAGE_HEIGHT - v_margin - (row_pos + 1) * label_h

            if idx > 0 and idx % (COLUMNS * ROWS) == 0:
                c.showPage()

            if show_label_outlines_var.get():
                c.setStrokeColorRGB(0.8, 0.2, 0.2)
                c.setLineWidth(0.5)
                c.rect(x, y, label_w, label_h)
                c.setStrokeColorRGB(0, 0, 0)

            code_fs = int(code_font_size_var.get())
            c.drawImage(barcode_filename, x + 4, y + 23, width=label_w - 8, height=label_h - 31)
            c.setFont("Helvetica", code_fs)
            c.drawCentredString(x + label_w / 2, y + 14, barcode_data)
            if display_text:
                c.drawCentredString(x + label_w / 2, y + 5, display_text)

            try:
                os.remove(barcode_filename)
            except PermissionError:
                time.sleep(0.1)
                try:
                    os.remove(barcode_filename)
                except:
                    pass
        except Exception as e:
            status_var.set(f"Error generating barcode {idx}: {str(e)}")
            continue

        progress_bar.set((idx + 1) / total_labels)
        status_var.set(f"Generating... {idx + 1}/{total_labels}")
        root.update()

    c.save()
    progress_bar.set(1)
    status_var.set(f"✓ {total_labels} labels from {len(multi_batches)} batches generated!")
    manual_link_btn.configure(state="normal")

def change_appearance_mode(new_mode):
    ctk.set_appearance_mode(new_mode)

# === Library Functions ===
def refresh_categories():
    """Refresh the category list in the Library tab."""
    categories = get_categories()
    category_listbox.configure(state="normal")
    category_listbox.delete("1.0", "end")
    for cat_id, cat_name in categories:
        category_listbox.insert("end", f"{cat_name}\n")
    category_listbox.configure(state="disabled")

def add_new_category():
    """Add a new category via dialog."""
    name = simpledialog.askstring("New Category", "Enter category name:")
    if name:
        result = add_category(name.strip())
        if result:
            refresh_categories()
            status_var.set(f"Category '{name}' created.")
        else:
            status_var.set("Category already exists.")

def on_category_select(event=None):
    """Handle category selection."""
    try:
        # Get selected line
        index = category_listbox.index("insert")
        line_start = f"{int(float(index))}.0"
        line_end = f"{int(float(index))}.end"
        selected_text = category_listbox.get(line_start, line_end).strip()

        if selected_text:
            # Find category ID
            categories = get_categories()
            for cat_id, cat_name in categories:
                if cat_name == selected_text:
                    selected_category_var.set(str(cat_id))
                    refresh_batches(cat_id)
                    break
    except:
        pass

def refresh_batches(category_id):
    """Refresh the batches list for selected category."""
    batches = get_batches_by_category(category_id)
    batch_listbox.configure(state="normal")
    batch_listbox.delete("1.0", "end")
    for batch in batches:
        batch_id, name, prefix, number, start, amount, display, codes, created = batch
        batch_listbox.insert("end", f"[{batch_id}] {name} ({amount} codes)\n")
    batch_listbox.configure(state="disabled")

def save_to_library():
    """Save current barcode batch to library."""
    codes, error = get_codes_list()
    if error:
        status_var.set(error)
        return

    categories = get_categories()
    if not categories:
        status_var.set("Please create a category first.")
        return

    # Ask for batch name
    name = simpledialog.askstring("Save Batch", "Enter a name for this batch:")
    if not name:
        return

    # Ask for category (simple approach - use first category or ask)
    cat_names = [c[1] for c in categories]
    # For simplicity, use first category or let user type
    cat_name = simpledialog.askstring("Select Category",
                                      f"Enter category name:\n({', '.join(cat_names)})")
    if not cat_name:
        return

    # Find category ID
    cat_id = None
    for c_id, c_name in categories:
        if c_name.lower() == cat_name.lower():
            cat_id = c_id
            break

    if cat_id is None:
        # Create new category
        cat_id = add_category(cat_name)
        if cat_id is None:
            status_var.set("Failed to create category.")
            return

    # Save batch
    prefix = prefix_var.get().strip()
    number = number_var.get()
    start_letter = start_letter_var.get().strip().upper()
    amount = int(amount_var.get())
    display_text = display_text_var.get().strip()

    save_barcode_batch(cat_id, name, prefix, number, start_letter, amount, display_text, codes)
    refresh_categories()
    status_var.set(f"Saved '{name}' to library!")

def load_batch_by_id(batch_id):
    """Load a batch by its ID."""
    try:
        batch = get_batch_by_id(batch_id)
        if batch:
            _, name, prefix, number, start_letter, amount, display_text, codes_json = batch

            # Update UI fields
            prefix_var.set(prefix or "")
            number_var.set(number or "1")
            start_letter_var.set(start_letter or "A")
            amount_var.set(str(amount or 1))
            display_text_var.set(display_text or "")

            # Switch to Manual Entry tab and show preview
            tabview.set("Manual Entry")
            preview_codes()
            status_var.set(f"Loaded '{name}' from library.")
            return True
    except Exception as e:
        status_var.set(f"Error loading batch: {str(e)}")
    return False

def on_batch_double_click(event=None):
    """Handle double-click on batch to load it directly."""
    try:
        index = batch_listbox.index("insert")
        line_start = f"{int(float(index))}.0"
        line_end = f"{int(float(index))}.end"
        selected = batch_listbox.get(line_start, line_end).strip()

        if selected and "[" in selected:
            batch_id = int(selected.split("]")[0].replace("[", ""))
            load_batch_by_id(batch_id)
    except:
        pass

def load_from_library():
    """Load selected batch from library."""
    try:
        index = batch_listbox.index("insert")
        line_start = f"{int(float(index))}.0"
        line_end = f"{int(float(index))}.end"
        selected = batch_listbox.get(line_start, line_end).strip()

        if not selected or "[" not in selected:
            status_var.set("Click on a batch first, then Load.")
            return

        batch_id = int(selected.split("]")[0].replace("[", ""))
        load_batch_by_id(batch_id)
    except Exception as e:
        status_var.set(f"Error: {str(e)}")

def delete_selected_batch():
    """Delete the selected batch."""
    try:
        batch_listbox.configure(state="normal")
        selection = batch_listbox.get("1.0", "end").strip().split("\n")
        batch_listbox.configure(state="disabled")

        index = batch_listbox.index("insert")
        line_num = int(float(index)) - 1
        if line_num < 0 or line_num >= len(selection):
            return

        selected = selection[line_num].strip()
        if not selected:
            return

        batch_id = int(selected.split("]")[0].replace("[", ""))

        if messagebox.askyesno("Delete Batch", "Are you sure you want to delete this batch?"):
            delete_batch(batch_id)
            cat_id = selected_category_var.get()
            if cat_id:
                refresh_batches(int(cat_id))
            status_var.set("Batch deleted.")
    except:
        pass

# === GUI Setup ===
root = ctk.CTk()
root.title("Barcode Label Generator")
root.geometry("550x780")

# Variables
show_text_var = BooleanVar(value=True)
show_label_outlines_var = BooleanVar(value=False)  # Debug: draw label cell borders
code_font_size_var = StringVar(value="11")           # Font size for barcode code text
h_margin_var = StringVar(value="0.1875")            # Avery 5160 left margin in inches
v_margin_var = StringVar(value="0.45")               # Avery 5160 top margin in inches
multi_batch_mode_var = BooleanVar(value=False)       # Multi-prefix batch mode
multi_batches = []                                   # [{'codes': [...], 'display_text': '', 'summary': ''}]
label_type_var = StringVar(value="Avery 5160")
status_var = StringVar(value="Ready to generate barcodes")
prefix_var = StringVar()
number_var = StringVar(value="1")
amount_var = StringVar(value="1")
start_letter_var = StringVar(value="A")  # Starting letter for sequence
display_text_var = StringVar()  # Custom text to show below barcode
selected_category_var = StringVar()  # For library category selection

# === Main Container ===
main_frame = ctk.CTkFrame(root, fg_color="transparent")
main_frame.pack(fill="both", expand=True, padx=20, pady=20)

# === Header ===
header_label = ctk.CTkLabel(main_frame, text="Barcode Label Generator",
                            font=ctk.CTkFont(size=24, weight="bold"))
header_label.pack(pady=(0, 20))

# === Tabview ===
tabview = ctk.CTkTabview(main_frame, width=500, height=480)
tabview.pack(fill="both", expand=True)

csv_tab = tabview.add("CSV Import")
manual_tab = tabview.add("Manual Entry")
library_tab = tabview.add("Library")
settings_tab = tabview.add("Settings")

# ==================== CSV Import Tab ====================
csv_frame = ctk.CTkFrame(csv_tab, fg_color="transparent")
csv_frame.pack(fill="both", expand=True, padx=10, pady=10)

ctk.CTkLabel(csv_frame, text="Import barcodes from CSV file",
             font=ctk.CTkFont(size=14), text_color="gray").pack(pady=(0, 15))

# File selection button
csv_btn = ctk.CTkButton(csv_frame, text="Choose CSV File",
                        command=select_file, width=200, height=40,
                        font=ctk.CTkFont(size=14))
csv_btn.pack(pady=10)

# Open PDF button
csv_link_btn = ctk.CTkButton(csv_frame, text="Open Generated PDF",
                             command=open_pdf, width=200, height=35,
                             font=ctk.CTkFont(size=12), state="disabled",
                             fg_color="transparent", border_width=2)
csv_link_btn.pack(pady=10)

# Preview section
ctk.CTkLabel(csv_frame, text="Preview", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=(20, 5))
csv_canvas = Canvas(csv_frame, width=300, height=100, bg="white",
                    highlightthickness=1, highlightbackground="#cccccc")
csv_canvas.pack(pady=5)

# ==================== Manual Entry Tab ====================
manual_frame = ctk.CTkFrame(manual_tab, fg_color="transparent")
manual_frame.pack(fill="both", expand=True, padx=10, pady=10)

ctk.CTkLabel(manual_frame, text="Format: {Prefix}D{Number}{Letter}",
             font=ctk.CTkFont(size=14), text_color="gray").pack(pady=(0, 15))

# Input fields frame
input_frame = ctk.CTkFrame(manual_frame, fg_color="transparent")
input_frame.pack(pady=10)

# Prefix
ctk.CTkLabel(input_frame, text="Prefix:", font=ctk.CTkFont(size=13)).grid(row=0, column=0, padx=10, pady=8, sticky="e")
prefix_entry = ctk.CTkEntry(input_frame, textvariable=prefix_var, width=180, height=35,
                            placeholder_text="Enter prefix...")
prefix_entry.grid(row=0, column=1, padx=10, pady=8)

# Number
ctk.CTkLabel(input_frame, text="Number:", font=ctk.CTkFont(size=13)).grid(row=1, column=0, padx=10, pady=8, sticky="e")
number_menu = ctk.CTkOptionMenu(input_frame, variable=number_var, values=["1", "2", "3", "4"],
                                width=180, height=35)
number_menu.grid(row=1, column=1, padx=10, pady=8)

# Amount
ctk.CTkLabel(input_frame, text="Amount:", font=ctk.CTkFont(size=13)).grid(row=2, column=0, padx=10, pady=8, sticky="e")
amount_entry = ctk.CTkEntry(input_frame, textvariable=amount_var, width=180, height=35,
                            placeholder_text="Number of labels...")
amount_entry.grid(row=2, column=1, padx=10, pady=8)

# Starting Letter
ctk.CTkLabel(input_frame, text="Start From:", font=ctk.CTkFont(size=13)).grid(row=3, column=0, padx=10, pady=8, sticky="e")
start_letter_entry = ctk.CTkEntry(input_frame, textvariable=start_letter_var, width=180, height=35,
                                  placeholder_text="A, B, AA, etc...")
start_letter_entry.grid(row=3, column=1, padx=10, pady=8)
ctk.CTkLabel(input_frame, text="(A, AA, BA...)", font=ctk.CTkFont(size=11),
             text_color="gray").grid(row=3, column=2, padx=5, sticky="w")

# Display Text (shown below barcode)
ctk.CTkLabel(input_frame, text="Label Text:", font=ctk.CTkFont(size=13)).grid(row=4, column=0, padx=10, pady=8, sticky="e")
display_text_entry = ctk.CTkEntry(input_frame, textvariable=display_text_var, width=180, height=35,
                                  placeholder_text="Text below barcode...")
display_text_entry.grid(row=4, column=1, padx=10, pady=8)
ctk.CTkLabel(input_frame, text="(optional)", font=ctk.CTkFont(size=11),
             text_color="gray").grid(row=4, column=2, padx=5, sticky="w")

# Buttons frame
buttons_frame = ctk.CTkFrame(manual_frame, fg_color="transparent")
buttons_frame.pack(pady=10)

# Preview button
preview_btn = ctk.CTkButton(buttons_frame, text="Preview Codes",
                            command=preview_codes, width=140, height=40,
                            font=ctk.CTkFont(size=13),
                            fg_color="transparent", border_width=2)
preview_btn.pack(side="left", padx=5)

# Generate button
generate_btn = ctk.CTkButton(buttons_frame, text="Generate Barcodes",
                             command=generate_manual, width=140, height=40,
                             font=ctk.CTkFont(size=13, weight="bold"))
generate_btn.pack(side="left", padx=5)

# Multi-batch mode checkbox
multi_batch_check = ctk.CTkCheckBox(manual_frame, text="Combine multiple prefixes",
                                    variable=multi_batch_mode_var,
                                    command=toggle_multi_batch_ui,
                                    font=ctk.CTkFont(size=13))
multi_batch_check.pack(pady=(10, 0))

# Multi-batch collapsible frame (hidden by default)
multi_batch_frame = ctk.CTkFrame(manual_frame)
# Not packed yet — shown via toggle_multi_batch_ui()

ctk.CTkLabel(multi_batch_frame, text="Batch Queue:",
             font=ctk.CTkFont(size=12, weight="bold")).pack(pady=(8, 3))
batch_queue_box = ctk.CTkTextbox(multi_batch_frame, width=380, height=90,
                                 state="disabled", font=ctk.CTkFont(size=11))
batch_queue_box.pack(padx=10, pady=(0, 5))

batch_ctrl_frame = ctk.CTkFrame(multi_batch_frame, fg_color="transparent")
batch_ctrl_frame.pack(pady=(0, 5))
ctk.CTkButton(batch_ctrl_frame, text="Add to Queue",
              command=add_to_batch_queue, width=130, height=35,
              font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=5)
ctk.CTkButton(batch_ctrl_frame, text="Clear Queue",
              command=clear_batch_queue, width=110, height=35,
              font=ctk.CTkFont(size=12),
              fg_color="transparent", border_width=2).pack(side="left", padx=5)

ctk.CTkButton(multi_batch_frame, text="Generate All Batches",
              command=generate_all_batches, width=220, height=40,
              font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(0, 10))

# Codes preview textbox
codes_preview_label = ctk.CTkLabel(manual_frame, text="Codes Preview:", font=ctk.CTkFont(size=12, weight="bold"))
codes_preview_label.pack(pady=(10, 5))
preview_textbox = ctk.CTkTextbox(manual_frame, width=300, height=80, state="disabled",
                                 font=ctk.CTkFont(size=11))
preview_textbox.pack(pady=5)

# Open PDF button
manual_link_btn = ctk.CTkButton(manual_frame, text="Open Generated PDF",
                                command=open_pdf, width=200, height=35,
                                font=ctk.CTkFont(size=12), state="disabled",
                                fg_color="transparent", border_width=2)
manual_link_btn.pack(pady=5)

# Preview section
ctk.CTkLabel(manual_frame, text="Preview", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=(15, 5))
manual_canvas = Canvas(manual_frame, width=300, height=100, bg="white",
                       highlightthickness=1, highlightbackground="#cccccc")
manual_canvas.pack(pady=5)

# ==================== Library Tab ====================
library_frame = ctk.CTkFrame(library_tab, fg_color="transparent")
library_frame.pack(fill="both", expand=True, padx=10, pady=10)

ctk.CTkLabel(library_frame, text="Saved Barcodes Library",
             font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(0, 10))

# Split frame for categories and batches
split_frame = ctk.CTkFrame(library_frame, fg_color="transparent")
split_frame.pack(fill="both", expand=True)

# Categories section (left)
cat_frame = ctk.CTkFrame(split_frame)
cat_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

ctk.CTkLabel(cat_frame, text="Categories", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=5)
category_listbox = ctk.CTkTextbox(cat_frame, width=150, height=150, font=ctk.CTkFont(size=12))
category_listbox.pack(pady=5, padx=5, fill="both", expand=True)
category_listbox.bind("<Button-1>", on_category_select)

cat_btn_frame = ctk.CTkFrame(cat_frame, fg_color="transparent")
cat_btn_frame.pack(pady=5)
ctk.CTkButton(cat_btn_frame, text="+ Add", command=add_new_category, width=70, height=30).pack(side="left", padx=2)

# Batches section (right)
batch_frame = ctk.CTkFrame(split_frame)
batch_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

ctk.CTkLabel(batch_frame, text="Saved Batches", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=5)
ctk.CTkLabel(batch_frame, text="(double-click to load)", font=ctk.CTkFont(size=10), text_color="gray").pack()
batch_listbox = ctk.CTkTextbox(batch_frame, width=200, height=150, font=ctk.CTkFont(size=11))
batch_listbox.pack(pady=5, padx=5, fill="both", expand=True)
batch_listbox.bind("<Double-Button-1>", on_batch_double_click)

batch_btn_frame = ctk.CTkFrame(batch_frame, fg_color="transparent")
batch_btn_frame.pack(pady=5)
ctk.CTkButton(batch_btn_frame, text="Load", command=load_from_library, width=70, height=30).pack(side="left", padx=2)
ctk.CTkButton(batch_btn_frame, text="Delete", command=delete_selected_batch, width=70, height=30,
              fg_color="#dc3545", hover_color="#c82333").pack(side="left", padx=2)

# Save to library button (in Manual Entry tab)
save_library_btn = ctk.CTkButton(library_frame, text="Save Current to Library",
                                 command=save_to_library, width=200, height=35,
                                 fg_color="transparent", border_width=2)
save_library_btn.pack(pady=15)

# Refresh categories on load
refresh_categories()

# ==================== Settings Tab ====================
settings_frame = ctk.CTkFrame(settings_tab, fg_color="transparent")
settings_frame.pack(fill="both", expand=True, padx=20, pady=20)

ctk.CTkLabel(settings_frame, text="Settings",
             font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 20))

# Appearance mode
appearance_frame = ctk.CTkFrame(settings_frame)
appearance_frame.pack(fill="x", pady=10, padx=20)
ctk.CTkLabel(appearance_frame, text="Appearance Mode:",
             font=ctk.CTkFont(size=13)).pack(side="left", padx=15, pady=15)
appearance_menu = ctk.CTkOptionMenu(appearance_frame, values=["Dark", "Light", "System"],
                                    command=change_appearance_mode, width=150)
appearance_menu.set("Dark")
appearance_menu.pack(side="right", padx=15, pady=15)

# Label type
label_frame = ctk.CTkFrame(settings_frame)
label_frame.pack(fill="x", pady=10, padx=20)
ctk.CTkLabel(label_frame, text="Label Type:",
             font=ctk.CTkFont(size=13)).pack(side="left", padx=15, pady=15)
label_menu = ctk.CTkOptionMenu(label_frame, variable=label_type_var,
                               values=list(label_types.keys()), width=150)
label_menu.pack(side="right", padx=15, pady=15)

# Show text checkbox
show_text_check = ctk.CTkCheckBox(settings_frame, text="Show text below barcode",
                                  variable=show_text_var, font=ctk.CTkFont(size=13))
show_text_check.pack(pady=15)

# Debug outlines checkbox
show_outlines_check = ctk.CTkCheckBox(settings_frame, text="Show label outlines (debug)",
                                      variable=show_label_outlines_var, font=ctk.CTkFont(size=13),
                                      text_color="gray")
show_outlines_check.pack(pady=(0, 15))

# Left margin
h_margin_frame = ctk.CTkFrame(settings_frame)
h_margin_frame.pack(fill="x", pady=10, padx=20)
ctk.CTkLabel(h_margin_frame, text="Left Margin (in):",
             font=ctk.CTkFont(size=13)).pack(side="left", padx=15, pady=15)
h_margin_entry = ctk.CTkEntry(h_margin_frame, textvariable=h_margin_var, width=150)
h_margin_entry.pack(side="right", padx=15, pady=15)

# Top margin
v_margin_frame = ctk.CTkFrame(settings_frame)
v_margin_frame.pack(fill="x", pady=10, padx=20)
ctk.CTkLabel(v_margin_frame, text="Top Margin (in):",
             font=ctk.CTkFont(size=13)).pack(side="left", padx=15, pady=15)
v_margin_entry = ctk.CTkEntry(v_margin_frame, textvariable=v_margin_var, width=150)
v_margin_entry.pack(side="right", padx=15, pady=15)

# Code text font size
font_size_frame = ctk.CTkFrame(settings_frame)
font_size_frame.pack(fill="x", pady=10, padx=20)
ctk.CTkLabel(font_size_frame, text="Code Text Size:",
             font=ctk.CTkFont(size=13)).pack(side="left", padx=15, pady=15)
font_size_menu = ctk.CTkOptionMenu(font_size_frame, variable=code_font_size_var,
                                   values=["6", "7", "8", "9", "10", "11", "12"],
                                   width=150)
font_size_menu.pack(side="right", padx=15, pady=15)

# === Status Bar ===
status_frame = ctk.CTkFrame(main_frame, height=60)
status_frame.pack(fill="x", pady=(15, 0))

progress_bar = ctk.CTkProgressBar(status_frame, width=400)
progress_bar.pack(pady=(10, 5))
progress_bar.set(0)

status_label = ctk.CTkLabel(status_frame, textvariable=status_var,
                            font=ctk.CTkFont(size=12))
status_label.pack(pady=(0, 10))

root.mainloop()
