import cv2
import easyocr
import numpy as np
from PIL import ImageGrab, Image, ImageTk
import tkinter as tk
from tkinter import ttk
from collections import defaultdict

max_height = 500

threshold = 0.25

images_list = []  # store PhotoImage references to prevent GC
row_counter = 1   # to number the rows

last_clicked_cell = None  # (row_id, col_index)


# FUNCTIONS
# --- OCR and drawing function ---
def process_clipboard_image(event=None):  # event=None allows key binding
    # Grab image from clipboard
    global images_list, row_counter
    img_pil = ImageGrab.grabclipboard()
    if img_pil is None:
        result_label.config(text="No image in clipboard!")
        return

    # Convert PIL image to OpenCV BGR
    img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)

    # Run EasyOCR
    reader = easyocr.Reader(['en'], gpu=False)
    text_ = reader.readtext(img)
    

    # Draw rectangles and prepare for line grouping
    y_threshold = 10  # pixels
    lines = defaultdict(list)
    
    for bbox, text, score in text_:
        if score < threshold:
            continue

        # Draw rectangle
        pt1 = tuple(map(int, bbox[0]))
        pt2 = tuple(map(int, bbox[2]))
        cv2.rectangle(img, pt1, pt2, (0, 255, 0), 2)
        cv2.putText(img, text, pt1, cv2.FONT_HERSHEY_COMPLEX, 0.65, (255, 0, 0), 2)

        # Calculate y_center for line grouping
        y_top = bbox[0][1]
        y_bottom = bbox[3][1]
        y_center = (y_top + y_bottom) / 2

        # Group text by line
        found_line = False
        for key in lines:
            if abs(key - y_center) <= y_threshold:
                lines[key].append((bbox, text))
                found_line = True
                break
        if not found_line:
            lines[y_center].append((bbox, text))

    # Prepare table data
    table_data = []
    for y in sorted(lines.keys()):
        line_text = sorted(lines[y], key=lambda x: x[0][0][0])  # sort left to right
        line_str = " ".join([w[1] for w in line_text])
        #table_data.append(line_str)

         # Convert to excel format
        if "pts" in line_str:
            parts = line_str.rsplit(' ', 2)
            name = parts[0]
            points = parts[1].replace(',', '')
            table_data.append([name, points])

    # Clear previous table
    # for row in tree.get_children():
    #     tree.delete(row)
    # Insert new rows
    for name, points in table_data:
        tree.insert("", "end", values=(row_counter, name, points))
        row_counter += 1
    
    # --- Scale image after processing ---
    max_height = 600
    h, w = img.shape[:2]
    if h > max_height:
        scale_ratio = max_height / h
        new_w = int(w * scale_ratio)
        new_h = max_height
        img = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_AREA)

    # Convert OpenCV BGR image to PIL Image for Tkinter
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    img_tk = ImageTk.PhotoImage(Image.fromarray(img_rgb))

    # Display in Tkinter
    images_list.append(img_tk)
    lbl = tk.Label(image_frame, image=img_tk)
    lbl.pack(pady=5)
    # image_label.config(image=img_tk)
    # image_label.image = img_tk
    result_label.config(text=f"OCR complete: {len(images_list)} image(s) uploaded.")

def on_tree_click(event):
    global last_clicked_cell
    row_id = tree.identify_row(event.y)
    column_id = tree.identify_column(event.x)
    if not row_id or not column_id:
        last_clicked_cell = None
        return
    col_index = int(column_id.replace("#", "")) - 1
    last_clicked_cell = (row_id, col_index)

# --- Copy selected cell function ---
def copy_selected_cell(event=None):
    """Copy value from the last clicked cell"""
    if not last_clicked_cell:
        return
    row_id, col_index = last_clicked_cell
    values = tree.item(row_id, "values")
    if col_index < len(values):
        root.clipboard_clear()
        root.clipboard_append(values[col_index])
        root.update()

# --- Sort columns
def treeview_sort_column(tv, col, reverse=False):
    """
    Sorts the Treeview by the given column.
    tv: the Treeview widget
    col: the column index or column id
    reverse: sort descending if True
    """
    # Get all items
    data_list = [(tv.set(k, col), k) for k in tv.get_children('')]
    
    # Try to convert to int for numeric sort
    try:
        data_list.sort(key=lambda t: int(t[0]), reverse=reverse)
    except ValueError:
        # Sort as string if not int
        data_list.sort(key=lambda t: t[0], reverse=reverse)
    
    # Reorder items in Treeview
    for index, (val, k) in enumerate(data_list):
        tv.move(k, '', index)
    
    # Reverse next time
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

def copy_all():
    """Copy all rows as 'Name,Points'"""
    rows = tree.get_children()
    all_text = []
    for row in rows:
        values = tree.item(row, "values")
        if len(values) >= 3:
            all_text.append(f"{values[1]},{values[2]}")
    if all_text:
        root.clipboard_clear()
        root.clipboard_append("\n".join(all_text))
        root.update()

def copy_points():
    """Copy Points from the selected row"""
    """Copy Points column from all rows"""
    rows = tree.get_children()
    points_list = []
    for row in rows:
        values = tree.item(row, "values")
        if len(values) >= 3:  # Ensure Points column exists
            points_list.append(values[2])  # Points is column index 2
    if points_list:
        root.clipboard_clear()
        root.clipboard_append("\n".join(points_list))
        root.update()

def clear_all():
    global row_counter, images_list

    # Clear the table
    for row in tree.get_children():
        tree.delete(row)
    
    # Clear images
    for widget in image_frame.winfo_children():
        widget.destroy()
    
    # Clear image references
    images_list = []
    
    # Reset row counter
    row_counter = 1

    result_label.config(text="Image(s) cleared.")

# GUI
root = tk.Tk()
root.title("Uma OCR")
root.geometry("800x500")

# Main frame
main_frame = tk.Frame(root)
main_frame.pack(padx=10, pady=10)

# Left frame for images with scrollbar

# image_label = tk.Label(left_frame)
# image_label.pack()
# result_label = tk.Label(left_frame, text="")
# result_label.pack(pady=5)

# Left frame for images with scrollbar
left_outer_frame = tk.Frame(main_frame)
left_outer_frame.pack(side="left", padx=10, fill="both", expand=True)

left_button_frame = tk.Frame(left_outer_frame)
left_button_frame.pack(pady=5, fill="x")

upload_btn = tk.Button(left_button_frame, text="Ctrl-V or Click to Paste from Clipboard", command=process_clipboard_image).pack(side="left", padx=5)
clear_btn = tk.Button(left_button_frame, text="Clear", command=clear_all).pack(side="left")

result_label = tk.Label(left_outer_frame, text="")
result_label.pack(pady=5)

image_canvas = tk.Canvas(left_outer_frame)
scrollbar = tk.Scrollbar(left_outer_frame, orient="vertical", command=image_canvas.yview)
image_canvas.configure(yscrollcommand=scrollbar.set)

scrollbar.pack(side="right", fill="y")
image_canvas.pack(side="left", fill="both", expand=True)

image_frame = tk.Frame(image_canvas)
image_canvas.create_window((0, 0), window=image_frame, anchor="nw")

def _on_mousewheel(event):
    image_canvas.yview_scroll(-int(event.delta/120), "units")

# Update scrollregion when new images are added
def update_scrollregion(event=None):
    image_canvas.configure(scrollregion=image_canvas.bbox("all"))


image_frame.bind("<Configure>", update_scrollregion)
image_canvas.bind_all("<MouseWheel>", _on_mousewheel)  # Windows

# Right: Table + Copy button
right_frame = tk.Frame(main_frame)
right_frame.pack(side="left", padx=10)

button_frame = tk.Frame(right_frame)
button_frame.pack(pady=5, fill="x")

copy_points_btn = tk.Button(button_frame, text="Copy Points", command=lambda: copy_points())
copy_points_btn.pack(side="right", padx=5)

copy_all_btn = tk.Button(button_frame, text="Copy All", command=lambda: copy_all())
copy_all_btn.pack(side="right", padx=5)

tree = ttk.Treeview(right_frame, columns=("No", "Name", "Points"), show="headings", height=20)
tree.heading("No", text="#")
tree.heading("Name", text="Name")
tree.heading("Points", text="Points")
tree.column("No", width=30, anchor="center")
tree.column("Name", width=150, anchor="w")
tree.column("Points", width=150, anchor="center")
tree.pack()

# Bind Ctrl+C for copying selected rows
tree.bind("<Control-c>", copy_selected_cell)
# Bind click to track last clicked cell
tree.bind("<Button-1>", on_tree_click)
# Bind Ctrl+V to process clipboard image
root.bind("<Control-v>", process_clipboard_image)

# E
for col in ("No", "Name", "Points"):
    tree.heading(col, text=col, command=lambda c=col: treeview_sort_column(tree, c, False))


root.mainloop()