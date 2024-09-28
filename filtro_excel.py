import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from version import VERSION
import pandas as pd
import os
import sys
import unidecode

def resource_path(path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, path)
    return path

def show_help():
    help_window = tk.Toplevel(root)
    help_window.title("Help")
    help_window.geometry("600x400")

    help_text = tk.Text(help_window, wrap=tk.WORD)
    help_text.pack(expand=True, fill="both", padx=10, pady=10)

    scroll_y = tk.Scrollbar(help_text, orient="vertical", command=help_text.yview)
    scroll_y.pack(side="right", fill="y")
    help_text.configure(yscrollcommand=scroll_y.set)

    try:
        with open(resource_path("readme.txt"), "r", encoding="utf-8") as file:
            content = file.read()
            help_text.insert(tk.END, content)
    except FileNotFoundError:
        help_text.insert(tk.END, "Help file is not available.")

    help_text.config(state="disabled")

    close_button = tk.Button(help_window, text="Close", command=help_window.destroy)
    close_button.pack(pady=10)

root = tk.Tk()
root.title("Excel File Filter")
root.geometry("900x600")

df = None
entries = {}
labels = {}
filtered_indices = []
elements_var = tk.StringVar()
current_file = None

def load_excel():
    global df, current_file
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file:
        try:
            current_file = file
            # Especificar que la primera fila es el encabezado
            df = pd.read_excel(file, header=0)
            # Eliminar filas que son todas NaN
            df = df.dropna(how='all')
            # Eliminar columnas que son todas NaN
            df = df.dropna(axis=1, how='all')
            # Convertir todos los valores a string, reemplazando NaN con "---" y eliminando ".0" en números
            df = df.applymap(lambda x: str(x).rstrip('.0') if pd.notna(x) else "---")
            initialize_interface()
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file: {str(e)}")

def remove_accents(text):
    return unidecode.unidecode(text)

def filter_elements(*args):
    global filtered_indices
    filter_condition = pd.Series([True] * len(df))
    for column, entry in entries.items():
        value = entry.get()
        if value and value != "---":
            filter_condition &= df[column].astype(str).str.lower() == value.lower()
    filtered_indices = df[filter_condition].index.tolist()
    
    for column, entry in entries.items():
        filtered_options = ["---"] + sorted(df.loc[filtered_indices, column].astype(str).unique().tolist())
        combobox = entry.widget
        current_value = combobox.get()
        combobox['values'] = filtered_options
        if current_value not in filtered_options:
            combobox.set("---")

    update_results()

def update_results():
    if len(filtered_indices) == 1:
        show_record(0)
        verify_editable_fields()
    elif len(filtered_indices) == 0:
        record_text.config(state='normal')
        record_text.delete(1.0, tk.END)
        record_text.insert(tk.END, "NO RESULTS")
        record_text.config(state='disabled')
    else:
        record_text.config(state='normal')
        record_text.delete(1.0, tk.END)
        record_text.insert(tk.END, f"{len(filtered_indices)} results found")
        record_text.config(state='disabled')

def show_record(element_id):
    real_index = filtered_indices[element_id]
    element_data = df.loc[real_index]

    record_text.config(state='normal')
    record_text.delete(1.0, tk.END)

    for column in df.columns:
        record_text.insert(tk.END, f"{column}: ", "label")
        value = str(element_data[column])
        tag_name = f"editable_{column.replace(' ', '_').replace('.', '_')}"
        record_text.insert(tk.END, f"{value}\n", tag_name)
        record_text.tag_configure(tag_name, background="light yellow", selectbackground="blue", selectforeground="white")
        record_text.tag_bind(tag_name, "<Double-Button-1>", 
                            lambda e, col=column: edit_field(e, real_index, col))

    record_text.config(state='disabled')

def edit_field(event, index, column):
    print(f"Attempting to edit field: {column}, index: {index}")
    tag_name = f"editable_{column.replace(' ', '_').replace('.', '_')}"
    current_text = record_text.get(f"{tag_name}.first", f"{tag_name}.last").strip()
    print(f"Current text: '{current_text}'")
    
    x, y = event.x_root, event.y_root
    
    edit_window = tk.Toplevel(root)
    edit_window.title(f"Edit {column}")
    edit_window.geometry(f"+{x}+{y}")
    
    entry = tk.Entry(edit_window, width=50)
    entry.insert(0, current_text)
    entry.pack(padx=10, pady=10)
    entry.focus_set()
    
    def save_change():
        new_value = entry.get()
        print(f"Saving new value: '{new_value}' for {column}")
        df.at[index, column] = new_value
        edit_window.destroy()
        filter_elements()
        show_record(filtered_indices.index(index))
    
    save_button = tk.Button(edit_window, text="Save", command=save_change)
    save_button.pack(pady=10)
    
    entry.bind("<Return>", lambda e: save_change())

def verify_editable_fields():
    print("Columns in the DataFrame:")
    for col in df.columns:
        print(f"  - {col}")
    
    print("\nEditable fields:")
    for tag in record_text.tag_names():
        if tag.startswith("editable_"):
            print(f"Editable field: {tag}")
            ranges = record_text.tag_ranges(tag)
            for i in range(0, len(ranges), 2):
                start, end = ranges[i], ranges[i+1]
                print(f"  Range: {start} - {end}")
                print(f"  Content: {record_text.get(start, end).strip()}")

def clear_fields():
    for entry in entries.values():
        entry.set('---')
    global filtered_indices
    filtered_indices = df.index.tolist()
    filter_elements()

load_frame = ttk.Frame(root, padding="10")
load_frame.grid(row=0, column=0, columnspan=3, sticky="w")

load_button = tk.Button(load_frame, text="Load Excel File", command=load_excel, bg="blue", fg="white")
load_button.pack(side="left", padx=10, pady=10)

clear_button = tk.Button(load_frame, text="Clear Fields", command=clear_fields, bg="green", fg="white")
clear_button.pack(side="left", padx=10, pady=10)

help_button = tk.Button(load_frame, text="Help", command=show_help, bg="orange", fg="white")
help_button.pack(side="left", padx=10, pady=10)

# boton_debug = tk.Button(cargar_frame, text="Debug", command=verificar_campos_editables, bg="red", fg="white")
# boton_debug.pack(side="left", padx=10, pady=10)

def add_logo_and_version():
    logo_path = resource_path("Logo.png")
    logo = Image.open(logo_path)
    logo_img = ImageTk.PhotoImage(logo)

    info_frame = ttk.Frame(root)
    info_frame.grid(row=0, column=2, columnspan=2, sticky="e", padx=10, pady=10)

    version_label = ttk.Label(info_frame, text=VERSION, font=("Arial", 7))
    version_label.pack(side="left", padx=(0, 5))

    text_label = ttk.Label(info_frame, text="by apedrajas for", font=("Arial", 7))
    text_label.pack(side="left", padx=(0, 5))

    logo_label = ttk.Label(info_frame, image=logo_img)
    logo_label.image = logo_img
    logo_label.pack(side="left")

add_logo_and_version()

main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=1, column=0, columnspan=4, sticky="nsew")

canvas = tk.Canvas(main_frame)
scroll_x = tk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
scroll_y = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)

canvas.configure(xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)

scroll_x.pack(side="bottom", fill="x")
scroll_y.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

content_frame = ttk.Frame(canvas)
canvas.create_window((0, 0), window=content_frame, anchor="nw")

search_frame = ttk.Frame(content_frame)
search_frame.grid(row=1, column=0, sticky="nsew")

def initialize_interface():
    global filtered_indices
    filtered_indices = df.index.tolist()

    for widget in search_frame.winfo_children():
        widget.destroy()

    for i, column in enumerate(df.columns):
        label = ttk.Label(search_frame, text=column + ":", anchor="e")
        label.grid(row=i, column=0, padx=5, pady=5, sticky="E")
        labels[column] = label

        options = ["---"] + sorted(df[column].unique().tolist())
        entry = tk.StringVar()
        combobox = ttk.Combobox(search_frame, textvariable=entry, values=options)
        combobox.set("---")
        combobox.grid(row=i, column=1, padx=5, pady=5)
        combobox.bind("<<ComboboxSelected>>", filter_elements)
        combobox.bind("<KeyRelease>", filter_elements)
        entries[column] = entry
        entry.widget = combobox

    update_results()

record_frame = ttk.Frame(content_frame, padding="10")
record_frame.grid(row=1, column=1, sticky="nsew")

record_text = tk.Text(record_frame, height=20, width=60, wrap="none")
record_text.grid(row=0, column=1, sticky="nsew")
record_text.config(state='disabled')

scroll_x_record = tk.Scrollbar(record_frame, orient="horizontal", command=record_text.xview)
scroll_x_record.grid(row=1, column=1, sticky="ew")
record_text.config(xscrollcommand=scroll_x_record.set)

scroll_y_record = tk.Scrollbar(record_frame, orient="vertical", command=record_text.yview)
scroll_y_record.grid(row=0, column=2, sticky="ns")
record_text.config(yscrollcommand=scroll_y_record.set)

def export_results(filtered_only=False):
    if df is None:
        messagebox.showwarning("Warning", "No data to export.")
        return
    
    save_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_file:
        if filtered_only:
            df_export = df.loc[filtered_indices]
        else:
            df_export = df
        
        df_export.to_excel(save_file, index=False)
        messagebox.showinfo("Success", "File exported successfully.")

export_all_button = tk.Button(root, text="Export Complete File", 
                              command=lambda: export_results(False), 
                              bg="gray", fg="white")
export_all_button.grid(row=2, column=1, sticky="e", padx=10, pady=10)

export_filtered_button = tk.Button(root, text="Export Filtered Results", 
                                   command=lambda: export_results(True), 
                                   bg="light gray", fg="black")
export_filtered_button.grid(row=2, column=2, sticky="w", padx=10, pady=10)

def on_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

content_frame.bind("<Configure>", on_configure)

root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=1)

root.mainloop()