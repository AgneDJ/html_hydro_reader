import logging
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import os
import glob
from openpyxl import load_workbook
from bs4 import BeautifulSoup
import pandas as pd
To add a welcome page or splash screen that displays when the application starts, we can create a small window that shows the app name and author. This window will automatically close after a specified duration, such as 2 seconds. Here’s how to implement it:

1. ** Create a Splash Screen**: Create a small window that displays the welcome message.
2. ** Auto-Close the Splash Screen**: Set a timer to close the splash screen after 2 seconds and then show the main application window.

Here’s the code with the splash screen added:

```python

# Set up logging
logging.basicConfig(filename='excel_updater.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Reads percentages from HTML file


def read_html_table(html_file):
    try:
        with open(html_file, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, 'html.parser')
    except Exception as e:
        logging.error(f"Error reading HTML file {html_file}: {e}")
        return []

    table = soup.find('table', {'id': 'table1'})
    if table:
        logging.info(f"Table found in HTML file: {html_file}")
    else:
        logging.warning(f"No table found in HTML file: {html_file}")
        return []  # Return empty list if no table is found

    data = []
    for row in table.find_all('tr')[1:]:  # Skip the header row by slicing
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        if cols:
            row_name = cols[0]  # First value is the row name
            # Remove percentage symbols and convert to float, skip the first column
            row_data = []
            for ele in cols[1:]:
                try:
                    value = float(ele.replace('%', ''))
                except ValueError:
                    value = None  # Set to None if conversion fails
                row_data.append(value)
            data.append((row_name, row_data))
    return data

# Unmerges cells in the given worksheet


def unmerge_cells(ws):
    merged_cells = list(ws.merged_cells)
    for merged_cell in merged_cells:
        ws.unmerge_cells(str(merged_cell))

# Get the column index based on the folder name


def get_column_index(ws, folder_name):
    for col in ws.iter_cols(min_row=7, max_row=7, min_col=4, max_col=ws.max_column):
        column_value = str(col[0].value)  # Convert column value to string
        logging.info(
            f"Checking column {col[0].column} with value {column_value} against folder {folder_name}")
        if column_value == folder_name:
            return col[0].column
    return None

# Collects data from all HTML files in the month folder


def collect_data(base_folder, month_folder):
    collected_data = {}
    folder_path = os.path.join(base_folder, month_folder, '*')
    folder_list = [f for f in glob.glob(folder_path) if os.path.isdir(f)]

    logging.info(f"Found day folders: {folder_list}")

    for folder in folder_list:
        folder_name = os.path.basename(folder)
        html_files = glob.glob(os.path.join(folder, '*.html'))

        folder_data = []
        for html_file in html_files:
            # Reads HTML table
            data = read_html_table(html_file)
            if data:
                folder_data.append((folder_name, data))
        collected_data[folder_name] = folder_data

    return collected_data

# Writes collected data to Excel file


def write_data_to_excel(excel_file, sheet_name, collected_data):
    try:
        # Load the workbook and worksheet
        wb = load_workbook(excel_file)
        logging.info("Available sheet names: " + ", ".join(wb.sheetnames))
        if sheet_name not in wb.sheetnames:
            logging.error(f"Worksheet {sheet_name} does not exist.")
            messagebox.showwarning(
                "Warning", f"Worksheet '{sheet_name}' does not exist.")
            return

        ws = wb[sheet_name]
        unmerge_cells(ws)

        for folder_name, data_entries in collected_data.items():
            column_index = get_column_index(ws, folder_name)
            if column_index is None:
                logging.warning(f"No matching column for folder {folder_name}")
                continue

            for _, data in data_entries:
                for row_name, values in data:
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=3):
                        if row[0].value == row_name:
                            row_idx = row[0].row
                            for j, value in enumerate(values):
                                if value is not None:  # Only write non-None values
                                    ws.cell(
                                        row=row_idx, column=column_index + j, value=value)
                                    logging.info(
                                        f"Writing value {value} to cell ({row_idx}, {column_index + j})")

        # Save the workbook
        wb.save(excel_file)
        logging.info(f"Workbook {excel_file} saved successfully.")
        messagebox.showinfo("Success", "Data added successfully!")
    except Exception as e:
        logging.error(f"An error occurred while updating the Excel file: {e}")
        messagebox.showwarning("Warning", f"An error occurred: {e}")

# Function to select the base folder and populate month dropdown


def select_base_folder():
    folder_selected = filedialog.askdirectory()
    base_folder_entry.delete(0, tk.END)
    base_folder_entry.insert(0, folder_selected)

    # Populate month dropdown with subfolder names
    if folder_selected:
        subfolders = [name for name in os.listdir(
            folder_selected) if os.path.isdir(os.path.join(folder_selected, name))]
        month_combo['values'] = subfolders
        month_combo.set('')  # Clear the current selection

# Function to select the Excel file and populate sheet dropdown


def select_excel_file():
    file_selected = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")])
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(0, file_selected)

    # Populate sheet dropdown with sheet names
    if file_selected:
        wb = load_workbook(file_selected, read_only=True)
        sheets = wb.sheetnames
        sheet_name_combo['values'] = sheets
        sheet_name_combo.set('')  # Clear the current selection

# Function to run the update process


def run_update():
    base_folder = base_folder_entry.get()
    month_folder = month_combo.get()
    excel_file = excel_file_entry.get()
    sheet_name = sheet_name_combo.get()

    if not base_folder or not excel_file or not sheet_name or not month_folder:
        messagebox.showerror("Error", "Please provide all inputs.")
        return

    # Collect data from all HTML files first
    collected_data = collect_data(base_folder, month_folder)

    # Write collected data to the Excel file
    write_data_to_excel(excel_file, sheet_name, collected_data)

# Function to show the splash screen


def show_splash():
    splash = tk.Toplevel()
    splash.title("Welcome")
    splash.geometry("300x150")
    splash_label = tk.Label(
        splash, text="Excel Updater App\nAuthor: Your Name", font=("Helvetica", 14))
    splash_label.pack(expand=True)

    # Hide splash after 2 seconds and show the main window
    root.after(2000, splash.destroy)

# Create the main GUI


def create_main_window():
    global base_folder_entry, month_combo, excel_file_entry, sheet_name_combo, root

    root = tk.Tk()
    root.title("Excel Updater")

    # Base folder selection
    tk.Label(root, text="Base Folder:").grid(row=0, column=0, padx=10, pady=10)
    base_folder_entry = tk.Entry(root, width=50)
    base_folder_entry.grid(row=0, column=1, padx=10, pady=10)
    tk.Button(root, text="Browse", command=select_base_folder).grid(
        row=0, column=2, padx=10, pady=10)

    # Month selection dropdown
    tk.Label(root, text="Select Month:").grid(
        row=1, column=0, padx=10, pady=10)
    month_combo = ttk.Combobox(root)
    month_combo.grid(row=1, column=1, padx=10, pady=10)

    # Excel file selection
    tk.Label(root, text="Excel File:").grid(row=2, column=0, padx=10, pady=10)
    excel_file_entry = tk.Entry(root, width=50)
    excel_file_entry.grid(row=2, column=1, padx=10, pady=10)
    tk.Button(root, text="Browse", command=select_excel_file).grid(
        row=2, column=2, padx=10, pady=10)

    # Sheet name dropdown
    tk.Label(root, text="Sheet Name:").grid(row=3, column=0, padx=10, pady=10)
    sheet_name_combo = ttk.Combobox(root)
    sheet_name_combo.grid(row=3, column=1, padx=10, pady=10)

    # Run button
    tk.Button(root, text="Run", command=run_update).grid(
        row=4, column=1, pady=20)

# Show
