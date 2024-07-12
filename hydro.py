import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import glob
import os

# Reads percentages from HTML file


def read_html_table(html_file):
    with open(html_file, 'r') as file:
        soup = BeautifulSoup(file, 'html.parser')
    table = soup.find('table', {'id': 'table1'})
    if table:
        print(f"Table found in HTML file: {html_file}")
    else:
        print(f"No table found in HTML file: {html_file}")
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

# Updates Excel file with percentages


def update_excel_with_percentages(excel_file, html_folder, sheet_name, start_col):
    # Get a list of all HTML files in the folder
    html_files = glob.glob(os.path.join(html_folder, '*.html'))
    print(f"Found HTML files: {html_files}")

    # Loads the workbook and worksheet
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    # Unmerges cells
    unmerge_cells(ws)

    for html_file in html_files:
        # Reads HTML table
        data = read_html_table(html_file)
        if not data:
            continue  # Skip if no data

        # Updates Excel sheet with percentages
        for row_name, values in data:
            # Column C is the 3rd column
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=3):
                if row[0].value == row_name:
                    row_idx = row[0].row
                    for j, value in enumerate(values):
                        if value is not None:  # Only write non-None values
                            ws.cell(row=row_idx, column=start_col +
                                    j, value=value)
                            print(
                                f"Writing value {value} to cell ({row_idx}, {start_col + j})")

    # Saves the workbook
    wb.save(excel_file)
    print(f"Workbook {excel_file} saved successfully.")


# Folder containing HTML files
html_folder = '/Users/djacenko/Desktop/html_failai/'

excel_file = '/Users/djacenko/Desktop/Pasitvirtinimo_vertinimas.xlsx'
sheet_name = 'Sheet1'
start_col = 37  # Column 27 corresponds to column 'AA' in Excel

update_excel_with_percentages(excel_file, html_folder, sheet_name, start_col)
