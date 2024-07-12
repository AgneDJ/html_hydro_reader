import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# Reads percentages from HTML file


def read_html_table(html_file):
    with open(html_file, 'r') as file:
        soup = BeautifulSoup(file, 'html.parser')
    table = soup.find('table')
    data = []
    for row in table.find_all('tr'):
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        if cols:
            data.append(cols)
    return pd.DataFrame(data)

# Unmerges cells in the given worksheet


def unmerge_cells(ws):
    merged_cells = list(ws.merged_cells)
    for merged_cell in merged_cells:
        ws.unmerge_cells(str(merged_cell))

# Updates Excel file with percentages


def update_excel_with_percentages(excel_file, html_file, sheet_name, start_row, start_col):
    # Reads HTML table
    df_html = read_html_table(html_file)

    # Loads the workbook and worksheet
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    # Unmerges cells
    unmerge_cells(ws)

    # Updates Excel sheet with percentages
    for i, row in df_html.iterrows():
        for j, value in enumerate(row):
            ws.cell(row=start_row + i + 1,
                    column=start_col + j + 1, value=value)

    # Saves the workbook
    wb.save(excel_file)


html_file = '/Users/djacenko/Desktop/html_failai/StatProbability_Project1_Bartuva.html'
excel_file = '/Users/djacenko/Desktop/Pasitvirtinimo_vertinimas.xlsx'
sheet_name = 'Sheet3'
start_row = 30
start_col = 36

update_excel_with_percentages(
    excel_file, html_file, sheet_name, start_row, start_col)
