import pandas as pd
from bs4 import BeautifulSoup

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

# Updates excel file with percentage info


def update_excel_with_percentages(excel_file, html_file, sheet_name, start_row, start_col):
    # Reads HTML table
    df_html = read_html_table(html_file)

    # Reads Excel file
    excel_df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Updates Excel with percentages
    for i, row in df_html.iterrows():
        for j, value in enumerate(row):
            excel_df.iat[start_row + i, start_col + j] = value

    # Writes the update back to the Excel file
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        excel_df.to_excel(writer, sheet_name=sheet_name, index=False)


html_file = '/Users/djacenko/Desktop/html_failai/StatProbability_Project1_Bartuva.html'
excel_file = '/Users/djacenko/Desktop/Pasitvirtinimo_vertinimas.xlsx'
sheet_name = 'Sheet3'
start_row = 31
start_col = 26

update_excel_with_percentages(
    excel_file, html_file, sheet_name, start_row, start_col)
