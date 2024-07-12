import pandas as pd
from bs4 import BeautifulSoup

# Function to read percentages from HTML file


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

# Function to update Excel file with percentages


def update_excel_with_percentages(excel_file, html_file, sheet_name, start_row, start_col):
    # Read HTML table
    df_html = read_html_table(html_file)

    # Read Excel file
    excel_df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Update Excel DataFrame with percentages
    for i, row in df_html.iterrows():
        for j, value in enumerate(row):
            excel_df.iat[start_row + i, start_col + j] = value

    # Write the updated DataFrame back to the Excel file
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        excel_df.to_excel(writer, sheet_name=sheet_name, index=False)


# Example usage
html_file = 'path_to_html_file.html'
excel_file = 'path_to_excel_file.xlsx'
sheet_name = 'Sheet1'  # Adjust as needed
start_row = 2  # Adjust as needed
start_col = 3  # Adjust as needed

update_excel_with_percentages(
    excel_file, html_file, sheet_name, start_row, start_col)
