# HTML HYDRO READER

Html hydro reader is excel updater is a desktop application built using Python and `tkinter` that allows users to update an Excel file with data extracted from HTML files. This application is compatible with both macOS and Windows.

## Features

- Select a base folder containing HTML files.
- Select an Excel file to update.
- Specify the sheet name in the Excel file.
- Preview the updated Excel file within the application.

## Prerequisites

Make sure you have Python installed. The application requires the following Python packages:

- `pandas`
- `beautifulsoup4`
- `openpyxl`
- `tkinter` (usually included with Python but needs to be installed separately on some systems like macOS)

## Installation

1.  **Clone the repository or download the script:**

    ```bash
    git clone https://github.com/yourusername/html_hydro_reader.git
    cd html_hydro_reader

    ```

2.  **Install the required Python packages:**

3.  **Install tkinter (if not already installed):**

        On macOS:
            brew install python-tk@3.11

    - Ensure you have the necessary environment variables set in your shell configuration file (~/.bash_profile, ~/.zshrc, or other):

      export PATH="/usr/local/opt/tcl-tk/bin:$PATH"
      export LDFLAGS="-L/usr/local/opt/tcl-tk/lib"
      export CPPFLAGS="-I/usr/local/opt/tcl-tk/include"
      export PKG_CONFIG_PATH="/usr/local/opt/tcl-tk/lib/pkgconfig"

    - Reload your shell configuration:
      source ~/.bash_profile # or `source ~/.zshrc`

      On Windows:
      tkinter is included with Python. No additional steps are required.

## Usage:

    1. **Run the application:**

        On macOS:
            python3.11 hydro_app.py

        On Windows:
            python hydro_app.py

    2. **Use the application:**

        - Click the "Browse" button next to "Base Folder" to select the folder containing your HTML files.
        - Click the "Browse" button next to "Excel File" to select the Excel file you want to update.
        - Enter the sheet name in the "Sheet Name" field.
        - Click the "Run" button to update the Excel file and preview the updated content within the application.

## Code Explanation

    - read_html_table: Reads percentages from an HTML file and returns the data as a list of tuples.
    - unmerge_cells: Unmerges any merged cells in the specified worksheet.
    - get_column_index: Gets the column index based on the folder name.
    - update_excel_with_percentages: Updates the Excel file with percentages from the HTML files.
    - select_base_folder: Opens a file dialog to select the base folder.
    - select_excel_file: Opens a file dialog to select the Excel file.
    - run_update: Runs the update process and previews the updated Excel file.
    - preview_excel_file: Previews the updated Excel file in a Treeview widget.

## Contributing

    Contributions are welcome! Please open an issue or submit a pull request for any changes.
