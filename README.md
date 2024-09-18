# HTML HYDRO READER

Html hydro reader is excel updater is a desktop application built using Python and `tkinter` that allows users to update an Excel file with data extracted from HTML files. The application is designed with a graphical user interface (GUI) using Tkinter. This application is compatible with both macOS and Windows.

## Features

- Reads HTML tables from specified folders and extracts data.
- Writes extracted data into an Excel file's specified worksheet.
- Displays a progress bar and messages during the data update process.
- Includes a splash screen on startup.

## Prerequisites

- Windows operating system (the application has been tested on Windows 10/11).
- No need to install Python or any additional libraries; the application is bundled with all necessary dependencies.

## Installation

1. **Download the Application**

   - Download the `HydroReader.zip` file from the provided source.

2. **Extract the Files**

   - Extract the contents of the `HydroReader.zip` file to a folder on your system.

3. **Running the Application**
   - Navigate to the extracted folder and double-click the `HydroReader.exe` file to run the application.
   - Make sure all additional files (images, configuration files, etc.) are in the same directory as the `.exe`.

## Modules and Libraries Used

The application uses the following modules and libraries:

- `logging`: For logging events and errors during execution.
- `tkinter` and `ttk`: For creating the graphical user interface.
- `filedialog` and `messagebox`: For file selection and displaying messages.
- `os` and `glob`: For file system operations.
- `openpyxl`: For reading and writing Excel files.
- `BeautifulSoup` from `bs4`: For parsing HTML files and extracting data from tables.
- `pandas`: For data manipulation (used in some data processing functions).
- `PIL` (Pillow): For handling and displaying images.
- `threading`: For running the data update process in a separate thread to keep the GUI responsive.

## Usage

1. Double-click `HydroReader.exe` to start the app.
2. Select the base folder containing the HTML files.
3. Choose the month from the dropdown list.
4. Select the Excel file you want to update.
5. Choose the worksheet from the dropdown list.
6. Click "Run" to start the data update process.

## Notes

- Ensure all files extracted from the `HydroReader.zip` archive are kept in the same directory. Moving files individually may cause the application to malfunction.
- If the application uses external files like configuration files, images, or databases, they should be included in the same directory as the `.exe` file.

## Troubleshooting

- If the application does not start, ensure your antivirus software is not blocking it. Sometimes, antivirus programs may flag `.exe` files as potential threats.
- If any DLL errors appear, make sure all necessary files from the `dist` directory are present in the same folder as the `.exe`.
- Check the `excel_updater.log` file for detailed logs and error messages if something goes wrong.

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.

## Acknowledgments

- [PyInstaller](https://www.pyinstaller.org/) - For packaging the Python script into an executable.
- [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) - For parsing HTML files.
- [OpenPyXL](https://openpyxl.readthedocs.io/) - For working with Excel files.
- [Pillow](https://python-pillow.org/) - For image handling in Tkinter.
