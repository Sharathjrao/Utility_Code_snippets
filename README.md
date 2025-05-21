# Utility_Code_snippets

# Multi-Utility Tkinter Application

A comprehensive desktop application built with Python and Tkinter that provides a suite of file and folder utilities, PDF processing tools, and CSV management features — all accessible through a clean tabbed interface.

---

## Features

### Folder Utilities
- **Remove Empty Subfolders:** Recursively remove empty subfolders (excluding the root folder).
- **Remove Empty Folders:** Recursively remove empty folders including the root folder if empty.
- **Browse Folder Contents:** View files and folders in a selected directory.

### CSV Utilities
- **Save Subfolder Paths to CSV:** Export full paths of all subfolders under a selected folder.
- **Save Subfolder Names to CSV:** Export only the names of subfolders.
- **Create Folders From CSV:** Create folders based on names listed in a CSV file.

### PDF Utilities
- **Split PDF by Pages:** Split a PDF into multiple files with user-defined pages per split; output filenames are based on header text extracted from the PDF.
- **Merge PDFs in Pairs:** Merge PDFs two at a time from a folder; output filenames are based on header text from the first page of the first PDF in each pair.

### File Search & Copy
- **Search Files by Extension & Save CSV:** Recursively search for files with a specified extension and save their paths to a CSV file.
- **Filter CSV & Copy DJI Images:** Filter a CSV for rows where filenames start with 'DJI' and end with 'jpeg' and copy those images to an output folder.

### CSV Viewer
- **Load and View CSV Files:** Load any CSV file and display its contents in a table within the app.

---

## Installation

### Prerequisites
- Python 3.8 or higher
- `pip` package manager

### Dependencies
Install required Python packages:

pip install PyPDF2

text

---

## Usage

1. Clone or download this repository.
2. Run the application:

python app.py

text

3. Use the tabbed interface to select the utility you want.
4. Follow on-screen prompts to select folders, files, or enter parameters.
5. Long-running operations run in the background to keep the UI responsive.
6. Status messages and results will be displayed within the app.

---

## Project Structure

app.py # Main application script with all utilities
README.md # This file

text

---

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to fork the repo and submit pull requests.

---

## License

This project is licensed under the MIT License.

---

## Contact

For questions or support, please open an issue or contact [Your Name](mailto:your.email@example.com).

---

## Acknowledgments

- [PyPDF2](https://pypi.org/project/PyPDF2/) for PDF processing.
- Python's standard libraries for file and CSV management.
- Tkinter for the GUI framework.
