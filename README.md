# PDF Data Extractor

## Overview
This Python program extracts structured data from PDF files using predefined templates. It processes multiple PDFs in a specified folder and extracts tables using Tabula, then saves the extracted data to an Excel file.

## Features
- Extracts tabular data from PDFs using predefined template coordinates.
- Supports batch processing of multiple PDFs in a selected folder.
- Saves extracted data to an Excel file for better readability.
- Uses a graphical file dialog for selecting input and output locations.

## Requirements
Before running the program, ensure you have the following dependencies installed:

- Python 3.x
- `pandas`
- `tabula-py`
- `openpyxl`

You can install the required packages using the following command:
```sh
pip install pandas tabula-py openpyxl
```

Additionally, `Java` must be installed and configured for `tabula-py` to work correctly.

## Installation & Usage

### 1. Clone or Download the Repository
```sh
git clone https://github.com/your-repository/pdf-data-extractor.git
cd pdf-data-extractor
```

### 2. Run the Program
Run the script using Python:
```sh
python main.py
```

### 3. Using the File Dialogs
1. **Select Folder:** Choose the folder containing the PDF files to process.
2. **Select Template File:** Choose a JSON file containing predefined extraction coordinates.
3. **Select Output File:** Specify an Excel file where extracted data will be saved.
4. The program will process the PDFs and save all extracted tables to the selected Excel file.

## Template File Format
The template file should be a JSON file containing coordinate areas for data extraction, as required by Tabula. Example format:
```json
[
    {"x1": 100, "y1": 200, "x2": 400, "y2": 250},
    {"x1": 150, "y1": 300, "x2": 450, "y2": 350}
]
```
These coordinates define areas where specific data fields (e.g., tables) are located in the PDF.

## Output Format
The extracted data is saved in Excel format with:
- **Sheet Name:** `All_Tables`
- **Columns:** Extracted table data from PDFs

## Troubleshooting
- If `tabula-py` fails to read PDFs, ensure Java is installed and added to the system PATH.
- If no data is extracted, check if the template coordinates match the PDF layout.

## License
This project is open-source.

## Author
Ilavarasan

