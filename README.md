# PDF Data Extractor

## Overview
This Python program extracts structured data from PDF files using predefined templates. It processes multiple PDFs in a specified folder and extracts values such as GST numbers and monetary amounts based on the given template. The extracted data is then saved as a CSV file.

## Features
- Extracts tabular data from PDFs using predefined template coordinates.
- Supports batch processing of multiple PDFs in a selected folder.
- Provides a graphical user interface (GUI) for easy file selection and execution.
- Saves extracted data to a CSV file for further analysis.

## Requirements
Before running the program, ensure you have the following dependencies installed:

- Python 3.x
- `tkinter` (built-in with Python)
- `pandas`
- `tabula-py`
- `PyPDF2`

You can install the required packages using the following command:
```sh
pip install pandas tabula-py PyPDF2
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

### 3. Using the GUI
1. **Select Source Folder:** Choose the folder containing the PDF files to process.
2. **Select Template File:** Choose a JSON file containing predefined extraction coordinates.
3. **Select Output File:** Specify a CSV file where extracted data will be saved.
4. **Click "Start Processing"** to extract data and save it.

## Template File Format
The template file should be a JSON file containing coordinate areas for data extraction. Example format:
```json
[
    {"x1": 100, "y1": 200, "x2": 400, "y2": 250},
    {"x1": 150, "y1": 300, "x2": 450, "y2": 350}
]
```
These coordinates define areas where specific data fields (e.g., GST numbers or monetary values) are located in the PDF.

## Output Format
The extracted data is saved in CSV format with columns:
- **File Name** - Name of the processed PDF file.
- **GST No** - Extracted GST number.
- **Goods Value** - Extracted monetary value.
- **Total Value** - Extracted total amount.

## Troubleshooting
- If `tabula-py` fails to read PDFs, ensure Java is installed and added to the system PATH.
- If no data is extracted, check if the template coordinates match the PDF layout.

## License
This project is open-source.

## Author
Ilavarasan 

