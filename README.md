# PDF Table Extractor to Excel using Tabula and Tkinter

This Python script allows you to **extract tables from multiple PDF files** using a Tabula JSON template and **consolidate** the extracted tables into a **single Excel workbook**. The script includes a simple **GUI-based file picker** using `tkinter`.

## Features

- Select a folder containing multiple PDFs.
- Select a Tabula JSON extraction template.
- Automatically extract tables from each PDF using the template.
- Combine all tables into one Excel sheet with:
  - Filenames as labels.
  - Hyperlinks to the original PDFs.
- Output is saved in a single `.xlsx` file.

## Requirements

### Python Libraries

Ensure you have the following libraries installed:

```bash
pip install pandas openpyxl tabula-py
```

Additionally, **Java must be installed and available in your system path**, as `tabula-py` is a wrapper for the Java-based Tabula.

## How to Use

1. **Run the script**:
   ```bash
   python extract_pdf_tables_to_excel.py
   ```

2. **Choose the input folder** when prompted – this should contain all the PDF files you want to extract tables from.

3. **Select your Tabula JSON template** – this controls how tables are detected and extracted from the PDFs.

4. **Choose the output location and file name** for the Excel file.

5. The script will process all PDFs in the folder and write the tables to the Excel sheet `All_Tables`.

## Output Format

- **Excel Sheet Name:** `All_Tables`
- Each row begins with the **PDF filename**.
- Tables are inserted in adjacent columns.
- A **hyperlink** to the original PDF is added at the end of each row.
- Rows are spaced based on the number of rows in the first table extracted from each file.

## Notes

- This script assumes all PDFs follow a similar format compatible with the selected Tabula template.
- If no tables are found or an error occurs for a PDF, the script will continue processing the rest.
- You can modify the script to handle multiple sheets or different layouts as needed.

## Example Tabula Template

To create a JSON template:
1. Open [Tabula GUI](https://tabula.technology/).
2. Load a sample PDF.
3. Select the areas containing tables.
4. Export the selection as a JSON template.

## Troubleshooting

- **Tabula Not Working?** Ensure Java is properly installed and accessible via `java -version`.
- **Incorrect Extraction?** Check your Tabula template. It must match the layout of the PDF tables.
- **Empty Excel Output?** Ensure the PDFs actually contain extractable tables (not scanned images).

## License
This project is open-source.

## Author
Ilavarasan

