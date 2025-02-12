import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import tabula

# Function to process PDFs
def process_pdfs_in_folder(input_folder, output_file, template_path):
    try:
        areas = load_template(template_path)  # Load the template and extract areas
        if not areas:
            messagebox.showerror("Error", "Invalid template file. No valid areas found.")
            return

        pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
        if not pdf_files:
            messagebox.showerror("Error", "No PDF files found in the selected folder.")
            return

        all_data = []
        for pdf_file in pdf_files:
            input_path = os.path.join(input_folder, pdf_file)
            tables = tabula.read_pdf(input_path, area=areas, pages='all', multiple_tables=True)
            row_data = {"File Name": pdf_file}
            if tables:
                # print(f"Tables found in {pdf_file}.")
                for table in tables:
                    if table is not None:  # Ensure the table is not None
                        # print(f"Extracted table data from {pdf_file}:\n{table}")
                        # Extract values from the column names
                        for col_name in table.columns:
                            if pd.notna(col_name):
                                value = str(col_name).strip()
                                # print(value)
                                # Assign values to specific keys based on patterns
                                if len(value) == 15 and value.isalnum():  # Likely a GST No
                                    row_data["GST No"] = value
                                elif "," in value and value.replace(",", "").isdigit():  # Likely a monetary value
                                    if "Goods Value" not in row_data:
                                        row_data["Goods Value"] = value
                                    else:
                                        row_data["Total Value"] = value
            all_data.append(row_data)

        if all_data:
            final_df = pd.DataFrame(all_data)
            final_df.to_csv(output_file, index=False)
            messagebox.showinfo("Success", f"Data has been successfully saved to {output_file}.")
        else:
            messagebox.showerror("Error", "No data extracted from the PDFs.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to load the template JSON file and extract coordinates
def load_template(template_path):
    import json
    with open(template_path, 'r') as f:
        template = json.load(f)

    areas = []
    if isinstance(template, list):
        for area in template:
            if 'x1' in area and 'y1' in area and 'x2' in area and 'y2' in area:
                areas.append([area['y1'], area['x1'], area['y2'], area['x2']])
    return areas

# Functions for "Browse" buttons
def browse_folder(entry_widget):
    folder = filedialog.askdirectory(title="Select Source Folder")
    if folder:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder)

def browse_file(entry_widget, file_type="file"):
    if file_type == "file":
        file_path = filedialog.askopenfilename(title="Select Template File", filetypes=[("JSON Files", "*.json")])
    else:
        file_path = filedialog.asksaveasfilename(title="Save Output File", defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

# Main GUI application
def main_gui():
    def start_processing():
        input_folder = input_folder_entry.get()
        template_path = template_file_entry.get()
        output_file = output_file_entry.get()

        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Please select a valid source folder.")
            return
        if not template_path or not os.path.isfile(template_path):
            messagebox.showerror("Error", "Please select a valid template JSON file.")
            return
        if not output_file:
            messagebox.showerror("Error", "Please specify a valid output file path.")
            return

        process_pdfs_in_folder(input_folder, output_file, template_path)

    root = tk.Tk()
    root.title("PDF Data Extractor")

    # Input Folder
    tk.Label(root, text="Select Source Folder:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
    input_folder_entry = tk.Entry(root, width=50)
    input_folder_entry.grid(row=0, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_folder(input_folder_entry)).grid(row=0, column=2, padx=10, pady=5)

    # Template File
    tk.Label(root, text="Select Template File:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
    template_file_entry = tk.Entry(root, width=50)
    template_file_entry.grid(row=1, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_file(template_file_entry, file_type="file")).grid(row=1, column=2, padx=10, pady=5)

    # Output File
    tk.Label(root, text="Select Output File:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
    output_file_entry = tk.Entry(root, width=50)
    output_file_entry.grid(row=2, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_file(output_file_entry, file_type="save")).grid(row=2, column=2, padx=10, pady=5)

    # Start Button
    tk.Button(root, text="Start Processing", command=start_processing, bg="green", fg="white").grid(row=3, column=0, columnspan=3, pady=10)

    root.mainloop()

# Run the GUI
if __name__ == "__main__":
    main_gui()
