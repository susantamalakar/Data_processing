# excel_to_pdf_row_splitter.py

import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os
from tqdm import tqdm
from fpdf import FPDF

def main():
    # Hide Tkinter root window
    Tk().withdraw()

    # Open file dialog to select Excel file
    excel_file = askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])

    if not excel_file:
        print("No file selected. Exiting...")
        return

    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)

    # Delete the second, third, and fourth columns
    df.drop(df.columns[[1, 2, 3]], axis=1, inplace=True)

    # Create a folder to save the new PDF files
    output_folder = os.path.join(os.path.dirname(excel_file), "Row_PDFs")
    os.makedirs(output_folder, exist_ok=True)

    # Create PDFs for each row
    for i, row in tqdm(df.iterrows(), total=df.shape[0], desc="Processing rows"):
        row_pdf = FPDF()
        row_pdf.add_page()

        # Set font and size
        row_pdf.set_font("Arial", size=12)

        # Add the row data to the PDF, skipping empty or NaN values
        for idx, col_value in enumerate(row):
            if pd.notna(col_value) and col_value != '':
                row_pdf.cell(200, 10, txt=f"{df.columns[idx]}: {col_value}", ln=True)

        # Save the PDF file
        row_name = f"Row_{i+1}.pdf"
        output_path = os.path.join(output_folder, row_name)
        row_pdf.output(output_path)

    print(f"All rows have been saved as PDF files in the folder: {output_folder}")

if __name__ == "__main__":
    main()
