import os
import tkinter as tk
from tkinter import filedialog
import PyPDF2
import pandas as pd
from tqdm import tqdm
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph


def extract_info_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    return text


def process_pdf(text):
    lines = text.split('\n')
    data = []
    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            data.append([key.strip(), value.strip()])
    return data


def save_table_as_pdf(data, output_path):
    pdf = SimpleDocTemplate(output_path, pagesize=letter)
    style = getSampleStyleSheet()

    table_data = [[Paragraph(cell, style['BodyText']) for cell in row] for row in data]

    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements = [table]
    pdf.build(elements)


def main():
    root = tk.Tk()
    root.withdraw()
    pdf_paths = filedialog.askopenfilenames(
        title="Select PDF Files",
        filetypes=[("PDF Files", "*.pdf")]
    )

    output_folder = filedialog.askdirectory(
        title="Select Output Folder"
    )

    if not pdf_paths or not output_folder:
        print("No files or output folder selected.")
        return

    for pdf_path in tqdm(pdf_paths, desc="Processing PDFs"):
        text = extract_info_from_pdf(pdf_path)
        data = process_pdf(text)
        output_path = os.path.join(output_folder, os.path.basename(pdf_path))
        save_table_as_pdf(data, output_path)

    print("PDF processing complete. Files saved in:", output_folder)


if __name__ == "__main__":
    main()
