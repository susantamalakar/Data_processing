import os
import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter
from tqdm import tqdm


def convert_pdf_to_word(pdf_path, output_path):
    cv = Converter(pdf_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()


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

    for pdf_path in tqdm(pdf_paths, desc="Converting PDFs to Word with Tables"):
        output_path = os.path.join(output_folder, os.path.splitext(os.path.basename(pdf_path))[0] + '.docx')
        convert_pdf_to_word(pdf_path, output_path)

    print("PDF to Word conversion complete. Files saved in:", output_folder)


if __name__ == "__main__":
    main()
