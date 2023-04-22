import tkinter as tk
from tkinter import filedialog
import pdf2docx
import os

def select_file():
    """Open a file dialog window to select a PDF file"""
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF Files", "*.pdf")])
    pdf_path_entry.delete(0, tk.END)
    pdf_path_entry.insert(0, file_path)

def convert_pdf_to_word():
    """Convert the selected PDF file to a Word file"""
    pdf_path = pdf_path_entry.get()
    word_path = word_path_entry.get()

    downloads_dir = os.path.expanduser("~/Downloads")

    initial_dir = downloads_dir

    file_path = filedialog.asksaveasfilename(
        initialdir=initial_dir,
        initialfile=word_path,
        title="Save Word File",
        defaultextension=".docx",
        filetypes=[("Word Files", "*.docx")],
    )

    pdf2docx.parse(pdf_path, file_path)

root = tk.Tk()
root.title("PDF to Word Converter")
root.geometry("500x500")
root.configure(padx=200, pady=200)
root.configure(padx=20, pady=20, bg="thistle")

font = ("TkDefaultFont", 14)

pdf_path_label = tk.Label(root, text="PDF File:", font=font,bg="LightBlue1")
pdf_path_label.grid(row=0, column=0, padx=10, pady=10)

pdf_path_entry = tk.Entry(root, font=font)
pdf_path_entry.grid(row=0, column=1, padx=10, pady=10)

pdf_path_button = tk.Button(root, text="Select File", command=select_file, font=font,bg="lightgoldenrod")
pdf_path_button.grid(row=0, column=2, padx=10, pady=10)

word_path_label = tk.Label(root, text="Word File:", font=font,bg="LightBlue1")
word_path_label.grid(row=1, column=0, padx=10, pady=10)

word_path_entry = tk.Entry(root, font=font)
word_path_entry.grid(row=1, column=1, padx=10, pady=10)

convert_button = tk.Button(root, text="Convert to Word", command=convert_pdf_to_word, font=font,bg="limegreen")
convert_button.grid(row=2, column=1, padx=10, pady=10)

root.mainloop()