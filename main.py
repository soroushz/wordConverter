import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx2pdf import convert


def select_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)


def select_word_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        word_entry.delete(0, tk.END)
        word_entry.insert(0, file_path)


def pdf_to_word_ui():
    pdf_file = pdf_entry.get()
    if not pdf_file:
        messagebox.showerror("Error", "Please select a PDF file")
        return

    word_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    if word_file:
        cv = Converter(pdf_file)
        cv.convert(word_file)
        cv.close()
        messagebox.showinfo("Success", "PDF converted to Word successfully")
        pdf_entry.delete(0, tk.END)


def word_to_pdf_ui():
    word_file = word_entry.get()
    if not word_file:
        messagebox.showerror("Error", "Please select a Word file")
        return

    pdf_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if pdf_file:
        convert(word_file, pdf_file)
        messagebox.showinfo("Success", "Word converted to PDF successfully")
        word_entry.delete(0, tk.END)


# Creating the main window
root = tk.Tk()
root.title("PDF to Word and Word to PDF Converter")

# PDF to Word section
tk.Label(root, text="PDF to Word").grid(row=0, column=0, padx=10, pady=10)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_pdf_file).grid(row=0, column=2, padx=10, pady=10)
tk.Button(root, text="Convert", command=pdf_to_word_ui).grid(row=0, column=3, padx=10, pady=10)

# Word to PDF section
tk.Label(root, text="Word to PDF").grid(row=1, column=0, padx=10, pady=10)
word_entry = tk.Entry(root, width=50)
word_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_word_file).grid(row=1, column=2, padx=10, pady=10)
tk.Button(root, text="Convert", command=word_to_pdf_ui).grid(row=1, column=3, padx=10, pady=10)

# Start the GUI event loop
root.mainloop()
