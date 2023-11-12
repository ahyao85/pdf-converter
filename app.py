import tkinter as tk
from tkinter import filedialog, messagebox
import camelot
import pandas as pd
from pdf2docx import Converter

def convert_pdf_to_word(pdf_path, output_word_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(output_word_path, start=0, end=None)
        cv.close()
        messagebox.showinfo("Success", f"PDF converted to Word and saved as {output_word_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def convert_pdf_to_excel(pdf_path, output_excel_path):
    try:
        tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        if len(tables) > 0:
            df = tables[0].df
            for table in tables[1:]:
                df = pd.concat([df, table.df], ignore_index=True)
        else:
            raise ValueError("No tables found on any page of the PDF.")
        df.to_excel(output_excel_path, index=False)
        messagebox.showinfo("Success", f"PDF converted to Excel and saved as {output_excel_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_path.set(file_path)

def save_output():
    file_path = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx"), ("Word files", "*.docx")])
    if file_path:
        output_path.set(file_path)  # Use a single StringVar for the output path

def convert_file():
    input_path = pdf_path.get()
    output_path_str = output_path.get()
    
    if not input_path:
        messagebox.showwarning("Warning", "Please select a PDF file.")
        return
        
    if not output_path_str:
        messagebox.showwarning("Warning", "Please specify an output file.")
        return

    # Determine the conversion type based on the file extension
    try:
        if output_path_str.endswith('.xlsx'):
            convert_pdf_to_excel(input_path, output_path_str)
        elif output_path_str.endswith('.docx'):
            convert_pdf_to_word(input_path, output_path_str)
        else:
            messagebox.showerror("Error", "Unknown file type selected. Please choose '.xlsx' or '.docx'.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

app = tk.Tk()
app.title("PDF Converter")

pdf_path = tk.StringVar()
output_path = tk.StringVar()  # Use a single StringVar for both Excel and Word output

# Create and place widgets
tk.Label(app, text="Select PDF File:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(app, textvariable=pdf_path, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=select_pdf).grid(row=0, column=2, padx=10, pady=10)

tk.Label(app, text="Save Output File As:").grid(row=1, column=0, padx=10, pady=10)
tk.Entry(app, textvariable=output_path, width=50).grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Save As", command=save_output).grid(row=1, column=2, padx=10, pady=10)

tk.Button(app, text="Convert", command=convert_file).grid(row=2, column=1, padx=10, pady=20)

app.mainloop()