import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def main(pdf_folder, excel_file, dest_folder, ticket_col, status_col):
    try:
        # Create destination folder if it doesn't exist
        os.makedirs(dest_folder, exist_ok=True)
        
        # Read the Excel file
        df = pd.read_excel(excel_file)
        
        # Get ticket numbers to move (Replied and Forwarded)
        statuses_to_move = ['Replied', 'Forwarded']
        tickets_to_move = set(df[df[status_col].isin(statuses_to_move)][ticket_col].astype(str))
        
        moved_count = 0
        # Iterate through PDF files
        for file in os.listdir(pdf_folder):
            if file.endswith('.pdf'):
                # Parse ticket number from filename
                name = file[:-4]  # Remove .pdf extension
                parts = name.split()
                ticket = parts[-1]  # Last part is ticket number
                
                if ticket in tickets_to_move:
                    src_path = os.path.join(pdf_folder, file)
                    dest_path = os.path.join(dest_folder, file)
                    shutil.move(src_path, dest_path)
                    moved_count += 1
        
        messagebox.showinfo("Success", f"Process completed. {moved_count} files moved.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def browse_excel():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(0, filename)

def browse_pdf_folder():
    folder = filedialog.askdirectory()
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, folder)

def browse_dest_folder():
    folder = filedialog.askdirectory()
    dest_entry.delete(0, tk.END)
    dest_entry.insert(0, folder)

def run():
    excel = excel_entry.get()
    pdf = pdf_entry.get()
    dest = dest_entry.get()
    ticket_col = ticket_entry.get()
    status_col = status_entry.get()
    
    if not all([excel, pdf, dest, ticket_col, status_col]):
        messagebox.showwarning("Input Error", "Please fill all fields.")
        return
    
    main(pdf, excel, dest, ticket_col, status_col)

# Create GUI
root = tk.Tk()
root.title("PDF Mover Tool")

# Excel file
tk.Label(root, text="Excel File Path:").grid(row=0, column=0, sticky="e")
excel_entry = tk.Entry(root, width=50)
excel_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=browse_excel).grid(row=0, column=2)

# PDF folder
tk.Label(root, text="PDF Folder Path:").grid(row=1, column=0, sticky="e")
pdf_entry = tk.Entry(root, width=50)
pdf_entry.grid(row=1, column=1)
tk.Button(root, text="Browse", command=browse_pdf_folder).grid(row=1, column=2)

# Destination folder
tk.Label(root, text="Destination Folder Path:").grid(row=2, column=0, sticky="e")
dest_entry = tk.Entry(root, width=50)
dest_entry.grid(row=2, column=1)
tk.Button(root, text="Browse", command=browse_dest_folder).grid(row=2, column=2)

# Ticket column
tk.Label(root, text="Ticket Column Name:").grid(row=3, column=0, sticky="e")
ticket_entry = tk.Entry(root, width=50)
ticket_entry.grid(row=3, column=1)
ticket_entry.insert(0, "Ticket Number")  # Default

# Status column
tk.Label(root, text="Status Column Name:").grid(row=4, column=0, sticky="e")
status_entry = tk.Entry(root, width=50)
status_entry.grid(row=4, column=1)
status_entry.insert(0, "Status")  # Default

# Run button
tk.Button(root, text="Run", command=run).grid(row=5, column=1)

root.mainloop()