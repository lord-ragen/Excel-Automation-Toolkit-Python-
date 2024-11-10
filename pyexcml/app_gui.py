import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pyexcml import excel_to_xml  


file_path = None

def upload_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        status_label.config(text=f"Selected file: {file_path}")
    else:
        messagebox.showwarning("No File", "Please select a file.")

def generate_xml():
    if file_path:
        try:
            
            output_path = os.path.splitext(file_path)[0] + ".xml"
            excel_to_xml(file_path, output_path)  
            
            messagebox.showinfo("Success", f"XML file generated: {output_path}")
            status_label.config(text=f"XML file generated at: {output_path}")
        except FileNotFoundError:
            messagebox.showerror("File Not Found", f"The file '{file_path}' could not be found.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            status_label.config(text="Error occurred during XML generation.")
    else:
        messagebox.showwarning("No File Selected", "Please upload an Excel file first.")


root = tk.Tk()
root.title("Exceml")
root.geometry("400x300")
root.configure(bg="#f0f0f5")


root.iconbitmap(r'C:\Users\Welcome\source\repos\pyexcml\pyexcml\exceml.ico')  # Icon path

# Header
header_label = tk.Label(root, text="Excel to XML Converter", font=("Helvetica", 16, "bold"), bg="#f0f0f5")
header_label.pack(pady=15)

# Upload Button
upload_button = tk.Button(root, text="Upload Excel File", command=upload_file, font=("Helvetica", 12),
                          bg="#007acc", fg="white", activebackground="#005f99", activeforeground="white")
upload_button.pack(pady=10, ipadx=10, ipady=5)

# Generate XML Button
generate_button = tk.Button(root, text="Generate XML", command=generate_xml, font=("Helvetica", 12),
                            bg="#007acc", fg="white", activebackground="#005f99", activeforeground="white")
generate_button.pack(pady=10, ipadx=20, ipady=5)

# Status Label
status_label = tk.Label(root, text="Please upload an Excel file to convert.", font=("Helvetica", 10), bg="#f0f0f5")
status_label.pack(pady=5)

# Run the Tkinter event loop
root.mainloop()
