import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Initialize global variables
excel_path = ""
base_path = ""
login_password = ""
report_var = None
upload_files_var = None
label_path = None
label_excel = None

def select_folder():
    global base_path, label_path
    base_path = filedialog.askdirectory(title="Select files Folder")
    if base_path:
        label_path.config(text=f"Selected: {base_path}")

def select_excel_file():
    global excel_path, label_excel
    file_path = filedialog.askopenfilename(title="Select Excel File",
                                           filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        excel_path = file_path
        label_excel.config(text=f"Selected: {excel_path}")

def submit(root):
    global report_var, upload_files_var, base_path, excel_path, login_password
    report = report_var.get()
    upload_files = upload_files_var.get()

    # Ask for password
    login_password = simpledialog.askstring("Password", "Enter the site password:", show="*")
    if not login_password:
        messagebox.showwarning("Warning", "Password is required!")
        return

    if not base_path:
        messagebox.showwarning("Warning", "Please select a base folder.")
        return

    if not excel_path:
        messagebox.showwarning("Warning", "Please select an Excel file.")
        return

    # Show confirmation
    messagebox.showinfo("Confirmation",
                        f"Base Path: {base_path}\n"
                        f"Excel File: {excel_path}\n"
                        f"Report: {'Yes' if report else 'No'}\n"
                        f"Upload Files: {'Yes' if upload_files else 'No'}\n"
                        f"Password: {f'{login_password}'}")

    root.destroy()  # Close the GUI after submission

def get_password():
    global login_password
    login_password = simpledialog.askstring("Password", "Enter the site password:", show="*")
    if login_password:
        messagebox.showinfo("Password", "Password set successfully.")

def get_basic_info():
    global report_var, upload_files_var, label_path, label_excel

    # Initialize GUI
    root = tk.Tk()
    root.title("Settings")

    # Select folder
    tk.Label(root, text="Select the files folder (where the recipes are located):").pack()
    label_path = tk.Label(root, text="No folder selected", fg="red")
    label_path.pack()
    tk.Button(root, text="Select Folder", command=select_folder).pack()

    # Select Excel file
    tk.Label(root, text="Excel File:").pack()
    label_excel = tk.Label(root, text="No file selected", fg="red")
    label_excel.pack()
    tk.Button(root, text="Select Excel File", command=select_excel_file).pack()

    # Checkboxes
    report_var = tk.BooleanVar()
    upload_files_var = tk.BooleanVar()

    tk.Checkbutton(root, text="Fill a REPORT of the patients?", variable=report_var).pack()
    tk.Checkbutton(root, text="Upload FILES for the patients?", variable=upload_files_var).pack()

    # Submit button
    tk.Button(root, text="Submit", command=lambda: submit(root)).pack()

    # Run the GUI
    root.mainloop()

    report = report_var.get()
    upload_files = upload_files_var.get()
    print("INPUT:", base_path, excel_path, report, upload_files)
    return base_path, excel_path, report, upload_files,login_password
