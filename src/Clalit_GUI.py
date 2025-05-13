import tkinter as tk
from tkinter import filedialog, messagebox

# Initialize global variables
excel_path = ""
base_path = ""
login_password = ""


def select_folder(label_path):
    global base_path
    base_path = filedialog.askdirectory(title="Select Files Folder")
    if base_path:
        label_path.config(text=f"📁 {base_path}", fg="green")


def select_excel_file(label_excel):
    global excel_path
    file_path = filedialog.askopenfilename(title="Select Excel File",
                                           filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        excel_path = file_path
        label_excel.config(text=f"📄 {excel_path}", fg="green")


def submit(root, password_entry, report_var, upload_files_var):
    global login_password, base_path, excel_path

    login_password = password_entry.get()
    report = report_var.get()
    upload_files = upload_files_var.get()

    if not login_password:
        messagebox.showwarning("⚠ Warning", "Password is required!")
        return
    if not base_path:
        messagebox.showwarning("⚠ Warning", "Please select a base folder.")
        return
    if not excel_path:
        messagebox.showwarning("⚠ Warning", "Please select an Excel file.")
        return

    # Show confirmation
    messagebox.showinfo("✅ Confirmation",
                        f"Base Path: {base_path}\n"
                        f"Excel File: {excel_path}\n"
                        f"Report: {'Yes' if report else 'No'}\n"
                        f"Upload Files: {'Yes' if upload_files else 'No'}\n"
                        f"Password: {'*' * len(login_password)}")

    root.destroy()  # Close the GUI after submission


def get_basic_info():
    global base_path, excel_path, login_password

    # Initialize GUI
    root = tk.Tk()
    root.title("Settings")
    root.geometry("600x800")  # Adjusted height to ensure space for Submit
    root.resizable(False, False)

    # Styling - Different Fonts for Labels, Buttons, and Entries
    label_font = ("Comic Sans MS", 14)

    button_font = ("Comic Sans MS", 10, "bold")
    entry_font = ("Verdana", 12)

    # Select Folder Section
    frame_folder = tk.Frame(root, padx=10, pady=5)
    frame_folder.pack(fill="x")

    tk.Label(frame_folder, text="Select Files Folder:", font=label_font).pack(anchor="w")
    label_path = tk.Label(frame_folder, text="No folder selected", fg="red", font=label_font)
    label_path.pack(anchor="w")
    tk.Button(frame_folder, text="📁 Select Folder", font=button_font,
              command=lambda: select_folder(label_path), cursor="hand2").pack(fill="x", pady=5)

    # Select Excel File Section
    frame_excel = tk.Frame(root, padx=10, pady=5)
    frame_excel.pack(fill="x")

    tk.Label(frame_excel, text="Excel File:", font=label_font).pack(anchor="w")
    label_excel = tk.Label(frame_excel, text="No file selected", fg="red", font=label_font)
    label_excel.pack(anchor="w")
    tk.Button(frame_excel, text="📄 Select Excel File", font=button_font,
              command=lambda: select_excel_file(label_excel), cursor="hand2").pack(fill="x", pady=5)

    # Password Input
    frame_password = tk.Frame(root, padx=10, pady=5)
    frame_password.pack(fill="x")

    tk.Label(frame_password, text="Enter Site Password:", font=label_font).pack(anchor="w")
    password_entry = tk.Entry(frame_password, show="*", font=entry_font)
    password_entry.pack(fill="x", pady=3)

    # Checkboxes Section
    report_var = tk.BooleanVar()
    upload_files_var = tk.BooleanVar()

    frame_check = tk.Frame(root, padx=10, pady=5)
    frame_check.pack(fill="x")

    tk.Checkbutton(frame_check, text="Generate Patient Report", variable=report_var, font=label_font, cursor="hand2").pack(anchor="w")
    tk.Checkbutton(frame_check, text="Upload Files for Patients", variable=upload_files_var, font=label_font, cursor="hand2").pack(
        anchor="w")

    # Submit Button - At the Bottom
    submit_btn = tk.Button(root, text="✅ Submit", font=button_font,
                           command=lambda: submit(root, password_entry, report_var, upload_files_var), cursor="hand2")
    submit_btn.pack(side="bottom", fill="x", pady=50, padx=20)

    # Run the GUI
    root.mainloop()

    return base_path, excel_path, report_var.get(), upload_files_var.get(), login_password

