import tkinter as tk
from tkinter import filedialog, messagebox

# Initialize global variables
excel_path = ""
login_password = ""


def select_excel_file(label_excel):
    global excel_path
    file_path = filedialog.askopenfilename(title="Select Excel File",
                                           filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        excel_path = file_path
        label_excel.config(text=f"📄 {excel_path}", fg="green")


def submit(root, password_entry):
    global login_password, excel_path

    login_password = password_entry.get()

    if not login_password:
        messagebox.showwarning("⚠ Warning", "Password is required!")
        return
    if not excel_path:
        messagebox.showwarning("⚠ Warning", "Please select an Excel file.")
        return

    # Show confirmation
    messagebox.showinfo("✅ Confirmation",
                        f"Excel File: {excel_path}\n"
                        f"Password: {'*' * len(login_password)}")

    root.destroy()  # Close the GUI after submission


def get_basic_info2():
    global excel_path, login_password

    # Initialize GUI
    root = tk.Tk()
    root.title("Excel File and Password")
    root.geometry("500x400")  # Adjusted height for the input fields
    root.resizable(False, False)

    # Styling
    label_font = ("Comic Sans MS", 14)
    button_font = ("Comic Sans MS", 10, "bold")
    entry_font = ("Verdana", 12)

    # Select Excel File Section
    frame_excel = tk.Frame(root, padx=10, pady=5)
    frame_excel.pack(fill="x")

    tk.Label(frame_excel, text="Excel File:", font=label_font).pack(anchor="w")
    label_excel = tk.Label(frame_excel, text="No file selected", fg="red", font=label_font)
    label_excel.pack(anchor="w")
    tk.Button(frame_excel, text="📄 Select Excel File", font=button_font,
              command=lambda: select_excel_file(label_excel), cursor="hand2").pack(fill="x", pady=5)

    # Password Input Section
    frame_password = tk.Frame(root, padx=10, pady=5)
    frame_password.pack(fill="x")

    tk.Label(frame_password, text="Enter Password:", font=label_font).pack(anchor="w")
    password_entry = tk.Entry(frame_password, show="*", font=entry_font)
    password_entry.pack(fill="x", pady=3)

    # Submit Button - At the Bottom
    submit_btn = tk.Button(root, text="✅ Submit", font=button_font,
                           command=lambda: submit(root, password_entry), cursor="hand2")
    submit_btn.pack(side="bottom", fill="x", pady=50, padx=20)

    # Run the GUI
    root.mainloop()

    return excel_path, login_password

