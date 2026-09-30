import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd

from core.reporter import Report
from core.validator import is_coordinates_valid
from excel.writer import write_workbook


def select_excel(): # needs restructurisation
    """Selection of the Excel file."""
    filename = fd.askopenfilename(title="Select the Exel Table", initialdir="./data", filetypes=[("Exel Table (.xlsx)", "*.xlsx")])

    if not filename:
        sys.exit()

    return filename

def select_database_file():
    """Selection of the data base file."""
    filename = fd.askopenfilename(title="Select the data base file", initialdir="./data", filetypes=[("addresses.json", "addresses.json")],
    )

    if not filename:
        sys.exit()

    return filename


def select_sheet(window, book, on_selected):
    popup_width = 250
    popup_height = 110

    window.update_idletasks()
    popup_x = window.winfo_rootx() + window.winfo_width() - popup_width
    popup_y = window.winfo_rooty()

    popup = tk.Toplevel(window)
    popup.resizable(width=False, height=False)
    popup.title("Select a Sheet")
    popup.geometry(f"{popup_width}x{popup_height}+{popup_x}+{popup_y}")
    popup["bg"] = "#333333"

    tk.Label(popup, text="Choose a sheet:", fg="white", bg="#333333").pack(pady=len(book.sheetnames))

    combo = ttk.Combobox(popup, values=book.sheetnames, state="readonly")
    combo.pack(pady=10)
    combo.current(0)

    def select_position():
        chosen = combo.current()
        sheet = book[book.sheetnames[chosen]]
        maxrow = sheet.max_row

        on_selected(sheet, maxrow)
        popup.destroy()

    tk.Button(popup, text="Select", command=select_position, fg="white", bg="#222222").pack(pady=5)

    popup.transient(window)
    popup.grab_set()
    window.wait_window(popup)


def enter_coordinates(window, missing_address, on_filled, structbutton):
    popup_width = 330
    popup_height = 130

    window.update_idletasks()
    popup_x = window.winfo_rootx() + window.winfo_width() - popup_width
    popup_y = window.winfo_rooty()

    popup = tk.Toplevel(window)
    popup.resizable(width=False, height=False)
    popup.title("Fill the missing coordinates")
    popup.geometry(f"{popup_width}x{popup_height}+{popup_x}+{popup_y}")
    popup["bg"] = "#333333"

    tk.Label(popup, text="Copy the geocoordinates", fg="white", bg="#333333").pack(pady=5)

    address_to_search = tk.Text(popup, width=40, height=1, state="normal")
    address_to_search.pack(pady=5)
    address_to_search.delete("1.0", tk.END)
    address_to_search.insert(tk.END, missing_address)
    address_to_search.config(state="disabled")

    text_box = tk.Text(popup, height=1, width=40)
    text_box.pack(pady=5)

    def print_to_textbox():
        value = text_box.get("1.0", tk.END).strip()

        if not is_coordinates_valid(value):
            return

        on_filled(value)
        popup.destroy()

    tk.Button(popup, text="Fill", command=print_to_textbox, fg="white", bg="#222222").pack(pady=5)

    structbutton.config(state="disabled")

    popup.transient(window)
    popup.grab_set()
    window.wait_window(popup)


def scan_folder(window, time):
    confirmed = False

    # Default data folders
    default_invoice_folder = "data/invoices"
    default_address_database_folder = "data"
    default_report_folder = "data"

    invoice_folder = None
    database_file = None
    report_folder = None
    starting_row = 1

    popup = tk.Toplevel(window)
    popup.title("Choose the following parameters")
    popup["bg"] = "#333333"

    tk.Label(
        popup, 
        text=" Choose the following parameters of close the window for defaults ",
        fg="white",
        bg="#333333"
    ).pack(pady=5)
    
    label_invoice = tk.Label(popup, text=default_invoice_folder, fg="black", bg="white")
    label_invoice.pack(pady=5)

    def invoice_folder_select():
        nonlocal invoice_folder
        invoice_folder = fd.askdirectory(title="Choose the invoice folder", initialdir=f"./{default_invoice_folder}")
        if invoice_folder:
            label_invoice.config(text=invoice_folder)

    tk.Button(popup, text="Invoice folder", command=invoice_folder_select, fg="white", bg="#222222").pack(pady=5)

    label_database = tk.Label(popup, text=default_address_database_folder, fg="black", bg="white")
    label_database.pack(pady=5)

    def database_folder_select():
        nonlocal database_file
        database_file = fd.askopenfilename(
        title="Choose the address database folder",
        initialdir=f"./{default_address_database_folder}",
        filetypes=[("addresses.json", "addresses.json")],
        )
        if database_file:
            label_database.config(text=database_file)

    tk.Button(popup, text="Database folder", command=database_folder_select, fg="white", bg="#222222").pack(pady=5)

    label_report = tk.Label(popup, text=default_report_folder, fg="black", bg="white")
    label_report.pack(pady=5)

    def report_folder_select():
        nonlocal report_folder
        report_folder = fd.askdirectory(title="Choose the report folder", initialdir=f"./{default_report_folder}")
        if report_folder:
            label_report.config(text=report_folder)

    tk.Button(popup, text="Report folder", command=report_folder_select, fg="white", bg="#222222").pack(pady=5)

    def confirm_settings():
        nonlocal invoice_folder, database_file, report_folder
        
        if invoice_folder == None:
            invoice_folder = default_invoice_folder
        if database_file == None:
            database_file = f"{default_address_database_folder}/addresses.json"
        if report_folder == None:
            report_folder = default_report_folder

        popup.destroy()
        window.destroy()

        reporter = Report(report_folder, time)

        write_workbook(invoice_folder, database_file, report_folder, starting_row, reporter, time)

        reporter.save()
        print(f"Process complete. You can find the report in {report_folder}")

    tk.Button(popup, text="Confirm", command=confirm_settings, fg="white", bg="#222222").pack(pady=5)
        
    popup.resizable(width=False, height=False)
    popup.wait_window()

    if not confirmed:
        sys.exit()