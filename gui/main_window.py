import re
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime
from pathlib import Path

from excel.writer import write_workbook
from excel.scanner import header_def_distionary
from core.address_database import AddressDatabase
from core.calculator import calculate_distances
from gui.dialog import *
from core.validator import validate_addresses
from core.reporter import Report

def main_window():
    window = tk.Tk()
    window.title("Invoice Travel Reporter")
    window["bg"] = "#222222"
    label_font = tkfont.Font(family="Calibri", size=14)

    now = datetime.now()

    tk.Label(window, 
            text=f" Would you like to scan a folder of invoices or select an existing \n Excel spreadsheet to enter distances or addresses? ",
            fg="white",
            bg="#222222",
            font=label_font,
            justify="center",
        ).pack(pady=5)

    def window_scan_folder():
        scan_folder(window, now)

    select_button_1 = tk.Button(window, text="Scan folder", state="active", command=window_scan_folder, fg="white", bg="#111111", font=label_font)
    select_button_1.pack(pady=5)

    def existing_file():
        filename = select_excel()
        if not filename:
            return

        book = openpyxl.load_workbook(filename)

        # Remove everything from the current window
        for widget in window.winfo_children():
            widget.destroy()

        # Reuse the existing window
        ManualWindow(window, filename, book, now)

    select_button_2 = tk.Button(window, text="Existing Excel file", state="active", command=existing_file, fg="white", bg="#111111", font=label_font)
    select_button_2.pack(pady=5)

    window.mainloop()


class ManualWindow:
    def __init__(self, window, filename, book, time):
        self.window = window
        self.filename = filename
        self.book = book
        self.time = time

        self.label_width = 405
        self.box_width = 170
        self.label_font = tkfont.Font(family="Calibri", size=14)

        self.sheet = None
        self.maxrow = None
        self.starting_row = 1

        # Build the interface.
        self._build_ui()


    def save_excel(self):
        """Save table and close"""
        self.book.save(self.filename)


    def truncation(self, text, width, font):
        "The truncated table name display (if it's too long)"
        text = str(text)
        if font.measure(text) <= width:
            return text

        for i in range(len(text)):
            truncated = "..." + text[i:]
            if font.measure(truncated) <= self.label_width:
                return truncated

        return "..."


    def _build_ui(self):
        self.window.resizable(width=False, height=False)
        self.window.title("Distance Calculator")
        self.window.geometry("720x255")
        self.window["bg"] = "#222222"

        # File button and label definition
        self.filebutton = tk.Button(self.window, text="File", state="disabled", width=1, fg="white", bg="#111111", font=self.label_font)
        self.filebutton.place(x=20, y=20, width=self.label_font.measure("  Sheet  "))

        self.filelabel = tk.Label(
            self.window,
            text=self.truncation(self.filename, self.label_width, self.label_font),
            width=1,
            anchor="e",
            fg="black",
            bg="white",
            font=self.label_font,
            justify="right",
        )
        self.filelabel.place(x=20+self.label_font.measure("  Sheet  ")+20, y=25, width=self.label_width)

        # Sheet button and label
        self.sheetbutton = tk.Button(self.window, text="Sheet", command=self.sheet_select, width=1, fg="white", bg="#111111", font=self.label_font)
        self.sheetbutton.place(x=20, y=65, width=self.label_font.measure("  Sheet  "))

        self.sheetlabel = tk.Label(self.window, width=1, anchor="e", fg="black", bg="white", font=self.label_font, justify="right")
        self.sheetlabel.place(x=20 + self.label_font.measure("  Sheet  ") + 20, y=70, width=self.label_width)

        # Table structure definition
        self.datebox = ttk.Combobox(self.window, state="readonly")
        self.datebox.place(x=20 + self.label_font.measure("  Structure  ") + 20, y=160, width=self.box_width)

        self.addressbox = ttk.Combobox(self.window, state="readonly")
        self.addressbox.place(x=20 + self.label_font.measure("  Structure  ") + 20 + self.box_width + 20, y=160, width=self.box_width)

        self.distbox = ttk.Combobox(self.window, state="readonly")
        self.distbox.place(x=20 + self.label_font.measure("  Structure  ") + 20 + self.box_width + 20 + self.box_width + 20, y=160, width=self.box_width)

        self.datelabel = tk.Label(self.window, width=1, text="Dates", fg="white", bg="#222222", font=self.label_font)
        self.datelabel.place(x=20 + self.label_font.measure("  Structure  ") + 20, y=130, width=self.box_width)

        self.addresslabel = tk.Label(self.window, width=1, text="Addresses", fg="white", bg="#222222", font=self.label_font)
        self.addresslabel.place(x=20 + self.label_font.measure("  Structure  ") + 20 + self.box_width + 20, y=130, width=self.box_width)

        self.distlabel = tk.Label(self.window, width=1, text="Distances", fg="white", bg="#222222", font=self.label_font)
        self.distlabel.place(x=20 + self.label_font.measure("  Structure  ") + 20 + self.box_width + 20 + self.box_width + 20, y=130, width=self.box_width)

        self.structbutton = tk.Button(self.window, text="Structure", state="disabled", command=self.data_select, width=1, fg="white", bg="#111111", font=self.label_font)
        self.structbutton.place(x=20, y=140, width=self.label_font.measure("  Structure  "))

        # Calculation
        self.processbutton = tk.Button(self.window, state="disabled", text="PROCESS", command=self.process_data, width=1, fg="white", bg="#111111")
        self.processbutton.place(x=20, y=200, width=self.label_font.measure("  Structure  "))

        self.resultlabel = tk.Label(self.window, width=1, fg="black", bg="white", font=self.label_font)
        self.resultlabel.place(x=20 + self.label_font.measure("  Structure  ") + 20, y=205, width=self.label_width - 30)


    def sheet_select(self):
        self.datebox.set("")
        self.addressbox.set("")
        self.distbox.set("")

        def on_selected(sheet, maxrow):
            self.sheet = sheet
            self.maxrow = maxrow

            self.sheetlabel.config(text=self.truncation(sheet, self.label_width, self.label_font))

            for row in self.sheet.iter_rows(min_col=1, max_col=self.sheet.max_column):
                if all(re.search(r"[a-zA-Z]", str(cell.value or "")) for cell in row):
                    self.starting_row = row[0].row
                    break

            row_list = [sheet.cell(row=self.starting_row, column=column).value for column in range(1, sheet.max_column + 1)]

            self.datebox["values"] = row_list
            self.addressbox["values"] = row_list
            self.distbox["values"] = row_list
                        
            self.structbutton.config(state="active")

        select_sheet(self.window, self.book, on_selected)


    def data_select(self):
        if self.addressbox.get() == "" or self.addressbox.get() == "None":
            return

        self.sheetbutton.config(state="disabled")

        addresscolumn = self.addressbox.current() + 1

        self.database_file = select_database_file()

        database = AddressDatabase(self.database_file)
        database.load()

        def get_coordinates(missing_address):
            result = {"value": None}

            def on_filled(value):
                result["value"] = value

            enter_coordinates(self.window, missing_address, on_filled, self.structbutton)
            return result["value"]

        database.add_missing_addresses(self.sheet, addresscolumn, self.starting_row, self.sheet.max_row + 1, get_coordinates)
        database.save()

        self.processbutton.config(state="active")
        self.structbutton.config(state="disabled")
        self.datebox.config(state="disabled")
        self.addressbox.config(state="disabled")
        self.distbox.config(state="disabled")


    def process_data(self):
        if (
            self.addressbox.get() == ""
            or self.addressbox.get() == "None"
            or self.datebox.get() == ""
            or self.datebox.get() == "None"
            or self.distbox.get() == ""
            or self.distbox.get() == "None"
        ):
            return

        self.processbutton.config(state="disabled")
        self.datebox.config(state="disabled")
        self.distbox.config(state="disabled")

        addresscolumn = self.addressbox.current() + 1
        datecolumn = self.datebox.current() + 1
        distcolumn = self.distbox.current() + 1

        file_folder = Path(self.filename).parent
        file_name = Path(self.filename).with_suffix(".pdf").name

        reporter = Report(file_folder, self.time)

        database = AddressDatabase(self.database_file)
        database.load()

        missing_addresses, missing_rows = validate_addresses(self.sheet, addresscolumn, self.starting_row + 1, database)

        reporter.report_missing_addresses(missing_addresses, missing_rows)

        distances = calculate_distances(self.sheet, self.maxrow, addresscolumn, datecolumn, distcolumn, database.get_data(), first_data_row=self.starting_row + 1, missing_addresses=missing_addresses)
        # distances: [home_customer_distance, customer_customer_dist, customer_home_dist]

        headers = [cell.value for cell in self.sheet[self.starting_row]]
        headers = headers[:-2]

        date_headers, float_headers, currency_headers, text_headers = header_def_distionary()

        reporter.report_financial(self.sheet, self.starting_row, headers, date_headers, float_headers, currency_headers)
        reporter.report_travel(self.sheet, self.starting_row, headers, date_headers, text_headers, distcolumn, distances)
        # Aggregate the financial and travel metrics from the populated
        # worksheet and the distance-calculation results.

        columns = sum(1 for cell in self.sheet[self.starting_row] if cell.value not in (None, ""))

        for row in range(self.starting_row + 1, self.sheet.max_row + 1):
            for column in list(range(1, len(headers) + 1)):
                if self.sheet.cell(row=row, column=column).value in (None, ""):
                    reporter.report_cell_missmatch(row)
                    self.sheet.cell(row=row, column=column).fill = PatternFill(fgColor="FF0000", fill_type="solid")

        reporter.save(filename=file_name, folder=file_folder)                    

        self.resultlabel.config(text="Process complete")

        self.save_excel()