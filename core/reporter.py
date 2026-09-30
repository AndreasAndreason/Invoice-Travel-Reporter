from datetime import datetime
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth


from core.validator import *


class Report:
    def __init__(self, report_folder, time):

        # Statistics
        # general
        self.invoice_number = 0
        self.report_folder = report_folder
        self.time = time

        # finance & work
        self.amount_sum = 0
        self.amount_prise_average = 0
        self.total_sum_prise = 0

        #travel
        self.work_days = 0
        self.customers_visited = 0
        self.total_distance = 0
        self.home_customer_dist = 0
        self.customer_customer_dist = 0
        self.customer_home_dist = 0


        # Bugs
        self.missing_invoice_number = 0
        self.missing_invoices = []

        self.header_mismatch_num = 0
        self.header_mismatch_file = []
        self.header_mismatch_row = []

        self.missing_address_num = 0
        self.missing_address = []
        self.missing_address_rows = []


    def report_invoice_files(self, invoice_files: list, invoice_folder):
        missing_invoices = validate_invoice_file_number(invoice_files, invoice_folder)
        missing_invoices = [Path(missing_invoice).name for missing_invoice in missing_invoices]

        self.missing_invoices = missing_invoices
        self.missing_invoice_number = len(missing_invoices)
        self.invoice_number = len(invoice_files)


    def report_cell_missmatch(self, row):
        self.header_mismatch_num += 1
        if row not in self.header_mismatch_row:
            self.header_mismatch_row.append(row)


    def report_header_missmatch(self, invoice_file, row):
        invoice_file = Path(invoice_file).name
        self.header_mismatch_num += 1
        if invoice_file not in self.header_mismatch_file:
            self.header_mismatch_file.append(invoice_file)
        if row not in self.header_mismatch_row:
            self.header_mismatch_row.append(row)


    def report_missing_addresses(self, missing_addresses, missing_rows):
        for address in missing_addresses:
            self.missing_address_num += 1
            if address not in self.missing_address:
                self.missing_address.append(address)
            if missing_rows not in self.missing_address_rows:
                self.missing_address_rows = missing_rows


    def report_financial(self, sheet, starting_row, headers, date_headers, float_headers, currency_headers):
        if sheet.max_row <= starting_row:
            raise ValueError("No valid invoice data was available for reporting.")

        date_col = headers.index(date_headers[0]) + 1
        amount_col = headers.index(float_headers[0]) + 1
        amount_prise_col = headers.index(currency_headers[0]) + 1
        total_sum_col = headers.index(currency_headers[1]) + 1

        data_starting_row = starting_row + 1
        avgn = 0

        unit_prices = []
        quantities = []
        total_prices = []

        for row in range(data_starting_row, sheet.max_row + 1):
            quantity = sheet.cell(row=row, column=amount_col).value
            unit_price = sheet.cell(row=row, column=amount_prise_col).value
            total_price = sheet.cell(row=row, column=total_sum_col).value

            if quantity is not None:
                quantity = float(quantity)
                quantities.append(quantity)

            if unit_price is not None:
                unit_price = float(unit_price)
                unit_prices.append(unit_price)

            if total_price is not None:
                total_price = float(total_price)
                total_prices.append(total_price)

        self.amount_sum = sum(quantities)
        self.total_sum_prise = sum(total_prices)
        self.amount_prise_average = (sum(unit_prices) / len(unit_prices) if unit_prices else 0)


    def report_travel(self, sheet, starting_row, headers, date_headers, text_headers, distcolumn, distances):
        self.home_customer_dist, self.customer_customer_dist, self.customer_home_dist = distances

        date_col = headers.index(date_headers[0]) + 1
        customer_col = headers.index(text_headers[0]) + 1
        dist_col = distcolumn

        unique_days = []
        unique_customers = []

        data_starting_row = starting_row + 1

        for row in range(data_starting_row, sheet.max_row + 1):
            value = sheet.cell(row=row, column=date_col).value
            if value not in unique_days:
                unique_days.append(value)

            value = sheet.cell(row=row, column=customer_col).value
            if value not in unique_customers:
                unique_customers.append(value)

            if sheet.cell(row=row, column=dist_col).value != None:
                self.total_distance += sheet.cell(row=row, column=dist_col).value
            if sheet.cell(row=row, column=dist_col+1).value != None:
                self.total_distance += sheet.cell(row=row, column=dist_col+1).value

        self.work_days = len(unique_days)
        self.customers_visited = len(unique_customers)


    def write_report(self):
        now = datetime.now()

        return f"""
Rechnungen Zusammenfassung
Verarbeitungsdatum: {self.time:%d.%m.%Y}
Rechnungen verarbeitet: {self.invoice_number}
FINANZEN
----------------------------------------------------------------------------------------
Menge insgesamt: {self.amount_sum:,.2f}
Einzelpreis durchschnitt: {self.amount_prise_average:,.2f}
Gesamtpreis: {self.total_sum_prise:,.2f}
 
ARBEITSWEG
----------------------------------------------------------------------------------------
Arbeitstage: {self.work_days}
Mandanten besucht: {self.customers_visited}
Distanzen insgesamt: {self.total_distance:.2f} km
 
von zu Hause -> Mandant: {self.home_customer_dist:.2f} km
Mandant -> Mandant: {self.customer_customer_dist:.2f} km
Mandant -> Nach Hause: {self.customer_home_dist:.2f} km
 
VALIDIERUNGSFEHLER
----------------------------------------------------------------------------------------
Anzahl fehlender Rechnungen: {self.missing_invoice_number}
Fehlende Rechnungen: {", ".join(str(invoice) for invoice in self.missing_invoices) if self.missing_invoices else "Keine"}
 
Anzahl der Header-Nichtübereinstimmungen: {self.header_mismatch_num}
Rechnungen mit Header-Nichtübereinstimmungen: {", ".join(str(header) for header in self.header_mismatch_file) if self.header_mismatch_file else "Keine"}
Tabellenzeilen mit Header-Nichtübereinstimmungen: {", ".join(str(header) for header in self.header_mismatch_row) if self.header_mismatch_row else "Keine"}
 
Anzahl fehlender Adressen: {self.missing_address_num}
Fehlende Adressen: {", ".join(str(address) for address in self.missing_address) if self.missing_address else "Keine"}
Tabellenzeilen mit fehlenden Adressen: {", ".join(str(row) for row in self.missing_address_rows) if self.missing_address_rows else "Keine"}
 
Hinweis: Fehlende Rechnungen und Adressen sollen manuel korrigiert werden. Fehlende Adressen und Distanzen können mit "Existing Excel file" korrigiert werden. Falls es Validierungsfehler gibt, sind alle andere Dateien falsch!
""".strip()


    def save(self, filename=None, folder=None):

        if filename == None:
            filename=f"Invoice Report {self.time:%d-%m-%Y} {self.time:%H-%M-%S}.pdf"

        if folder == None:
            folder = self.report_folder

        Path(folder).mkdir(parents=True, exist_ok=True)

        output_path = f"{folder}/{filename}"

        text = self.write_report()

        # Will work only with windows!!!
        pdfmetrics.registerFont(TTFont("Calibri", r"C:\Windows\Fonts\calibri.ttf"))
        pdfmetrics.registerFont(TTFont("Calibri-Bold", r"C:\Windows\Fonts\calibrib.ttf"))

        c = canvas.Canvas(output_path, pagesize=A4)

        width, height = A4
        margin = 20 * mm
        y = height - margin
        line_height = 7 * mm
        max_width = width - 2 * margin

        for line in text.splitlines():
            if not line.strip():
                y -= line_height  # extra space for paragraph
                continue

            # Choose font
            if line == "Rechnungen Zusammenfassung":
                font = "Calibri-Bold"
                font_size = 18
                centered = True

            elif line in {
                "HAUPTINFORMATION",
                "FINANZEN",
                "ARBEITSWEG",
                "VALIDIERUNGSFEHLER",
            }:
                font = "Calibri-Bold"
                font_size = 14
                centered = False

            else:
                font = "Calibri"
                font_size = 14
                centered = False

            c.setFont(font, font_size)

            # Split line into words
            words = line.split()
            current_line = ""

            for word in words:
                test_line = f"{current_line} {word}".strip()

                if stringWidth(test_line, font, font_size) <= max_width:
                    current_line = test_line
                else:
                    # Draw current line before starting a new one
                    if y < margin:
                        c.showPage()
                        y = height - margin
                        c.setFont(font, font_size)

                    if centered:
                        c.drawCentredString(width / 2, y, current_line)
                    else:
                        c.drawString(margin, y, current_line)

                    y -= line_height
                    current_line = word

            # Draw remaining text
            if current_line:
                # New page if necessary
                if y < margin:
                    c.showPage()
                    y = height - margin
                    c.setFont(font, font_size)

                if centered:
                    c.drawCentredString(width / 2, y, current_line)
                else:
                    c.drawString(margin, y, current_line)

                y -= line_height
            # # New page if necessary
            # if y < margin:
            #     c.showPage()
            #     y = height - margin

            # # Titles / section headers
            # if line == "Rechnungen Zusammenfassung":
            #     c.setFont("Calibri-Bold", 18)
            #     c.drawCentredString(width / 2, y, line)

            # elif line in {
            #     "Rechnungen Zusammenfassung",
            #     "HAUPTINFORMATION",
            #     "FINANZEN",
            #     "ARBEITSWEG",
            #     "VALIDIERUNGSFEHLER",
            # }:
            #     c.setFont("Calibri-Bold", 14)
            #     c.drawString(margin, y, line)

            # # Normal text
            # else:
            #     c.setFont("Calibri", 14)
            #     c.drawString(margin, y, line)

            # # Move to next line
            # y -= line_height

        c.save()
        return output_path