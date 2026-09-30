import re
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from excel.scanner import extract_text
from excel.scanner import parse_invoice
from core.validator import *
from core.calculator import calculate_distances
from core.address_database import AddressDatabase


INVOICE_PATTERN = re.compile(
    r"^invoice\s+(\d+)\.pdf$",
    re.IGNORECASE,
)


def sort_invoices(path):
    """Return invoice PDFs sorted by their numeric invoice number."""
    folder = Path(path)
    invoices = []

    for file_path in folder.glob("*.pdf"):
        match = INVOICE_PATTERN.fullmatch(file_path.name)

        if match:
            invoice_number = int(match.group(1))
            invoices.append((invoice_number, file_path))

    if not invoices:
        raise ValueError("No valid invoice PDFs were found.")

    return [
        file_path
        for _, file_path in sorted(invoices)
    ]


def create_workbook():
    """
    Create and initialize the Excel workbook used for the invoice report.

    Returns:
        tuple: A tuple containing the workbook and the active worksheet.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices Report"

    return wb, ws


def write_headers(ws, headers, starting_row):
    """
    Write the invoice field names as column headers in the worksheet.

    Args:
        ws: The worksheet where the headers will be written.
        headers: List of column header names.
        starting_row: Excel row where the headers should be written.
    """
    for col, header in enumerate(headers, start=1):
        ws.cell(row=starting_row, column=col, value=header)


def convert_value(header, value, int_headers,date_headers, float_headers, currency_headers):
    """
    Convert an invoice value into the appropriate Python data type.

    Conversion is based on the header classification returned by
    the invoice parser.

    Supported conversions: Integer, Date, Float, Currency

    Args:
        header: Name of the invoice field.
        value: Raw value extracted from the invoice.
        int_headers: Headers containing integer values.
        date_headers: Headers containing date values.
        float_headers: Headers containing floating-point values.
        currency_headers: Headers containing currency values.

    Returns:
        The converted value, or None if the input value is None.
    """
    if value is None:
        return value

    if header in int_headers:
        try:
            return int(value)
        except Exception:
            return False

    if header in date_headers:
        try:
            return datetime.strptime(str(value), "%d.%m.%Y").date()
        except Exception:
            return False

    if header in float_headers:
        try:
            return float(value.replace(",", "."))
        except Exception:
            return False

    if header in currency_headers:
        try:
            return float(str(value).replace("€", "").replace(".", "").replace(",", ".").strip())
        except Exception:
            return False

    return value


def apply_number_format(cell, header, int_headers, date_headers, float_headers, currency_headers):
    """
    Apply the appropriate Excel number format to a worksheet cell.

    The format is selected according to the field type identified
    by the invoice parser.
    """
    if header in int_headers:
        cell.number_format = "0"

    if header in date_headers:
        cell.number_format = "dd/mm/yyyy"

    if header in float_headers:
        cell.number_format = "0.00"

    if header in currency_headers:
        cell.number_format = "#,##0.00 €"


def write_invoice_data(ws, invoice_file, data, headers, row, int_headers, date_headers, float_headers, currency_headers, reporter):
    """
    Write one invoice's data into a worksheet row.

    Each value is converted to the correct Python type before being
    written to Excel. The corresponding Excel number format is then
    applied to the cell.
    """
    for col, header in enumerate(headers, start=1):

        value = data.get(header)

        # Convert the raw invoice value to the appropriate Python type
        value = convert_value(header, value, int_headers, date_headers, float_headers, currency_headers)

        if value == False:
            reporter.report_header_missmatch(invoice_file, row)
            continue

        cell = ws.cell(row=row, column=col, value=value)

        # Apply the corresponding Excel display format
        apply_number_format(cell, header, int_headers, date_headers, float_headers, currency_headers)


def process_invoice(ws, invoice_file, headers, row, int_headers, date_headers, float_headers, currency_headers, reporter):
    """
    Extract, parse, validate, and write a single invoice.

    This function represents the complete processing pipeline for
    one invoice after the workbook structure has been established.
    """
    # Extract the raw text from the PDF invoice
    invoice_text = extract_text(invoice_file)

    # Parse the invoice text into structured data
    parse_dictionary = parse_invoice(invoice_text)

    data = parse_dictionary[0]

    # Ensure that this invoice has the same structure as previous invoices
    validator = validate_headers(data, headers, invoice_file)
    if validator == False:
        reporter.report_header_missmatch(invoice_file, row)
        return False

    # Write the parsed invoice data into the worksheet
    write_invoice_data(ws, invoice_file, data, headers, row, int_headers, date_headers, float_headers, currency_headers, reporter)

    return True


def load_invoice_metadata(invoice_file):
    """
    Parse an invoice and extract its metadata definitions.

    The first invoice establishes the structure of the Excel report.
    Its headers and field type classifications are used for all
    subsequent invoices.

    Returns:
        tuple:
            headers,
            int_headers,
            date_headers,
            float_headers,
            currency_headers,
            text_headers
    """
    invoice_text = extract_text(invoice_file)
    parse_dictionary = parse_invoice(invoice_text)

    data = parse_dictionary[0]

     # The first invoice defines the structure of the report
    headers = list(data.keys())

    int_headers = parse_dictionary[1]
    date_headers = parse_dictionary[2]
    float_headers = parse_dictionary[3]
    currency_headers = parse_dictionary[4]
    text_headers = parse_dictionary[5]

    return headers, int_headers, date_headers, float_headers, currency_headers, text_headers


def write_invoices(ws, invoice_files, starting_row, reporter):
    """
    Process all invoice files and write their data to the worksheet.

    The first invoice determines the column structure and field types.
    Every subsequent invoice is validated against that structure.

    Returns:
        tuple:
            headers,
            date_headers,
            text_headers
    """
    first_invoice = True
    data_starting_row = starting_row + 1

    headers = None
    int_headers = None
    date_headers = None
    float_headers = None
    currency_headers = None
    text_headers = None

    for invoice_file in invoice_files:

        if first_invoice:
            # Use the first invoice to establish the report structure
            headers, int_headers, date_headers, float_headers, currency_headers, text_headers = load_invoice_metadata(invoice_file)

            write_headers(ws, headers, starting_row)

            first_invoice = False

        # Process the current invoice using the established structure
        flag = process_invoice(ws, invoice_file, headers, data_starting_row, int_headers, date_headers, float_headers, currency_headers, reporter)

        if flag == True:
            data_starting_row += 1

    return headers, int_headers, date_headers, float_headers, currency_headers, text_headers


def add_distance_columns(ws, starting_row):
    """
    Add the distance calculation columns to the invoice report.

    Two columns are added:
        - Distance 1 (km)
        - Distance 2 (km)

    Returns:
        int: The column number of the first distance column.
    """
    last_col = ws.max_column

    ws.cell(row=starting_row, column=last_col + 1, value="Distance 1 (km)")

    ws.cell(row=starting_row, column=last_col + 2, value="Distance 2 (km)")

    return last_col + 1


def calculate_invoice_distances(database_file, ws, headers, date_headers, text_headers, distcolumn, starting_row, reporter):
    """
    Calculate and populate invoice distances using the address database.

    The address and invoice date columns are identified from the
    parsed invoice metadata rather than being hard-coded.
    """
    # Load the address database used by the distance calculator
    database = AddressDatabase(database_file)
    database.load()

    addresscolumn = headers.index(text_headers[1]) + 1
    datecolumn = headers.index(date_headers[0]) + 1

    database.load()

    missing_addresses, missing_rows = validate_addresses(ws, addresscolumn, starting_row +1, database)

    reporter.report_missing_addresses(missing_addresses, missing_rows)

    data = calculate_distances(ws, ws.max_row, addresscolumn, datecolumn, distcolumn, database.get_data(), first_data_row=starting_row + 1, missing_addresses=missing_addresses)

    # data: [home_customer_distance, customer_customer_dist, customer_home_dist]
    return data


def autofit_columns(ws):
    """
    Automatically adjust worksheet column widths based on cell contents.

    The width is calculated from the longest value found in each column,
    with a small amount of additional padding for readability.
    """
    for column in ws.columns:

        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)) + 1)

        ws.column_dimensions[column_letter].width = (max_length + 2)


def generate_output_filename(report_folder, time):
    """
    Generate a timestamped filename for the invoice report.

    Returns:
        str: Output path including the current date and time.
    """

    return f"{report_folder}/Invoice Report {time:%d-%m-%Y} {time:%H-%M-%S}.xlsx"


def save_workbook(wb, report_folder, time):
    """
    Save the generated workbook using a timestamped output filename.

    Returns:
        str: Path of the saved Excel file.
    """
    output_file = generate_output_filename(report_folder, time)
    wb.save(output_file)

    return output_file


def write_workbook(invoice_folder, database_file, report_folder, starting_row, reporter, time):
    """
    Generate the complete invoice Excel spreadsheet.

    This function acts as the main orchestration layer. It coordinates
    workbook creation, invoice processing, distance calculations,
    column formatting, and file output.

    Args:
        invoice_folder: invoice folder path
        database_file: data base file path
        report_folder: folder where to save Excel workbook and report
        starting_row: Excel row where the report headers should begin.
        reporter: Reporter class instance to constract summary and bug reports
    Returns:
        str: Path of the generated Excel report.
    """
    wb, ws = create_workbook()

    invoice_files = sort_invoices(invoice_folder)

    reporter.report_invoice_files(invoice_files, invoice_folder)

    headers, int_headers, date_headers, float_headers, currency_headers, text_headers = write_invoices(ws, invoice_files, starting_row, reporter)
    # Invoice-level validation errors are recorded by the reporter.
    # Only successfully validated invoices advance the output row.

    distcolumn = add_distance_columns(ws, starting_row)

    distances = calculate_invoice_distances(database_file, ws, headers, date_headers, text_headers, distcolumn, starting_row, reporter)
    # reporter and validator work inside calculate_invoice_distances(...)
    # distances: [home_customer_distance, customer_customer_dist, customer_home_dist]

    reporter.report_financial(ws, starting_row, headers, date_headers, float_headers, currency_headers)
    reporter.report_travel(ws, starting_row, headers, date_headers, text_headers, distcolumn, distances)
    # Aggregate the financial and travel metrics from the populated
    # worksheet and the distance-calculation results.

    autofit_columns(ws)

    return save_workbook(wb, report_folder, time)