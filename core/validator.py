import re
from pathlib import Path


def is_coordinates_valid(value):
    pattern = r"-?\d+(\.\d+)?, -?\d+(\.\d+)?"
    return re.fullmatch(pattern, value.strip()) is not None


def validate_invoice_file_number(invoice_files, folder_files):
    """Return PDF filenames missing from the supplied invoice list."""
    directory = Path(folder_files)

    pdf_files = {
        path.name
        for path in directory.glob("*.pdf")
    }
    provided_files = {
        Path(path).name
        for path in invoice_files
    }

    if not provided_files.issubset(pdf_files):
        raise ValueError(
            "The invoice list contains files "
            "that are not present in the directory."
        )

    return sorted(pdf_files - provided_files)


def validate_headers(data, headers, invoice_file):
    """
    Verify that an invoice contains the same fields as the
    previously processed invoices.

    All invoices must have a consistent structure so that their
    data can safely be written into the same Excel columns.

    Raises:
        ValueError: If the invoice headers do not match.
    """
    if set(data.keys()) != set(headers):
        return False
        # raise ValueError(f"Headers in {invoice_file} do not match!")
    else:
        return True


def validate_addresses(sheet, addresscolumn, starting_row, database):
    missing = []
    rows = []

    for i in range(starting_row, sheet.max_row+1):
        if sheet.cell(row=i, column=addresscolumn).value not in database.data:
            if sheet.cell(row=i, column=addresscolumn).value not in missing:
                missing.append(sheet.cell(row=i, column=addresscolumn).value)
            rows.append(i)

    return missing, rows