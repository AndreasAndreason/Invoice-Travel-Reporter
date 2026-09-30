import json
from pathlib import Path


class AddressDatabase: # needs improvement
    """
    The address database maps normalized address strings to coordinates.

    Design rationale:
    - The database contains no customer IDs or other customer-specific
      identifiers.
    - The simple key-value structure is sufficient for the project's
      small, relatively stable customer base.
    - Exact string matching is intentional in the current implementation.

    Limitations:
    - Address matching is sensitive to spelling, whitespace, and
      punctuation.
    - The database is not designed for large-scale customer lookup.
    """

    def __init__(self, filename):
        self.filename = filename
        self.data = {}

    def load(self):
        with open(self.filename, "r", encoding="utf-8") as file:
            self.data = json.load(file)
        return self.data

    def add_missing_addresses(self, sheet, address_column, starting_row, maxrow, get_coordinates):
        # Database-filling algorithm.
        for i in range(starting_row + 1, maxrow):
            value = sheet.cell(row=i, column=address_column).value

            if value is None or not str(value).strip():
                continue

            address = str(value).strip()

            if address not in self.data:
                newcords = get_coordinates(sheet.cell(row=i, column=address_column).value)

                if newcords == None or newcords == "":
                    continue
                    # raise Exception("Incorrect coordinates!")
                self.data.update({address: newcords})

    def save(self):
        with open(self.filename, "w", encoding="utf-8") as file:
            json.dump(self.data, file, ensure_ascii=False, indent=4)

    def get_data(self):
        return self.data