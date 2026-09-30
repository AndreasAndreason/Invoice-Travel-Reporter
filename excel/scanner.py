import pymupdf as pmp
import re

# The scan alg was designed for a certain invoice document structure. It needs to be
# changed for different types of documents

"""
Data structure:

KUNDE       KUNDENNUMMER        RECHNUNGSNUMMER     RECHNUNGSDATUM      LIEFERDATUM

multiple    int                 int                 date                date
lines!

MENGE       EINZELPREIS         GESAMTPREIS
float       currency            currency

GESAMTBETRAG BRUTTO     currency
MWST                    currency
GESAMTBETRAG NETTO      currency
"""

METADATA = [ # to count and filter until KUNDE data with multiple lines
    "KUNDENNUMMER",
    "RECHNUNGSNUMMER",
    "RECHNUNGSDATUM",
    "LIEFERDATUM",
]

TABLE_2 = [
    "MENGE",
    "EINZELPREIS",
    "GESAMTPREIS",
]

TABLE_VERTICAL = [
    "GESAMTBETRAG BRUTTO",
    "MWST",
    "GESAMTBETRAG NETTO",
]

headers_for_report = [
    ["Rechnungsnummer", "Kundennummer"],
    ["Lieferdatum"],
    ["Menge"],
    ["Einzelpreis", "Gesamtpreis"],
    ["Kundenname", "Kundenadresse"],
]


# Typical German street types
STREET_TYPES = re.compile(
    r"(?:"
    r"straße|strasse|str\.|"
    r"allee|chaussee|"
    r"weg|platz|ring|gasse|"
    r"damm|ufer|steig|hof|"
    r"höhe|markt|brücke|"
    r"promenade"
    r")(?=\s|$)",
    re.IGNORECASE
)

# German postal code + city
POSTAL_CODE = re.compile(
    r"\b\d{5}\s+[A-Za-zÄÖÜäöüß]"
)

# House number, e.g.
# 17
# 17a
# 17-19
# 17a-19b
HOUSE_NUMBER = re.compile(
    r"\b\d+[a-zA-Z]?(?:-\d+[a-zA-Z]?)?\b"
)


def extract_text(pdf_file):
    with pmp.open(pdf_file) as doc:
        return "\n".join(page.get_text() for page in doc) # type: ignore


def reconstruct_customer(lines):
    """
    Reconstruct lines into a logical string.
    A line ending in '-' is joined directly to the next line:
    """
    result = ""

    for line in lines:
        line = line.strip()

        if not line:
            continue

        if not result:
            result = line

        elif result.endswith("-"):
            # Hyphen is part of the word
            result += line

        else:
            result += " " + line

    return result


def normalize_text(text):
    # PDF extraction sometimes puts a space before ß ("Stra ße" -> "Straße")
    text = re.sub(r"\s+ß", "ß", text)
    # Normalize non-breaking spaces
    text = text.replace("\u00a0", " ")
    return text


def split_name_address(customer):
    """
    Split customer information into name and address.
    """
    customer_string = reconstruct_customer(customer)

    # Find postal code
    postal_match = POSTAL_CODE.search(customer_string)

    if not postal_match:
        return customer_string, None

    # Everything before postal code
    before_postal = customer_string[:postal_match.start()].strip()

    # Find house number before postal code
    house_matches = list(HOUSE_NUMBER.finditer(before_postal))

    if not house_matches:
        return before_postal, customer_string[postal_match.start():].strip()

    # Normally the last number before postal code is the house number
    house_match = house_matches[-1]

    # Street must contain a known street type
    street_type_matches = list(
        STREET_TYPES.finditer(
            before_postal[:house_match.end()]
        )
    )

    if not street_type_matches:
        return before_postal, customer_string[postal_match.start():].strip()

    # We need to find where the street name starts. For now, use the word immediately
    # before the street type and include hyphenated words as part of the street.
    
    tokens = before_postal.split()

    # Find token containing the street type
    street_token_index = next(
        (
            i for i, token in enumerate(tokens)
            if STREET_TYPES.search(token)
        ),
        None
    )

    if street_token_index is None:
        return before_postal, customer_string[postal_match.start():].strip()

    # For abbreviated/explicit street types such as "Str." and
    # "Straße", the street starts at the street-type token.
    #
    # For types such as "Allee", "Weg", "Platz", etc., the word
    # before it is usually part of the street name.
    street_type_token = tokens[street_token_index].lower()

    street_type_token = tokens[street_token_index].lower()

    # True for "Damm", "Allee", "Str." ... False for "Stubenrauchstraße"
    is_bare_type = STREET_TYPES.fullmatch(street_type_token) is not None

    if is_bare_type and street_type_token not in {"str.", "straße", "strasse"}:
        # "Blumberger Damm": the word before is part of the street name
        street_start = max(1, street_token_index - 1)
    else:
        # "Stubenrauchstraße" (compound) or "Str. des 17. Juni"
        street_start = street_token_index

    name = " ".join(tokens[:street_start]).strip()

    address_before_postal = " ".join(
        tokens[street_start:]
    ).strip()

    address_after_postal = customer_string[
        postal_match.start():
    ].strip()

    address = f"{address_before_postal} {address_after_postal}"

    return name, address


def parse_invoice(text):

    text = normalize_text(text)

    # ignore unnecessary text
    filtered = re.search(r"RECHNUNG.*?(?=Zahlungsinformation:)", text, re.DOTALL)
    if filtered:
        text = filtered.group(0)

    # Remove empty rows
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    # Locate the customer section. The parser relies on the invoice
    # template's fixed section order, so fail clearly if it is missing.
    try:
        customer_index = lines.index("KUNDE")
    except ValueError as exc:
        raise ValueError("Invoice does not contain the expected KUNDE section.") from exc

    # Customer data starts after KUNDE + metadata headers
    customer_start = customer_index + 1 + len(METADATA)

    # In the supported invoice template, the first integer-only line
    # after the customer section marks the beginning of the invoice metadata.
    customer_end = customer_start

    customer_end = next(
        (
            i for i in range(customer_start, len(lines))
            if re.fullmatch(r"\d+", lines[i])
        ),
        len(lines)
    )

    customer = lines[customer_start:customer_end]
    name, address = split_name_address(customer)

    if address:
        address = address.replace("Deutschland", "").strip()

    # KUNDENNUMMER RECHNUNGSNUMMER RECHNUNGSDATUM LIEFERDATUM
    metadata_end = customer_end + len(METADATA)
    metadata_values = lines[customer_end:metadata_end]
    metadata = dict(zip(METADATA, metadata_values))

    # MENGE EINZELPREIS GESAMTPREIS
    table_2_start = metadata_end + len(TABLE_2)
    table_2_end = table_2_start + len(TABLE_2)
    table_2_values = lines[table_2_start:table_2_end]
    table_2 = dict(zip(TABLE_2, table_2_values))

    # GESAMTBETRAG BRUTTO   MWST    GESAMTBETRAG NETTO
    table_vertical_values = lines[table_2_end:table_2_end + len(TABLE_VERTICAL) * 2]
    table_vertical = dict(zip(TABLE_VERTICAL, table_vertical_values[1::2]))

    """
    Need to select the necessary headers. Here now are only
    necessary positions. Compare with headers_for_report!!!
    """
    dictionary = {
        "Rechnungsnummer": metadata["RECHNUNGSNUMMER"],
        "Lieferdatum": metadata["LIEFERDATUM"],
        "Kundennummer": metadata["KUNDENNUMMER"],
        "Menge": table_2["MENGE"],
        "Einzelpreis": table_2["EINZELPREIS"],
        "Gesamtpreis": table_2["GESAMTPREIS"],
        "Kundenname": name,
        "Kundenadresse": address,
    }

    int_headers = headers_for_report[0]

    date_headers = headers_for_report[1]

    float_headers = headers_for_report[2]

    currency_headers = headers_for_report[3]

    text_headers = headers_for_report[4]

    return dictionary, int_headers, date_headers, float_headers, currency_headers, text_headers


def header_def_distionary():
    date_headers = headers_for_report[1]
    float_headers = headers_for_report[2]
    currency_headers = headers_for_report[3]
    text_headers = headers_for_report[4]

    return date_headers, float_headers, currency_headers, text_headers