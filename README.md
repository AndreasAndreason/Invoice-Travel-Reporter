# Invoice & Travel Report Automation

A Python desktop application for automating invoice data extraction, Excel reporting, and business travel distance calculations.

The application processes PDF invoices, extracts structured customer and financial information, consolidates the results into an Excel workbook, and calculates straight-line distances between customer locations and a configured home address. It also generates a PDF summary containing financial, travel, and validation statistics.

## Overview

Small businesses that process recurring invoices may need to consolidate invoice data and prepare annual financial and business travel summaries.

This project automates parts of that workflow by combining PDF text extraction, structured data validation, Excel report generation, address-to-coordinate mapping, and geodesic distance calculations in a desktop application.

The application is designed around a recurring invoice workflow with a relatively stable customer base and a consistent invoice template.

## Features

### Invoice processing

- Batch processing of PDF invoices from a selected directory.
- Extraction of invoice data using PyMuPDF.
- Parsing of customer information, invoice numbers, dates, quantities, and prices.
- Validation of invoice fields against a consistent report structure.
- Reporting of missing invoice files and invoice header mismatches.

### Excel reporting

- Automated Excel workbook generation using openpyxl.
- Conversion of extracted values into appropriate data types.
- Formatting of dates, quantities, and currency values.
- Automatic adjustment of worksheet column widths.
- Support for selecting an existing workbook for manual address database fulfillment and distance calculations.

### Address management

- JSON-based address-to-coordinate database.
- Detection of addresses missing from the database.
- Manual entry of coordinates through the desktop interface.
- Reuse of previously stored coordinates.

### Travel distance calculations

- Geodesic distance calculations between customer locations.
- Calculation of home-to-customer, customer-to-customer, and customer-to-home distances.
- Grouping of travel calculations by invoice date.
- Skipping of distance calculations involving unresolved addresses.

### Reporting

- Excel report containing the processed invoice data and calculated distances.
- PDF summary containing financial, travel, and validation statistics.
- Reporting of missing invoices, inconsistent invoice structures, and unresolved addresses.

## Technology Stack

| Technology    | Purpose                              
| ------------- | -------------------------------------
| Python        | Application logic and data processing
| PyMuPDF       | PDF text extraction                  
| openpyxl      | Excel workbook processing            
| Tkinter / ttk | Desktop graphical interface          
| ReportLab     | PDF report generation                
| geopy         | Geodesic distance calculations       
| JSON          | Address-to-coordinate storage        

## Application Workflow

1. Select a folder containing invoice PDFs or open an existing Excel workbook.
2. Extract text from the supported invoice documents.
3. Parse and validate the invoice data.
4. Consolidate valid invoice records into an Excel workbook.
5. Check customer addresses against the local coordinate database.
6. Enter coordinates for unresolved addresses when required.
7. Calculate straight-line distances based on invoice dates and customer locations.
8. Generate the Excel report and PDF summary.

## Project Structure

```text
invoice-travel-reporter/
    core/
        address_database.py
        calculator.py
        reporter.py
        validator.py
    data/
        invoices/
        addresses.json
    excel/
        scanner.py
        writer.py
    gui/
        dialog.py
        main_window.py
    .gitattributes
    .gitignore
    LICENSE.txt
    main.py
    README.md
    requirements.txt
```

*The structure above represents the intended organization; adjust the paths to match the actual repository.*

## Installation

### Requirements

- Python 3.14.
- A supported desktop environment with Tkinter.
- PDF invoices following the supported document template.

### 1. Clone the repository
git clone https://github.com/AndreasAndreason/Invoice-Travel-Reporter.git

cd Invoice-Travel-Reporter
### 2. Create a virtual environment

Windows:

python -m venv .venv
.venv\Scripts\activate

Linux / macOS:

python3 -m venv .venv
source .venv/bin/activate
### 3. Install dependencies
pip install -r requirements.txt
### 4. Configure the application

Prepare the local data directory and address database.

The address database uses a JSON mapping between address strings and coordinate pairs. Configure the home coordinates and customer coordinates according to the format expected by the application.

Do not use real customer data when experimenting with the project.

## Usage

Start the desktop application using the project's entry point.

python main.py

The application provides two workflows:

**Invoice folder processing:** Select an invoice directory, an address database, and an output directory or confirm to apply default settings. The application processes the invoices and generates the reports.

**Existing Excel workbook:** Select a workbook, choose the worksheet and relevant columns, provide any missing coordinates, and calculate travel distances.

*Update the startup command if the application is packaged or launched through a different entry point.*

## Input Data

### Invoice PDFs

The current parser is designed for a specific invoice layout. It relies on expected section labels, field ordering, and the arrangement of extracted PDF text.

Invoices using a different template may require changes to the parsing logic.

### Address database

The address database maps address strings to coordinate pairs.

Example structure:

```text
{
    "home": "52.5200, 13.4050",
    "Example Street 10, 10115 Berlin": "52.5321, 13.3849",
    "Example Avenue 5, 10117 Berlin": "52.5170, 13.3889"    
}
```

The coordinates above are illustrative. Replace them with valid coordinates for the relevant locations.

Address matching currently relies on exact string matching. Differences in whitespace, punctuation, or spelling may result in an address being treated as missing.

## Distance Methodology

Distances are calculated using geodesic distance between geographic coordinates.

These values represent straight-line distances, not road, cycling, or walking distances.

The total distance may be incomplete when an address is missing or cannot be resolved. The calculations should therefore be interpreted in the context of the input data and the application's validation results.

## Privacy and Data Handling

Invoice PDFs, customer addresses, coordinates, and generated reports may contain sensitive business or personal information.

- Keep real invoices and customer-specific databases out of version control.
- Use synthetic data for development, testing, and demonstrations.
- Keep local business data separate from application code.
- Review generated reports before sharing them.
- Do not describe the address database as anonymous merely because it does not contain customer IDs.

The address database is intentionally implemented as a simple address-to-coordinate mapping for a small, relatively stable customer base. This design is not intended as a scalable customer management system.

## Known Limitations

- The invoice parser depends on a specific invoice template.
- Address extraction relies on heuristic recognition of German street names, house numbers, and postal codes.
- Address matching is sensitive to text formatting.
- The current distance calculations use straight-line distances rather than actual travel routes.
- Missing addresses can lead to incomplete distance totals.
- Financial calculations depend on the correctness and consistency of the extracted invoice values.
- The desktop interface and font configuration may require platform-specific adjustments.


## Future Improvements

- More robust parsing of different invoice layouts.
- Stronger validation and normalization of addresses.
- Improved handling of missing or malformed invoice fields.
- More reliable monetary calculations using decimal arithmetic.
- Automated tests and continuous integration.
- Cross-platform font handling and application packaging.
- Improved error reporting and recovery in the desktop interface.
- Eliminate the TTF Windows font dependency.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
