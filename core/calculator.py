from geopy.distance import geodesic


def check_dates(sheet, sheet_row, date_column):
    """
    Compares dates between the current row and the next one.
    """
    if sheet.cell(row=sheet_row, column=date_column).value == sheet.cell(
            row=sheet_row + 1, column=date_column).value:
        return True
    else:
        return False


def calculate_address_home(sheet, data, address_column, output_row, output_column, input_row):
    """
    Distance calculation between addresses of a row and the home address.
    """
    distance = round(
        geodesic(
            data["home"],
            data[sheet.cell(row=input_row, column=address_column).value],
        ).kilometers,
        2,
    )

    sheet.cell(row=output_row, column=output_column).value = distance

    return distance


def calculate_address_address(sheet, data, address_column, output_row, output_column, input_row_1, input_row_2):
    """
    Distance calculation between two neighboring addresses.
    """
    distance = round(
        geodesic(
            data[sheet.cell(row=input_row_1, column=address_column).value],
            data[sheet.cell(row=input_row_2, column=address_column).value],
        ).kilometers,
        2,
    )

    sheet.cell(row=output_row, column=output_column).value = distance

    return distance


def calculate_distances(sheet, maxrow, address_column, date_column, distance_column, data, first_data_row = 3, missing_addresses=None):
    """
    The geo-distance calculation algorithm between latitudes and longtitudes of two addresses. 
    The initial calculations are provided for the first and the last working day from home to 
    the coresponding address. The main cycle checks if the next address relate to the same 
    working day. The distance will be calculated between two addresses if positive. The route 
    home from the last address of the current working day and the first address of the next 
    working day calculates otherwise.
    """
    # Distances are calculated using geodesic distance between coordinates. They represent straight-line distances, not
    # road or walking distances. Missing addresses are skipped so that one unresolved location does not prevent the 
    # remaining distances from being calculated. This means the reported total may be incomplete.

    if sheet.max_row <= first_data_row:
        raise ValueError("No valid invoice data was available for reporting.")

    if missing_addresses is None:
        missing_addresses = []

    home_customer_distance = 0
    customer_customer_dist = 0
    customer_home_dist = 0

    # initial calcculations for the first and the last row:
    if sheet.cell(row=first_data_row, column=address_column).value not in missing_addresses:
        distance = calculate_address_home(sheet, data, address_column, first_data_row, distance_column, first_data_row)
        home_customer_distance += distance

    if sheet.cell(row=maxrow, column=address_column).value not in missing_addresses:
        distance = calculate_address_home(sheet, data, address_column, maxrow, distance_column + 1, maxrow)
        customer_home_dist += distance

    # main cycle
    for i in range(first_data_row, maxrow):

        if check_dates(sheet, i, date_column) == True:

            if sheet.cell(row=i + 1, column=address_column).value not in missing_addresses:
                distance = calculate_address_address(sheet, data, address_column, i + 1, distance_column, i, i + 1)
                customer_customer_dist += distance

        else:

            if sheet.cell(row=i, column=address_column).value not in missing_addresses:
                distance = calculate_address_home(sheet, data, address_column, i, distance_column + 1, i)
                customer_home_dist += distance

            if sheet.cell(row=i + 1, column=address_column).value not in missing_addresses:
                distance = calculate_address_home(sheet, data, address_column, i + 1, distance_column, i + 1)
                home_customer_distance += distance

    return home_customer_distance, customer_customer_dist, customer_home_dist