import re
from excel_legacy_utils import get_xls_cell_value,update_xls_cell,get_xls_last_row,get_xls_cell_reference_by_value

def get_column_values(file_path: str, sheet_name: str, column_name: str) -> list:
    """
    Extracts all values from a specified column in a legacy .xls file.

    This function reads an Excel sheet to find a specific column by its header name,
    then collects all the data in that column from the row after the header to the last
    row containing data in the sheet.

    Args:
        file_path (str): The full path to the .xls file.
        sheet_name (str): The name of the sheet to read from (e.g., "Mass Update").
        column_name (str): The exact name of the column header to find 
                           (e.g., "Shipping Order Number *").

    Returns:
        list: A list of values from the specified column. Returns an empty list
              if the column or sheet is not found or if an error occurs.
    """
    print(f"Starting extraction from '{file_path}'...")
    
    # Find the cell reference of the column header (e.g., 'A15')
    header_cell_ref = get_xls_cell_reference_by_value(file_path, sheet_name, column_name)
    
    if "Error:" in header_cell_ref:
        print(f"Error finding column header: {header_cell_ref}")
        return []
        
    print(f"Found column header '{column_name}' at cell {header_cell_ref}.")

    # Extract the column letter(s) and header row number from the reference
    match = re.match(r"([A-Za-z]+)([0-9]+)", header_cell_ref)
    if not match:
        print("Error: Could not parse the header cell reference.")
        return []
    
    col_letters, header_row_num_str = match.groups()
    start_row = int(header_row_num_str) + 1
    
    # Find the last row with any data in the sheet
    last_row, _ = get_xls_last_row(file_path, sheet_name)
    
    if last_row is None:
        print("Error: Could not determine the last row of the sheet.")
        return []
        
    print(f"Data extraction will run from row {start_row} to {last_row}.")

    # Loop from the row after the header to the last row and collect cell values
    column_values = []
    for current_row in range(start_row, last_row + 1):
        cell_to_read = f"{col_letters}{current_row}"
        value = get_xls_cell_value(file_path, sheet_name, cell_to_read)
        
        # Stop if we hit an error or an empty cell, assuming data is contiguous
        if isinstance(value, str) and "Error:" in value:
            print(f"Encountered an error reading {cell_to_read}: {value}")
            break
        if value == "":
            continue # Skip empty cells

        column_values.append(value)
        
    print("Extraction complete.")
    return column_values