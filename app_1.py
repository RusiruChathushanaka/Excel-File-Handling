import os
import re
import glob
import xlrd
import xlwt
from xlutils.copy import copy

def get_latest_file(folder, file_format):
    """
    Returns the full path of the latest file in the folder matching the given file format.
    
    Args:
        folder (str): Path to the directory.
        file_format (str): File extension, e.g., '.xlsx', '.csv', '.txt'.
    
    Returns:
        str: Full path of the latest file, or None if no matching files are found.
    """
    # Build the search pattern
    pattern = os.path.join(folder, f"*{file_format}")
    # Get list of matching files
    files = glob.glob(pattern)
    if not files:
        return None
    # Find the latest file by creation time (or use os.path.getmtime for modification time)
    latest_file = max(files, key=os.path.getctime)
    return latest_file

def update_xls_cell(file_path, sheet_name, cell_ref, update_value):
    """
    Updates a cell value in an .xls file.
    
    Args:
        file_path (str): Path to the .xls file.
        sheet_name (str): Name of the worksheet.
        cell_ref (str): Cell reference (e.g., 'C5').
        update_value: Value to set in the cell.
    
    Returns:
        str: 'Success' or error reason.
    """
    # Validate file existence
    if not os.path.isfile(file_path):
        return f"Failed: File not found - {file_path}"
    
    # Validate file extension
    if not file_path.lower().endswith('.xls'):
        return "Failed: Only .xls files are supported."
    
    try:
        # Open the workbook
        rb = xlrd.open_workbook(file_path, formatting_info=True)
        if sheet_name not in rb.sheet_names():
            return f"Failed: Sheet '{sheet_name}' not found."
        
        # Find cell coordinates
        match = re.match(r"([A-Za-z]+)([0-9]+)", cell_ref)
        if not match:
            return "Failed: Invalid cell reference format."
        col_letters, row_num = match.groups()
        col_idx = sum([(ord(char.upper()) - 64) * (26 ** i) for i, char in enumerate(reversed(col_letters))]) - 1
        row_idx = int(row_num) - 1
        
        # Copy workbook for writing
        wb = copy(rb)
        ws = wb.get_sheet(rb.sheet_names().index(sheet_name))
        ws.write(row_idx, col_idx, update_value)
        wb.save(file_path)
        return "Success"
    except Exception as e:
        return f"Failed: {str(e)}"

def get_last_row_and_first_col_value(file_path, sheet_name):
    """
    Finds the last row with data and returns its row number and the value in column A.
    
    Args:
        file_path (str): Path to the .xls file.
        sheet_name (str): Name of the worksheet.
    
    Returns:
        tuple: (last_row_number, first_column_value) or (None, None) if not found or error.
    """
    try:
        wb = xlrd.open_workbook(file_path)
        if sheet_name not in wb.sheet_names():
            return None, "Sheet not found"
        ws = wb.sheet_by_name(sheet_name)
        last_row = None
        # Iterate from the bottom up to find the last row with any data
        for row_idx in range(ws.nrows - 1, -1, -1):
            if any(ws.cell_value(row_idx, col_idx) != "" for col_idx in range(ws.ncols)):
                last_row = row_idx
                break
        if last_row is not None:
            first_col_value = ws.cell_value(last_row, 0)
            # Convert to Excel's 1-based row number
            return last_row + 1, first_col_value
        else:
            return None, None
    except Exception as e:
        return None, str(e)
    
    
def get_xls_cell_value(file_path, sheet_name, cell_ref):
    """
    Returns the value of a cell in an .xls Excel file.
    
    Args:
        file_path (str): Path to the .xls file.
        sheet_name (str): Name of the worksheet.
        cell_ref (str): Cell reference (e.g., 'C6').
    
    Returns:
        The value of the cell, or an error message if not found.
    """
    try:
        # Open the workbook and worksheet
        wb = xlrd.open_workbook(file_path)
        if sheet_name not in wb.sheet_names():
            return f"Error: Sheet '{sheet_name}' not found."
        ws = wb.sheet_by_name(sheet_name)
        
        # Parse the cell reference (e.g., 'C6')
        match = re.match(r"([A-Za-z]+)([0-9]+)", cell_ref)
        if not match:
            return "Error: Invalid cell reference format."
        col_letters, row_num = match.groups()
        
        # Convert column letters to 0-based index
        col_idx = sum([(ord(char.upper()) - 64) * (26 ** i) for i, char in enumerate(reversed(col_letters))]) - 1
        row_idx = int(row_num) - 1  # Excel rows are 1-based
        
        # Check bounds
        if row_idx >= ws.nrows or col_idx >= ws.ncols:
            return "Error: Cell reference out of range."
        
        return ws.cell_value(row_idx, col_idx)
    except Exception as e:
        return f"Error: {str(e)}"

def col_idx_to_letters(col_idx):
    """Convert a zero-based column index to Excel-style letters."""
    letters = ''
    while col_idx >= 0:
        letters = chr(col_idx % 26 + ord('A')) + letters
        col_idx = col_idx // 26 - 1
    return letters

def get_cell_reference_by_value(file_path, sheet_name, cell_value):
    """
    Returns the cell reference (e.g., 'A2', 'C8') for the first cell matching the given value.
    
    Args:
        file_path (str): Path to the .xls file.
        sheet_name (str): Name of the worksheet.
        cell_value: Value to search for (case and type sensitive).
    
    Returns:
        str: Cell reference (e.g., 'B5') or an error message.
    """
    try:
        wb = xlrd.open_workbook(file_path)
        if sheet_name not in wb.sheet_names():
            return f"Error: Sheet '{sheet_name}' not found."
        ws = wb.sheet_by_name(sheet_name)
        for row_idx in range(ws.nrows):
            for col_idx in range(ws.ncols):
                if ws.cell_value(row_idx, col_idx) == cell_value:
                    col_letters = col_idx_to_letters(col_idx)
                    cell_ref = f"{col_letters}{row_idx + 1}"
                    return cell_ref
        return "Error: Value not found in sheet."
    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == "__main__":
    # Example usage
    Mass_Update_folder_path = "docs/Mass_Update" 
    Mass_Update_file_format = ".xls"
    Mass_Update_latest_file = get_latest_file(Mass_Update_folder_path, Mass_Update_file_format)

    Load_Plan_folder_path = "docs/Load_Plan" 
    Load_Plan_file_format = ".xlsx"
    Load_Plan_latest_file = get_latest_file(Load_Plan_folder_path, Load_Plan_file_format)

    if Mass_Update_latest_file:
        print(f"The latest file is: {Mass_Update_latest_file}")
    else:
        print("No files found matching the specified format.")

    if Load_Plan_latest_file:
        print(f"The latest file is: {Load_Plan_latest_file}")
    else:
        print("No files found matching the specified format.")

    status = update_xls_cell(
    file_path=Mass_Update_latest_file,
    sheet_name="Mass Update",
    cell_ref="D6",
    update_value="06/09/2027"
            )
    print(status)

    row_num, first_col_value = get_last_row_and_first_col_value(Mass_Update_latest_file, "Mass Update")
    print("Last row number:", row_num)
    print("First column value in last row:", first_col_value)

    value = get_xls_cell_value(Mass_Update_latest_file, "Mass Update", "F6")
    print("Value in F6:", value)


    cell_ref = get_cell_reference_by_value(Mass_Update_latest_file, "Mass Update", "Booked HOD *")
    print("Cell reference:", cell_ref)
