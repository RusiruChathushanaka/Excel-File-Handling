from file_utils import get_latest_file
from excel_legacy_utils import (get_xls_cell_value,update_xls_cell,get_xls_last_row,get_xls_cell_reference_by_value)
from excel_new_utils import get_xlsx_cell_value, update_xlsx_cell, get_xlsx_last_row, get_xlsx_cell_reference_by_value
from func_utils import get_column_values

if __name__ == "__main__":
    # Example usage
    Mass_Update_folder_path = "docs/Mass_Update" 
    Mass_Update_file_format = ".xls"
    Mass_Update_Sheet_name = "Mass Update"
    Mass_Update_latest_file = get_latest_file(Mass_Update_folder_path, Mass_Update_file_format)


    if Mass_Update_latest_file:
        print(f"The latest file is: {Mass_Update_latest_file}")
    else:
        print("No files found matching the specified format.")


    status = update_xls_cell(
    file_path=Mass_Update_latest_file,
    sheet_name="Mass Update",
    cell_ref="D6",
    update_value="06/09/2027"
            )
    print(status)

    row_num, first_col_value = get_xls_last_row(Mass_Update_latest_file, "Mass Update")
    print("Last row number:", row_num)
    print("First column value in last row:", first_col_value)

    value = get_xls_cell_value(Mass_Update_latest_file, "Mass Update", "F6")
    print("Value in F6:", value)


    cell_ref = get_xls_cell_reference_by_value(Mass_Update_latest_file, "Mass Update", "Booked HOD *")
    print("Cell reference:", cell_ref)


    print("Load Plan file handling----------------")
    Load_Plan_folder_path = "docs/Load_Plan" 
    Load_Plan_file_format = ".xlsx"
    Load_Plan_Sheet_name = "LLL Load Plan - 16 June 25"
    Load_Plan_latest_file = get_latest_file(Load_Plan_folder_path, Load_Plan_file_format)


    if Load_Plan_latest_file:
        print(f"The latest file is: {Load_Plan_latest_file}")
    else:
        print("No files found matching the specified format.")

    status = update_xlsx_cell(
        file_path=Load_Plan_latest_file,
        sheet_name=Load_Plan_Sheet_name,
        cell_ref="AQ6",
        update_value="Updated Value"
    )
    print(status)

    row_num, first_col_value = get_xlsx_last_row(Load_Plan_latest_file, Load_Plan_Sheet_name)
    print("Last row number:", row_num)
    print("First column value in last row:", first_col_value)

    value = get_xlsx_cell_value(Load_Plan_latest_file, Load_Plan_Sheet_name, "F6")
    print("Value in F6:", value)

    cell_ref = get_xlsx_cell_reference_by_value(Load_Plan_latest_file, Load_Plan_Sheet_name, "Book HOD")
    print("Cell reference:", cell_ref)


    print("func utils library----------------")

    COLUMN_HEADER = "Shipping Order Number *"

    # Call the function and get the list of shipping order numbers
    shipping_orders = get_column_values(Mass_Update_latest_file, Mass_Update_Sheet_name, COLUMN_HEADER)

    # Print the results
    if shipping_orders:
        print("\n--- Extracted Shipping Order Numbers ---")
        for order in shipping_orders:
            print(order)
        print(f"\nTotal orders found: {len(shipping_orders)}")
    else:
        print("\nCould not extract any shipping order numbers.")
