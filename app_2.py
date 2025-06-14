import json
from file_utils import get_latest_file
from excel_legacy_utils import update_xls_cell
from func_utils import  find_row_and_get_values, clean_number,get_column_values_with_row_numbers


if __name__ == "__main__":
    with open('column_mappings.json', 'r') as column_mapping_file:
        column_mapping = json.load(column_mapping_file)

    print(list(column_mapping.values()))

    # Example usage
    Mass_Update_folder_path = "docs/Mass_Update"
    Mass_Update_file_format = ".xls"
    Mass_Update_Sheet_name = "Mass Update"
    Mass_Update_latest_file = get_latest_file(Mass_Update_folder_path, Mass_Update_file_format)


    print("Load Plan file handling----------------")
    Load_Plan_folder_path = "docs/Load_Plan"
    Load_Plan_file_format = ".xlsx"
    Load_Plan_Sheet_name = "LLL Load Plan - 16 June 25"
    Load_Plan_latest_file = get_latest_file(Load_Plan_folder_path, Load_Plan_file_format)

    COLUMN_HEADER = "Shipping Order Number *"

    # # Call the function and get the list of shipping order numbers
    shipping_orders = get_column_values_with_row_numbers(Mass_Update_latest_file, Mass_Update_Sheet_name, COLUMN_HEADER)

    # # Print the results
    if shipping_orders:
        print("\n--- Extracted Shipping Order Numbers ---")
        for order in shipping_orders:
            print(order)
        print(f"\nTotal orders found: {len(shipping_orders)}")
    else:
        print("\nCould not extract any shipping order numbers.")

    SEARCH_COLUMN = "SO#"
    COLUMNS_TO_GET = list(column_mapping.values())

    # Call the function with the parameters
    for order in shipping_orders:
        found_row, data = find_row_and_get_values(Load_Plan_latest_file, Load_Plan_Sheet_name, SEARCH_COLUMN, clean_number(order[1]), COLUMNS_TO_GET)
        update_xls_cell(
            file_path=Mass_Update_latest_file,
            sheet_name=Mass_Update_Sheet_name,
            cell_ref="J" + str(order[0]),
            update_value=data['LSP Requested HOD']
        )
        update_xls_cell(
            file_path=Mass_Update_latest_file,
            sheet_name=Mass_Update_Sheet_name,
            cell_ref="Q" + str(order[0]),
            update_value=data['ETD Port Of Load Date']
        )
        update_xls_cell(
            file_path=Mass_Update_latest_file,
            sheet_name=Mass_Update_Sheet_name,
            cell_ref="R" + str(order[0]),
            update_value=data['ETA Port Of Discharge Date']
        )
        update_xls_cell(
            file_path=Mass_Update_latest_file,
            sheet_name=Mass_Update_Sheet_name,
            cell_ref="S" + str(order[0]),
            update_value=data['ETA IN DC Date']
        )
        update_xls_cell(
            file_path=Mass_Update_latest_file,
            sheet_name=Mass_Update_Sheet_name,
            cell_ref="U" + str(order[0]),
            update_value=data['ETA IN DC Date']
        )
        print(order)
        print(found_row, data)
