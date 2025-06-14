import json
import sys
from file_utils import get_latest_file
from excel_legacy_utils import update_xls_cell
from func_utils import find_row_and_get_values, clean_number, get_column_values_with_row_numbers

# --- Configuration Constants ---
# Grouping all settings here makes the script easier to configure.

# Mass Update File Settings
MASS_UPDATE_FOLDER_PATH = "docs/Mass_Update"
MASS_UPDATE_FILE_FORMAT = ".xls"
MASS_UPDATE_SHEET_NAME = "Mass Update"
MASS_UPDATE_SHIPPING_ORDER_COLUMN = "Shipping Order Number *"

# Load Plan File Settings
LOAD_PLAN_FOLDER_PATH = "docs/Load_Plan"
LOAD_PLAN_FILE_FORMAT = ".xlsx"
LOAD_PLAN_SHEET_NAME = "LLL Load Plan - 16 June 25"
LOAD_PLAN_SEARCH_COLUMN = "SO#"

# Mappings File
COLUMN_MAPPINGS_FILE = 'column_mappings.json'


def main():
    """
    Main function to orchestrate the process of updating the Mass Update sheet
    with data from the Load Plan sheet.
    """
    # --- 1. Load Column Mappings ---
    try:
        with open(COLUMN_MAPPINGS_FILE, 'r') as f:
            column_mapping = json.load(f)
        print("Successfully loaded column mappings.")
    except FileNotFoundError:
        print(f"Error: The mapping file '{COLUMN_MAPPINGS_FILE}' was not found.")
        sys.exit(1) # Exit the script if the mapping is missing
    except json.JSONDecodeError:
        print(f"Error: The mapping file '{COLUMN_MAPPINGS_FILE}' is not a valid JSON file.")
        sys.exit(1)

    # --- 2. Get Latest Files ---
    print("\nLocating latest files...")
    mass_update_file = get_latest_file(MASS_UPDATE_FOLDER_PATH, MASS_UPDATE_FILE_FORMAT)
    load_plan_file = get_latest_file(LOAD_PLAN_FOLDER_PATH, LOAD_PLAN_FILE_FORMAT)

    if not mass_update_file or not load_plan_file:
        print("Error: Could not find one or both of the required Excel files. Exiting.")
        sys.exit(1)
        
    print(f"Found Mass Update file: {mass_update_file}")
    print(f"Found Load Plan file: {load_plan_file}")

    # --- 3. Extract Shipping Orders from Mass Update File ---
    print(f"\nExtracting shipping orders from '{MASS_UPDATE_SHIPPING_ORDER_COLUMN}' column...")
    shipping_orders = get_column_values_with_row_numbers(
        mass_update_file,
        MASS_UPDATE_SHEET_NAME,
        MASS_UPDATE_SHIPPING_ORDER_COLUMN
    )

    if not shipping_orders:
        print("Could not find any shipping order numbers to process. Exiting.")
        return

    print(f"Found {len(shipping_orders)} shipping orders to process.")

    # --- 4. Process Each Order ---
    print("\n--- Starting Update Process ---")
    columns_to_get_from_load_plan = list(set(column_mapping.values())) # Use set to avoid duplicate column names

    for row_num, order_number in shipping_orders:
        cleaned_order = clean_number(order_number)
        
        print(f"\nProcessing Order: '{cleaned_order}' from row {row_num}...")

        # Find the corresponding data in the Load Plan file
        found_row, data = find_row_and_get_values(
            load_plan_file,
            LOAD_PLAN_SHEET_NAME,
            LOAD_PLAN_SEARCH_COLUMN,
            cleaned_order,
            columns_to_get_from_load_plan
        )

        if not found_row:
            print(f"  -> WARNING: Could not find a match for order '{cleaned_order}' in the Load Plan file.")
            continue

        print(f"  -> Found matching data in Load Plan at row {found_row}.")

        # Iterate through the column mappings to update the Mass Update file
        for target_col, source_col in column_mapping.items():
            update_value = data.get(source_col)
            
            if update_value is not None:
                cell_ref = f"{target_col}{row_num}"
                print(f"    - Updating cell {cell_ref} with value: '{update_value}'")
                update_xls_cell(
                    file_path=mass_update_file,
                    sheet_name=MASS_UPDATE_SHEET_NAME,
                    cell_ref=cell_ref,
                    update_value=update_value
                )
            else:
                print(f"  -> WARNING: Source column '{source_col}' not found in data for order '{cleaned_order}'.")

    print("\n--- Update Process Finished ---")


if __name__ == "__main__":
    main()