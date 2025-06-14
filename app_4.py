# main_script.py

import json
import sys
from file_utils import get_latest_file
from excel_legacy_utils import update_xls_cell
from func_utils import find_row_and_get_values, clean_number, get_column_values_with_row_numbers
from logger_utils import setup_logger # <-- IMPORT THE NEW LOGGER

# --- Configuration Constants ---
MASS_UPDATE_FOLDER_PATH = "docs/Mass_Update"
MASS_UPDATE_FILE_FORMAT = ".xls"
MASS_UPDATE_SHEET_NAME = "Mass Update"
MASS_UPDATE_SHIPPING_ORDER_COLUMN = "Shipping Order Number *"

LOAD_PLAN_FOLDER_PATH = "docs/Load_Plan"
LOAD_PLAN_FILE_FORMAT = ".xlsx"
LOAD_PLAN_SHEET_NAME = "LLL Load Plan - 16 June 25"
LOAD_PLAN_SEARCH_COLUMN = "SO#"

COLUMN_MAPPINGS_FILE = 'column_mappings.json'

def main():
    """
    Main function to orchestrate the process of updating the Mass Update sheet
    with data from the Load Plan sheet.
    """
    # --- 0. Setup Logger ---
    # Gather all constants into a dictionary to pass to the logger
    config_for_logging = {
        "Mass Update Folder": MASS_UPDATE_FOLDER_PATH,
        "Mass Update File Format": MASS_UPDATE_FILE_FORMAT,
        "Mass Update Sheet": MASS_UPDATE_SHEET_NAME,
        "Load Plan Folder": LOAD_PLAN_FOLDER_PATH,
        "Load Plan File Format": LOAD_PLAN_FILE_FORMAT,
        "Load Plan Sheet": LOAD_PLAN_SHEET_NAME,
        "Column Mappings File": COLUMN_MAPPINGS_FILE
    }
    logger = setup_logger(config_for_logging)

    # --- 1. Load Column Mappings ---
    try:
        with open(COLUMN_MAPPINGS_FILE, 'r') as f:
            column_mapping = json.load(f)
        logger.info("Successfully loaded column mappings.")
    except FileNotFoundError:
        logger.error(f"Error: The mapping file '{COLUMN_MAPPINGS_FILE}' was not found.")
        sys.exit(1)
    except json.JSONDecodeError:
        logger.error(f"Error: The mapping file '{COLUMN_MAPPINGS_FILE}' is not a valid JSON file.")
        sys.exit(1)

    # --- 2. Get Latest Files ---
    logger.info("Locating latest files...")
    mass_update_file = get_latest_file(MASS_UPDATE_FOLDER_PATH, MASS_UPDATE_FILE_FORMAT)
    load_plan_file = get_latest_file(LOAD_PLAN_FOLDER_PATH, LOAD_PLAN_FILE_FORMAT)

    if not mass_update_file or not load_plan_file:
        logger.critical("Error: Could not find one or both of the required Excel files. Exiting.")
        sys.exit(1)
        
    logger.info(f"Found Mass Update file: {mass_update_file}")
    logger.info(f"Found Load Plan file: {load_plan_file}")

    # --- 3. Extract Shipping Orders from Mass Update File ---
    logger.info(f"Extracting shipping orders from '{MASS_UPDATE_SHIPPING_ORDER_COLUMN}' column...")
    shipping_orders = get_column_values_with_row_numbers(
        mass_update_file,
        MASS_UPDATE_SHEET_NAME,
        MASS_UPDATE_SHIPPING_ORDER_COLUMN
    )

    if not shipping_orders:
        logger.warning("Could not find any shipping order numbers to process. Exiting.")
        return

    logger.info(f"Found {len(shipping_orders)} shipping orders to process.")

    # --- 4. Process Each Order ---
    logger.info("--- Starting Update Process ---")
    columns_to_get_from_load_plan = list(set(column_mapping.values()))

    for row_num, order_number in shipping_orders:
        cleaned_order = clean_number(order_number)
        
        logger.info(f"Processing Order: '{cleaned_order}' from row {row_num}...")

        found_row, data = find_row_and_get_values(
            load_plan_file,
            LOAD_PLAN_SHEET_NAME,
            LOAD_PLAN_SEARCH_COLUMN,
            cleaned_order,
            columns_to_get_from_load_plan
        )

        if not found_row:
            logger.warning(f"  -> Could not find a match for order '{cleaned_order}' in the Load Plan file.")
            continue

        logger.info(f"  -> Found matching data in Load Plan at row {found_row}.")

        for target_col, source_col in column_mapping.items():
            update_value = data.get(source_col)
            
            if update_value is not None:
                cell_ref = f"{target_col}{row_num}"
                logger.info(f"    - Updating cell {cell_ref} with value: '{update_value}'")
                update_xls_cell(
                    file_path=mass_update_file,
                    sheet_name=MASS_UPDATE_SHEET_NAME,
                    cell_ref=cell_ref,
                    update_value=update_value
                )
            else:
                logger.warning(f"  -> Source column '{source_col}' not found in data for order '{cleaned_order}'.")

    logger.info("--- Update Process Finished ---")


if __name__ == "__main__":
    main()