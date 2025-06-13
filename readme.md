# Excel File Handling Project

A Python utility for handling Excel (.xls) files with functions for finding the latest file, updating cell values, and retrieving row information.

## Features

- Find the latest file in a directory by file extension
- Update specific cell values in .xls files
- Get the last row number and its first column value from a worksheet

## Prerequisites

- Python 3.x
- Required packages:
  - xlrd
  - xlwt
  - xlutils

## Installation

Install the required packages using pip:

```bash
pip install xlrd xlwt xlutils
```

## Usage

### Finding the Latest File

```python
from app_1 import get_latest_file

latest_file = get_latest_file(folder="docs", file_format=".xls")
print(latest_file)
```

### Updating a Cell Value

```python
from app_1 import update_xls_cell

status = update_xls_cell(
    file_path="path/to/file.xls",
    sheet_name="Sheet1",
    cell_ref="D6",
    update_value="06/09/2027"
)
print(status)
```

### Getting Last Row Information

```python
from app_1 import get_last_row_and_first_col_value

row_num, first_col_value = get_last_row_and_first_col_value(
    file_path="path/to/file.xls",
    sheet_name="Sheet1"
)
print(f"Last row: {row_num}, First column value: {first_col_value}")
```

## Function Documentation

### get_latest_file(folder, file_format)

Returns the full path of the latest file in the specified folder matching the given file format.

### update_xls_cell(file_path, sheet_name, cell_ref, update_value)

Updates a specific cell value in an .xls file. Returns 'Success' or an error message.

### get_last_row_and_first_col_value(file_path, sheet_name)

Returns the last row number and its first column value from the specified worksheet.

## Limitations

- Only supports .xls file format (not .xlsx)
- Requires write permissions for file updates
- Maintains existing cell formatting when updating values

## Contributing

Feel free to submit issues and enhancement requests.

## License

This project is available for use under the MIT
