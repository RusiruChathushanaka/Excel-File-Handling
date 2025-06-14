import os
import glob

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