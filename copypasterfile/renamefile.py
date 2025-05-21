import os
from tkinter import Tk, filedialog

def rename_excel_files():
    """
    Rename Excel files in a user-selected directory and its subdirectories by replacing 'xx24' with 'xx25'.
    """
    try:
        # Open a file dialog to select the directory
        Tk().withdraw()  # Hide the root window
        directory = filedialog.askdirectory(title="Select Directory")

        if not directory:
            print("No directory selected. Exiting.")
            return

        for root, _, files in os.walk(directory):
            for filename in files:
                # Check if the file is an Excel file and contains 'xx24'
                if filename.endswith(('.xlsx', '.xls')) and 'xx24' in filename:
                    new_filename = filename.replace('xx24', 'xx25')
                    # Construct full file paths
                    old_path = os.path.join(root, filename)
                    new_path = os.path.join(root, new_filename)
                    # Rename the file
                    os.rename(old_path, new_path)
                    print(f"Renamed: {old_path} -> {new_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
rename_excel_files()
