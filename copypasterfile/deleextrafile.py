import os
import shutil

def clean_empty_apr_folders(base_path):
    """
    Recursively searches for '4. Apr' folders inside any path that includes
    'format' or 'history' (case-insensitive), and deletes them ONLY if they are empty.

    Args:
        base_path (str): The base directory to start searching from.
    """
    for root, dirs, files in os.walk(base_path):
        root_lower = root.lower()

        # Check if path includes 'format' or 'history'
        if 'format' in root_lower or 'history' in root_lower:
            for folder in dirs:
                if folder.lower() == '5. may':
                    folder_path = os.path.join(root, folder)

                    # Check if folder is empty
                    if not os.listdir(folder_path):
                        try:
                            os.rmdir(folder_path)
                            print(f"✅ Deleted empty folder: {folder_path}")
                        except Exception as e:
                            print(f"❌ Failed to delete {folder_path}: {e}")
                    else:
                        print(f"⚠️ Skipped (not empty): {folder_path}")

# Example usage
if __name__ == "__main__":
    base_directory = r"C:\\Users\\User\\Dropbox\\DO & INV\\DO & INV 2025"
    clean_empty_apr_folders(base_directory)
