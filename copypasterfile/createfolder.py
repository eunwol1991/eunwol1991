import os

def find_and_create_folders(base_path):
    """
    Searches for folders containing Excel files with 'xx25' in their names.
    If the folder path contains 'format' or 'history', create '4. Apr' in its parent folder.
    Otherwise, create it in the current folder.
    """
    folders_to_create = set()

    for root, dirs, files in os.walk(base_path):
        abs_root = os.path.abspath(root)
        root_lower = abs_root.lower()

        # 找出包含 xx25 文件的路径
        if any(file.lower().endswith('.xlsx') and 'xx25' in file.lower() for file in files):
            # 如果路径中有 format 或 history，则记录“父路径”
            if 'format' in root_lower or 'history' in root_lower:
                target_path = os.path.dirname(abs_root)
                print(f"📂 xx25 found in format/history path, will create in parent: {target_path}")
            else:
                target_path = abs_root
                print(f"📁 xx25 found in: {abs_root}")

            folders_to_create.add(target_path)

    # 创建 4. Apr 文件夹（如果不存在）
    for path in folders_to_create:
        new_folder_path = os.path.join(path, '5. May')
        if os.path.exists(new_folder_path):
            print(f"✅ Folder already exists: {new_folder_path}")
        else:
            os.mkdir(new_folder_path)
            print(f"🚀 Created folder: {new_folder_path}")

# Example usage
if __name__ == "__main__":
    base_directory = r"C:\\Users\\User\\Dropbox\\DO & INV\\DO & INV 2025"
    find_and_create_folders(base_directory)
