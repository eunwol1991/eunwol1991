import os
import zipfile
import shutil
from pathlib import Path

def extract_and_cleanup(folder_path):
    folder = Path(folder_path)
    # 临时解压目录
    temp_dir = folder / "_temp_extract"
    temp_dir.mkdir(exist_ok=True)

    # 遍历所有zip文件
    for zip_path in folder.glob('*.zip'):
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # 将解压后的所有文件移动到目标文件夹
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    src_file = Path(root) / file
                    dest_file = folder / file
                    # 若目标已存在同名文件，可选择覆盖或重命名
                    if dest_file.exists():
                        dest_file.unlink()
                    shutil.move(str(src_file), str(dest_file))

            # 删除zip文件
            zip_path.unlink()

            # 清空临时目录
            for item in temp_dir.iterdir():
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()

        except Exception as e:
            print(f"处理 {zip_path.name} 时出错: {e}")

if __name__ == "__main__":
    target_folder = r"C:\Users\User\Desktop\txt to epub\txt file"
    extract_and_cleanup(target_folder)
    print("全部zip文件已处理完毕，已清理临时文件。")
