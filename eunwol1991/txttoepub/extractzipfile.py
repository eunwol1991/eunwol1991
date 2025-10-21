import os
import zipfile
import shutil
from pathlib import Path

def extract_and_cleanup(folder_path, src_encoding='cp437', target_encoding='gbk'):
    folder = Path(folder_path)
    temp_dir = folder / "_temp_extract"
    temp_dir.mkdir(exist_ok=True)

    for zip_path in folder.glob('*.zip'):
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                # 手动按编码转换文件名并解压
                for zip_info in zip_ref.infolist():
                    # 跳过根目录项
                    if zip_info.filename.endswith('/'):
                        # 创建目录
                        decoded_dir = zip_info.filename.encode(src_encoding).decode(target_encoding, 'ignore')
                        (temp_dir / decoded_dir).mkdir(parents=True, exist_ok=True)
                        continue
                    # 解码文件名
                    decoded_name = zip_info.filename.encode(src_encoding).decode(target_encoding, 'ignore')
                    dest_path = temp_dir / decoded_name
                    dest_path.parent.mkdir(parents=True, exist_ok=True)
                    # 写入文件内容
                    with zip_ref.open(zip_info) as src_file, open(dest_path, 'wb') as out_file:
                        shutil.copyfileobj(src_file, out_file)

            # 将解压后的所有文件移动到目标文件夹
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    src_file = Path(root) / file
                    dest_file = folder / file
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