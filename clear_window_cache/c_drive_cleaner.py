import os
import shutil
from pathlib import Path
from datetime import datetime

# 设置要清理的路径（名称: 路径）
CLEANUP_PATHS = {
    "User Temp (AppData)": str(Path.home() / "AppData" / "Local" / "Temp"),
    "System Temp": r"C:\Windows\Temp",
    "Downloads": str(Path.home() / "Downloads"),
    "Windows Update Cache": r"C:\Windows\SoftwareDistribution\Download",
    "Recycle Bin": r"C:\$Recycle.Bin"
}

# 将 byte 转为 MB 显示
def bytes_to_mb(b):
    return round(b / (1024 * 1024), 2)

# 获取文件夹大小
def get_directory_size(path):
    total = 0
    for dirpath, _, filenames in os.walk(path):
        for f in filenames:
            try:
                fp = os.path.join(dirpath, f)
                if os.path.exists(fp):
                    total += os.path.getsize(fp)
            except:
                pass
    return total

# 删除目录内容
def delete_contents(path, log_list):
    for root, dirs, files in os.walk(path):
        for file in files:
            try:
                full_path = os.path.join(root, file)
                os.remove(full_path)
                log_list.append(full_path)
            except:
                continue
        for dir in dirs:
            try:
                shutil.rmtree(os.path.join(root, dir), ignore_errors=True)
            except:
                continue

# 主程序
def main():
    print("🧠 C盘清理预览模式")
    print("=============================")
    delete_log = []
    for name, path in CLEANUP_PATHS.items():
        print(f"\n📁 [{name}]")
        if not os.path.exists(path):
            print("⛔ 路径不存在，跳过")
            continue

        size = get_directory_size(path)
        print(f"📦 占用空间：{bytes_to_mb(size)} MB")
        choice = input("是否要删除这个目录内容？(y/n): ").strip().lower()

        if choice == "y":
            print(f"🧹 正在删除 {name} ...")
            delete_contents(path, delete_log)
            print("✅ 删除完成")
        else:
            print("🚫 已跳过")

    # 生成日志
    if delete_log:
        log_file = f"c_clean_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(log_file, "w", encoding="utf-8") as f:
            f.write("🧾 删除文件列表：\n")
            for item in delete_log:
                f.write(item + "\n")
        print(f"\n📝 删除记录已写入：{log_file}")
    else:
        print("\n🧼 本次未执行任何删除。")

    print("\n🎉 清理程序结束")

if __name__ == "__main__":
    main()
