import os
import shutil
import re

def move_files_by_date_pattern(source_folder, target_folder, pattern, debug=False):
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
        print(f"📁 创建目标资料夹：{target_folder}")

    moved_count = 0

    for filename in os.listdir(source_folder):
        if re.search(pattern, filename, re.IGNORECASE):
            source_path = os.path.join(source_folder, filename)
            target_path = os.path.join(target_folder, filename)

            try:
                shutil.move(source_path, target_path)
                moved_count += 1
                print(f"✅ 移动：{filename}")
            except Exception as e:
                print(f"❌ 移动失败：{filename}，错误：{e}")

        elif debug:
            print(f"🟡 忽略：{filename}")

    print(f"\n🎯 总共成功移动 {moved_count} 个文件。")

# 路径设定
source = r"C:\Users\User\Dropbox\Halal Update\Halal\4. Innofresh\Year 2025(exp)"
target = r"C:\Users\User\Dropbox\Halal Update\Halal\4. Innofresh\Exp 4 May 2025"

# 匹配 "Exp 1~31 May 25/2025"
pattern = r"Exp\s*(?:[1-9]|[12][0-9]|3[01])\s*May\s*(?:25|2025)"

move_files_by_date_pattern(source, target, pattern, debug=True)
