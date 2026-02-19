import os
from PyPDF2 import PdfMerger


def _is_wsl() -> bool:
    if os.name == "nt":
        return False
    if os.environ.get("WSL_DISTRO_NAME"):
        return True
    try:
        with open("/proc/version", "r", encoding="utf-8") as f:
            return "microsoft" in f.read().lower()
    except OSError:
        return False


def _platform_drive_root() -> str:
    if os.name == "nt":
        return "c:/"
    if _is_wsl():
        return "/mnt/c"
    return "/"


def _from_c(path_tail: str) -> str:
    tail = (path_tail or "").lstrip("/")
    root = _platform_drive_root()
    if root.endswith("/"):
        return f"{root}{tail}"
    return f"{root}/{tail}"


def merge_cover_to_all(input_folder, output_folder, cover_filename):
    # 获取封面完整路径
    cover_path = os.path.join(input_folder, cover_filename)

    # 检查封面是否存在
    if not os.path.exists(cover_path):
        print(f"❌ 找不到封面文件：{cover_path}")
        return

    # 创建输出文件夹（如不存在）
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历所有 PDF
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf") and filename != cover_filename:
            target_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, filename)

            try:
                merger = PdfMerger()
                merger.append(cover_path)
                merger.append(target_path)
                merger.write(output_path)
                merger.close()
                print(f"✅ 已合并 → {output_path}")
            except Exception as e:
                print(f"❌ 合并失败 → {filename}，错误：{e}")


# === 参数设置 ===
input_path = _from_c(
    "Users/jhunj/Downloads/Re_ Innofrsh_Savori request for updated Halal cert (Exp_ May & Jun'25)"
)
output_path = _from_c("Users/jhunj/Downloads/halal cert")
cover_pdf_name = "Halal A477-2548 valid 05.08.2026 EN cover.pdf"

# === 执行函数 ===
merge_cover_to_all(input_path, output_path, cover_pdf_name)
