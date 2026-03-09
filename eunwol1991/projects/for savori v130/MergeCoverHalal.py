import os
import importlib
import importlib.util


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


def merge_cover_with_file(
    cover_pdf_path: str, target_pdf_path: str, output_pdf_path: str
):
    pypdf_spec = importlib.util.find_spec("pypdf")
    if pypdf_spec is None:
        print("❌ 缺少 PDF 库，请先安装：pip install pypdf")
        return

    pypdf_module = importlib.import_module("pypdf")
    writer_cls = getattr(pypdf_module, "PdfWriter", None)
    if writer_cls is None:
        print("❌ 当前 pypdf 不包含 PdfWriter，请升级：pip install -U pypdf")
        return

    if not os.path.exists(cover_pdf_path):
        print(f"❌ 找不到封面文件：{cover_pdf_path}")
        return
    if not os.path.exists(target_pdf_path):
        print(f"❌ 找不到目标文件：{target_pdf_path}")
        return

    output_dir = os.path.dirname(output_pdf_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        cover_real = os.path.realpath(cover_pdf_path)
        target_real = os.path.realpath(target_pdf_path)
        output_real = os.path.realpath(output_pdf_path)
    except OSError as e:
        print(f"❌ 路径解析失败，错误：{e}")
        return

    if output_real in {cover_real, target_real}:
        print("❌ 输出文件不能和封面/目标文件相同")
        return

    try:
        writer = writer_cls()
        writer.append(cover_pdf_path)
        writer.append(target_pdf_path)
        _ = writer.write(output_pdf_path)
        print(f"✅ 已合并 → {output_pdf_path}")
    except Exception as e:
        print(f"❌ 合并失败，错误：{e}")


def merge_cover_with_each_page(
    cover_pdf_path: str,
    target_pdf_path: str,
    output_folder: str,
    output_prefix: str,
):
    pypdf_spec = importlib.util.find_spec("pypdf")
    if pypdf_spec is None:
        print("❌ 缺少 PDF 库，请先安装：pip install pypdf")
        return []

    pypdf_module = importlib.import_module("pypdf")
    writer_cls = getattr(pypdf_module, "PdfWriter", None)
    reader_cls = getattr(pypdf_module, "PdfReader", None)
    if writer_cls is None or reader_cls is None:
        print("❌ 当前 pypdf 不包含 PdfWriter/PdfReader，请升级：pip install -U pypdf")
        return []

    if not os.path.exists(cover_pdf_path):
        print(f"❌ 找不到封面文件：{cover_pdf_path}")
        return []
    if not os.path.exists(target_pdf_path):
        print(f"❌ 找不到目标文件：{target_pdf_path}")
        return []

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        cover_reader = reader_cls(cover_pdf_path)
        target_reader = reader_cls(target_pdf_path)
    except Exception as e:
        print(f"❌ 读取 PDF 失败，错误：{e}")
        return []

    target_page_count = len(target_reader.pages)
    if target_page_count == 0:
        print(f"⚠️ 目标文件没有页面：{target_pdf_path}")
        return []

    output_paths = []
    for idx in range(target_page_count):
        page_no = idx + 1
        output_path = os.path.join(output_folder, f"{output_prefix}_p{page_no:03d}.pdf")

        try:
            writer = writer_cls()
            for cover_page in cover_reader.pages:
                writer.add_page(cover_page)
            writer.add_page(target_reader.pages[idx])
            _ = writer.write(output_path)
            output_paths.append(output_path)
            print(f"✅ 已生成 → {output_path}")
        except Exception as e:
            print(f"❌ 生成失败（第 {page_no} 页），错误：{e}")

    return output_paths


# === 参数设置 ===
cover_pdf_path = _from_c(
    "Users/jhunj/Dropbox/for jj/Halal/A477_ Innofresh_cover exp.30-05-2027 EN.pdf"
)
target_pdf_path = _from_c(
    "Users/jhunj/Dropbox/for jj/Halal/A477_ Innofresh_total list exp.30-05-2027.pdf"
)
output_pdf_path = _from_c(
    "Users/jhunj/Dropbox/for jj/Halal/merged/A477_ Innofresh_total list exp.30-05-2027 (with cover).pdf"
)
output_folder = _from_c("Users/jhunj/Dropbox/for jj/Halal/merged/pages")
output_prefix = "A477_ Innofresh_total list exp.30-05-2027"

# === 执行函数 ===
if __name__ == "__main__":
    merge_cover_with_each_page(
        cover_pdf_path=cover_pdf_path,
        target_pdf_path=target_pdf_path,
        output_folder=output_folder,
        output_prefix=output_prefix,
    )
