from pathlib import Path
import shutil
import fitz  # PyMuPDF

# 目标文件夹（按你的路径填写，前面加 r 防止转义）
ROOT = Path(r"C:\Users\User\Dropbox\Halal Update\Halal\4. Innofresh\[2025-2026] Halal certificate of Innofresh")

# 是否给原文件做 .bak 备份（True/False）
MAKE_BACKUP = True

def rotate_second_page(pdf_path: Path, angle: int = 90) -> None:
    """将 pdf 的第 2 页顺时针旋转 angle 度并覆盖保存。"""
    try:
        with fitz.open(pdf_path) as doc:
            if doc.page_count < 2:
                print(f"[跳过] 页数不足 2 页：{pdf_path.name}")
                return

            page = doc[1]  # 第二张（索引从 0 开始）
            # 保留已有旋转，再加 90 度
            new_rot = (page.rotation + angle) % 360
            page.set_rotation(new_rot)

            # 安全保存：先存到临时文件，再替换原文件
            tmp_path = pdf_path.with_suffix(".tmp.pdf")
            doc.save(tmp_path)

        if MAKE_BACKUP:
            bak_path = pdf_path.with_suffix(".bak.pdf")
            if not bak_path.exists():
                shutil.copy2(pdf_path, bak_path)

        tmp_path.replace(pdf_path)
        print(f"[完成] 已旋转第 2 页：{pdf_path.name}")

    except Exception as e:
        print(f"[失败] {pdf_path.name} -> {e}")

def main():
    if not ROOT.exists():
        print(f"目标文件夹不存在：{ROOT}")
        return

    # 仅遍历该目录下的 PDF（不递归子文件夹，若要递归改为 ROOT.rglob('*.pdf')）
    for pdf in ROOT.glob("*.pdf"):
        rotate_second_page(pdf, angle=90)

if __name__ == "__main__":
    main()
