"""
固定路径 PDF 读取预览
------------------------------------------------
依赖：pip install pymupdf
"""

from pathlib import Path
import re
import fitz  # PyMuPDF

# 固定 PDF 路径
PDF_PATH = Path(r"C:\Users\User\Downloads\mos_pdfs\Savori (2808).pdf")

# 正则
UNIT_RE = r"(?:ctn|ctns|pkt|pkts|tin|tins|can|cans|box|boxes|btl|btls|pc|pcs)"
PO_LINE_RE = re.compile(r"^\s*To\s*(\d{5,})\s*$", re.I)
ISSUED_DELIVERED_RE = re.compile(r"Issued\s+Delivered\s+(MOS Burger.+?\(\d+\))", re.I)
ITEM_LINE_RE = re.compile(rf"^(\d{{5,}})\s+(.+?)\s+(\d+)\s+{UNIT_RE}\b", re.I)


def main():
    with fitz.open(PDF_PATH) as doc:
        for pno, page in enumerate(doc, start=1):
            text = page.get_text("text") or ""
            lines = [ln for ln in text.splitlines()]

            print(f"\n=== Page {pno} 原始行（前 50 行） ===")
            for idx, ln in enumerate(lines[:50], start=1):
                print(f"{idx:03d}: {ln}")

            # 提取 PO
            po_hits = [m.group(1) for ln in lines if (m := PO_LINE_RE.match(ln.strip()))]
            po = po_hits[-1] if po_hits else ""

            # 提取门店
            store = ""
            for ln in lines:
                m = ISSUED_DELIVERED_RE.search(ln)
                if m:
                    store = m.group(1).strip()
                    break
            if not store:
                for ln in lines:
                    mm = re.search(r"(MOS Burger.+?\(\d+\))", ln, re.I)
                    if mm:
                        store = mm.group(1).strip()
                        break

            print(f"\n--- Page {pno} Store / PO ---")
            print(f"Store_raw: {store or '-'}")
            print(f"PO_raw: {po or '-'}")

            # 提取明细行
            print(f"\n--- Page {pno} 明细行 ---")
            for ln in lines:
                m = ITEM_LINE_RE.match(ln.strip())
                if not m:
                    continue
                code = m.group(1)
                desc = m.group(2)
                qty  = int(m.group(3))
                unit_m = re.search(rf"\b({UNIT_RE})\b", ln, re.I)
                unit = unit_m.group(1) if unit_m else ""
                print(f"code={code} | qty={qty} {unit} | desc={desc}")


if __name__ == "__main__":
    main()
