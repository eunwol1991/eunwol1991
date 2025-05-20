# -*- coding: utf-8 -*-
"""Convert Chinese novel text files to EPUB.
Usage:
    # Interactive mode (run with no arguments)
    python3 epub_converter.py

    # Command line
    python3 epub_converter.py -i input_dir -o output_dir
"""

import argparse
from html import escape
import os
import re
import sys
import uuid
import zipfile
from datetime import datetime, timezone


def detect_encoding(file_path: str) -> str:
    """Try to detect encoding: UTF-8 first, then GBK."""
    for enc in ("utf-8", "gbk", "gb18030"):
        try:
            with open(file_path, "r", encoding=enc) as f:
                f.read()
            return enc
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("Unable to decode file with common encodings")


def detect_author(file_path: str) -> str:
    """Try to detect the author from the first few lines of the file."""
    encoding = detect_encoding(file_path)
    try:
        with open(file_path, "r", encoding=encoding) as f:
            for _ in range(20):
                line = f.readline()
                if not line:
                    break
                m = re.search(r"(?:作者|author)[:：]?\s*(\S+)", line, re.I)
                if m:
                    return m.group(1).strip()
    except Exception:
        pass
    return "无"


def is_chapter_heading(line: str) -> bool:
    """Heuristic check if a line is a chapter heading."""
    line = line.strip()
    if not line:
        return False
    patterns = [
        r"^第[0-9一二三四五六七八九十百千万零〇两]+[章节卷回篇集].*",
        r"^第[0-9一二三四五六七八九十百千万零〇两]+[卷部篇季].*",
        r"^[0-9]{1,4}\s*.*",
        r"^本[书集篇卷季]?完.*",
    ]
    for pattern in patterns:
        if re.match(pattern, line):
            return True
    return False


def parse_chapters(file_path: str):
    encoding = detect_encoding(file_path)
    chapters = []
    title = None
    content_lines = []
    with open(file_path, "r", encoding=encoding) as f:
        for raw_line in f:
            line = raw_line.rstrip()  # keep spaces for indentation
            if is_chapter_heading(line):
                if title or content_lines:
                    chapters.append((title or "前言", "\n".join(content_lines)))
                title = line.strip()
                content_lines = []
            else:
                content_lines.append(line)
    if title or content_lines:
        chapters.append((title or "正文", "\n".join(content_lines)))
    return chapters


def chapter_to_xhtml(index: int, title: str, text: str) -> str:
    paragraphs = "\n".join(
        f"    <p>{escape(p)}</p>" for p in text.splitlines() if p.strip()
    )
    safe_title = escape(title)
    return f"""<?xml version='1.0' encoding='utf-8'?>
<html xmlns='http://www.w3.org/1999/xhtml'>
<head>
  <title>{safe_title}</title>
  <link rel='stylesheet' type='text/css' href='style.css'/>
</head>
<body>
  <h2 id='chapter{index}'>{safe_title}</h2>
{paragraphs}
</body>
</html>"""


def create_epub(title: str, author: str, chapters, output_path: str):
    unique_id = str(uuid.uuid4())
    modified = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    css = "body { font-family: SimSun, serif; line-height: 1.5; }"

    with zipfile.ZipFile(output_path, "w") as epub:
        # mimetype must be first and uncompressed
        epub.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)

        container = """<?xml version='1.0' encoding='UTF-8'?>
<container version='1.0' xmlns='urn:oasis:names:tc:opendocument:xmlns:container'>
  <rootfiles>
    <rootfile full-path='OEBPS/content.opf' media-type='application/oebps-package+xml'/>
  </rootfiles>
</container>"""
        epub.writestr("META-INF/container.xml", container)
        epub.writestr("OEBPS/style.css", css)

        manifest_items = ["<item id='nav' href='nav.xhtml' properties='nav' media-type='application/xhtml+xml'/>",
                          "<item id='css' href='style.css' media-type='text/css'/>",
                          "<item id='toc' href='toc.ncx' media-type='application/x-dtbncx+xml'/>"]
        spine_items = []
        nav_ol = []

        for i, (ch_title, text) in enumerate(chapters, 1):
            filename = f"chapter{i}.xhtml"
            epub.writestr(f"OEBPS/{filename}", chapter_to_xhtml(i, ch_title, text))
            manifest_items.append(
                f"<item id='chapter{i}' href='{filename}' media-type='application/xhtml+xml'/>"
            )
            spine_items.append(f"<itemref idref='chapter{i}'/>")
            nav_ol.append(
                f"      <li><a href='{filename}#chapter{i}'>{escape(ch_title)}</a></li>"
            )

        nav_xhtml = """<?xml version='1.0' encoding='utf-8'?>
<html xmlns='http://www.w3.org/1999/xhtml'>
<head>
  <title>目录</title>
  <link rel='stylesheet' type='text/css' href='style.css'/>
</head>
<body>
  <nav epub:type='toc' id='toc'>
    <h1>目录</h1>
    <ol>
{nav_items}
    </ol>
  </nav>
</body>
</html>""".format(nav_items="\n".join(nav_ol))
        epub.writestr("OEBPS/nav.xhtml", nav_xhtml)

        toc_entries = []
        for i, (ch_title, _) in enumerate(chapters, 1):
            toc_entries.append(
                f"    <navPoint id='navPoint-{i}' playOrder='{i}'>\n      <navLabel><text>{escape(ch_title)}</text></navLabel>\n      <content src='chapter{i}.xhtml'/></navPoint>"
            )
        toc_ncx = """<?xml version='1.0' encoding='utf-8'?>
<!DOCTYPE ncx PUBLIC '-//NISO//DTD ncx 2005-1//EN' 'http://www.daisy.org/z3986/2005/ncx-2005-1.dtd'>
<ncx xmlns='http://www.daisy.org/z3986/2005/ncx/' version='2005-1'>
  <head>
    <meta name='dtb:uid' content='{uid}'/>
    <meta name='dtb:depth' content='1'/>
    <meta name='dtb:totalPageCount' content='0'/>
    <meta name='dtb:maxPageNumber' content='0'/>
  </head>
  <docTitle><text>{escape(title)}</text></docTitle>
  <navMap>
{points}
  </navMap>
</ncx>""".format(uid=unique_id, title=title, points="\n".join(toc_entries))
        epub.writestr("OEBPS/toc.ncx", toc_ncx)

        safe_title = escape(title)
        safe_author = escape(author)
        content_opf = """<?xml version='1.0' encoding='utf-8'?>
<package xmlns='http://www.idpf.org/2007/opf' unique-identifier='bookid' version='3.0'>
  <metadata xmlns:dc='http://purl.org/dc/elements/1.1/'>
    <dc:identifier id='bookid'>{uid}</dc:identifier>
    <dc:title>{safe_title}</dc:title>
    <dc:creator>{safe_author}</dc:creator>
    <dc:language>zh</dc:language>
    <meta property='dcterms:modified'>{modified}</meta>
  </metadata>
  <manifest>
{manifest}
  </manifest>
  <spine toc='toc'>
{spine}
  </spine>
</package>""".format(
            uid=unique_id,
            title=safe_title,
            author=safe_author,
            modified=modified,
            manifest="\n".join(f"    {item}" for item in manifest_items),
            spine="\n".join(f"    {item}" for item in spine_items),
        )
        epub.writestr("OEBPS/content.opf", content_opf)


def convert_txt_to_epub(input_file: str, output_dir: str, title: str = None, author: str = ""):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    if title is None:
        title = os.path.splitext(os.path.basename(input_file))[0]
    if not author:
        author = detect_author(input_file)
    chapters = parse_chapters(input_file)
    output_path = os.path.join(output_dir, f"{title}.epub")
    if os.path.exists(output_path):
        os.remove(output_path)
    create_epub(title, author, chapters, output_path)
    print(f"EPUB created: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Convert Chinese novel .txt to EPUB")
    parser.add_argument("-i", "--input", help="Input file or directory")
    parser.add_argument("-o", "--output", help="Output directory")
    parser.add_argument("-t", "--title", help="Book title (single file mode)")
    parser.add_argument("-a", "--author", help="Author name (single file mode)")
    args = parser.parse_args()

    input_path = args.input
    output_dir = args.output
    if not input_path:
        input_path = input("输入txt文件夹路径: ").strip()
    if not output_dir:
        output_dir = input("输出EPUB文件夹路径: ").strip()

    if os.path.isdir(input_path):
        for name in os.listdir(input_path):
            if name.lower().endswith(".txt"):
                convert_txt_to_epub(os.path.join(input_path, name), output_dir)
    else:
        convert_txt_to_epub(input_path, output_dir, args.title, args.author or "")


if __name__ == "__main__":
    main()