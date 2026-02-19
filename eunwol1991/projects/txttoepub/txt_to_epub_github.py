#!/usr/bin/env python3
# -*- coding: utf-8 -*-


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
"""
Simple TXT → EPUB converter.

This script scans the configured input directory, converts every UTF/GB encoded
TXT file into an EPUB with a plain-text cover, then moves the EPUB and source
TXT into the configured destination folders.
"""

import codecs
import html
import logging
import os
import re
import shutil
import uuid
import zipfile
from datetime import datetime, timezone
from typing import List, Pattern, Tuple

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")

# Configuration – change these paths if needed.
INPUT_DIR = _from_c("Users/jhunj/Desktop/txt to epub/txt file")
TEMP_DIR = _from_c("Users/jhunj/Desktop/txt to epub/epub file")
HISTORY_DIR = _from_c("Users/jhunj/Desktop/txt to epub/history")
FINAL_EPUB_DIR = _from_c("Users/jhunj/iCloudDrive/Downloads/中文小说")
MAX_TITLE_LEN = 50
FALLBACK_SPLIT_CHARS = 12000
CHAPTER_PATTERN_STRS = [
    r"^楔子$",
    r"^序章$",
    r"^引子$",
    r"^尾声$",
    r"^后记$",
    r"^番外",
    r"^终章$",
    r"^第[\d零一二三四五六七八九十百千万〇两]+[章节卷回篇部]",
    r"^[\d零一二三四五六七八九零两]+回",
    r"^Chapter\s*\d+",
]
NOISE_PATTERNS = [re.compile(p) for p in (r"手机访问", r"请记住本站域名", r"未完待续")]
SENTENCE_END_RE = re.compile(r"[。！？\?]+")


def detect_encoding(file_path: str, sample_size: int = 4096) -> str:
    with open(file_path, "rb") as f:
        sample = f.read(sample_size)
    boms = {
        codecs.BOM_UTF8: "utf-8-sig",
        codecs.BOM_UTF16_LE: "utf-16-le",
        codecs.BOM_UTF16_BE: "utf-16-be",
    }
    for bom, enc in boms.items():
        if sample.startswith(bom):
            return enc
    for enc in ("utf-8", "gb18030", "gbk", "big5"):
        try:
            sample.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    try:
        import chardet

        guess = chardet.detect(sample)
        if guess and guess["encoding"]:
            return guess["encoding"]
    except ImportError:
        pass
    logging.warning(
        "Encoding detection failed for %s, falling back to gb18030", file_path)
    return "gb18030"


def clean_text(line: str) -> str:
    return line.strip()


def compile_patterns() -> List[Pattern]:
    return [re.compile(p, re.IGNORECASE) for p in CHAPTER_PATTERN_STRS]


def is_chapter_heading(line: str, patterns: List[Pattern]) -> bool:
    text = clean_text(line)
    if not text or len(text) > MAX_TITLE_LEN:
        return False
    simple = re.sub(r"^[^\w\u4e00-\u9fff]+", "", text)
    simple = re.sub(r"[\s：:【】]+", "", simple)
    for pat in patterns:
        if pat.match(simple):
            return True
    return False


def parse_chapters(file_path: str, patterns: List[Pattern]) -> List[Tuple[str, str]]:
    enc = detect_encoding(file_path)
    chapters: List[Tuple[str, str]] = []
    title = None
    buffer: List[str] = []
    with open(file_path, "r", encoding=enc, errors="replace") as f:
        for line in f:
            line = clean_text(line)
            if is_chapter_heading(line, patterns):
                if title or buffer:
                    chapters.append((title or "前言", "\n".join(buffer).strip()))
                title = line
                buffer = []
            else:
                buffer.append(line)
    if title or buffer:
        chapters.append((title or "正文", "\n".join(buffer).strip()))
    if not chapters:
        return [("正文", open(file_path, "r", encoding=enc, errors="replace").read())]
    return chapters


def ensure_chapter_text(chapters: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    if len(chapters) == 1 and len(chapters[0][1]) > FALLBACK_SPLIT_CHARS:
        text = chapters[0][1]
        segs: List[str] = []
        while len(text) > FALLBACK_SPLIT_CHARS:
            cut = text.rfind("\n", 0, FALLBACK_SPLIT_CHARS)
            if cut == -1:
                cut = FALLBACK_SPLIT_CHARS
            segs.append(text[:cut].strip())
            text = text[cut:].strip()
        segs.append(text.strip())
        return [(f"第{i}章", seg) for i, seg in enumerate(segs, 1)]
    return chapters


def _split_sentences(chunk: str) -> List[str]:
    parts = []
    start = 0
    for match in SENTENCE_END_RE.finditer(chunk):
        end = match.end()
        part = chunk[start:end].strip()
        if part:
            parts.append(part)
        start = end
    tail = chunk[start:].strip()
    if tail:
        parts.append(tail)
    return parts


def _paragraphs(text: str) -> List[str]:
    paras: List[str] = []
    for block in re.split(r"\n\s*\n", text):
        block = block.strip()
        if not block:
            continue
        if any(p.search(block) for p in NOISE_PATTERNS):
            continue
        chunk = block.replace("\n", " ")
        paras.extend(_split_sentences(chunk))
    return paras


def chapter_to_xhtml(idx: int, title: str, text: str) -> str:
    esc_title = html.escape(title)
    body = "\n".join(f"    <p>{html.escape(p)}</p>" for p in _paragraphs(text))
    return (
        "<?xml version='1.0' encoding='utf-8'?>\n"
        "<html xmlns='http://www.w3.org/1999/xhtml'>\n"
        "<head>\n"
        "  <meta charset='utf-8'/>\n"
        f"  <title>{esc_title}</title>\n"
        "  <link rel='stylesheet' href='style.css'/>\n"
        "</head>\n"
        "<body>\n"
        f"  <h2 id='chap{idx}'>{esc_title}</h2>\n"
        f"{body}\n"
        "</body>\n"
        "</html>"
    )


def create_epub(title: str, author: str, chapters: List[Tuple[str, str]], out_path: str) -> None:
    tmp_path = out_path + ".tmp"
    uid = str(uuid.uuid4())
    modified = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    css = (
        'body{margin:0 1em;font-family:"Noto Serif SC","Source Han Serif SC","Songti SC","STSong",serif;'
        "line-height:1.8;text-indent:2em;font-size:1rem;text-align:justify;text-justify:inter-ideograph;}"
        "p{margin:0.8em 0;}h1{text-align:center;margin-top:4em;}"
    )

    nav_items = []
    manifest = [
        "<item id='nav' href='nav.xhtml' properties='nav' media-type='application/xhtml+xml'/>",
        "<item id='css' href='style.css' media-type='text/css'/>",
    ]
    spine = ["<itemref idref='cover' linear='yes'/>",
             "<itemref idref='nav' linear='no'/>"]

    with zipfile.ZipFile(tmp_path, "w") as epub:
        epub.writestr("mimetype", "application/epub+zip",
                      compress_type=zipfile.ZIP_STORED)
        epub.writestr("META-INF/container.xml", """<?xml version='1.0' encoding='UTF-8'?>
<container version='1.0' xmlns='urn:oasis:names:tc:opendocument:xmlns:container'>
  <rootfiles>
    <rootfile full-path='OEBPS/content.opf' media-type='application/oebps-package+xml'/>
  </rootfiles>
</container>""")
        epub.writestr("OEBPS/style.css", css)

        cover_xhtml = (
            "<?xml version='1.0' encoding='utf-8'?>\n"
            "<html xmlns='http://www.w3.org/1999/xhtml'>\n"
            "<head><meta charset='utf-8'/><title>封面</title><link rel='stylesheet' href='style.css'/></head>\n"
            "<body>\n"
            f"  <h1>{html.escape(title)}</h1>\n"
            f"  <p>{html.escape(author)}</p>\n"
            "</body>\n"
            "</html>"
        )
        epub.writestr("OEBPS/cover.xhtml", cover_xhtml)
        manifest.append(
            "<item id='cover' href='cover.xhtml' media-type='application/xhtml+xml'/>")

        nav_list = []
        for idx, (ch_title, ch_text) in enumerate(chapters, 1):
            fname = f"chap{idx}.xhtml"
            epub.writestr(f"OEBPS/{fname}",
                          chapter_to_xhtml(idx, ch_title, ch_text))
            manifest.append(
                f"<item id='c{idx}' href='{fname}' media-type='application/xhtml+xml'/>"
            )
            spine.append(f"<itemref idref='c{idx}'/>")
            nav_list.append(
                f"    <li><a href='{fname}#chap{idx}'>{html.escape(ch_title)}</a></li>")
            nav_items.append(
                f"    <navPoint id='navPoint-{idx}' playOrder='{idx}'>"
                f"<navLabel><text>{html.escape(ch_title)}</text></navLabel>"
                f"<content src='{fname}#chap{idx}'/></navPoint>"
            )

        epub.writestr(
            "OEBPS/nav.xhtml",
            "<?xml version='1.0' encoding='utf-8'?>\n"
            "<html xmlns='http://www.w3.org/1999/xhtml' xmlns:epub='http://www.idpf.org/2007/ops'>\n"
            "<head><meta charset='utf-8'/><title>目录</title><link rel='stylesheet' href='style.css'/></head>\n"
            "<body><nav epub:type='toc'><h1>目录</h1><ol>\n"
            + "\n".join(nav_list)
            + "\n</ol></nav></body></html>"
        )

        epub.writestr(
            "OEBPS/toc.ncx",
            "<?xml version='1.0' encoding='utf-8'?>\n"
            "<!DOCTYPE ncx PUBLIC '-//NISO//DTD ncx 2005-1//EN' "
            "'http://www.daisy.org/z3986/2005/ncx-2005-1.dtd'>\n"
            "<ncx xmlns='http://www.daisy.org/z3986/2005/ncx/' version='2005-1'>\n"
            "  <head>\n"
            f"    <meta name='dtb:uid' content='{uid}'/>\n"
            "    <meta name='dtb:depth' content='1'/>\n"
            "    <meta name='dtb:totalPageCount' content='0'/>\n"
            "    <meta name='dtb:maxPageNumber' content='0'/>\n"
            "  </head>\n"
            f"  <docTitle><text>{html.escape(title)}</text></docTitle>\n"
            "  <navMap>\n"
            + "\n".join(nav_items)
            + "\n  </navMap>\n"
            "</ncx>"
        )
        manifest.append(
            "<item id='ncx' href='toc.ncx' media-type='application/x-dtbncx+xml'/>")

        manifest_str = "\n    ".join(manifest)
        spine_str = "\n    ".join(spine)

        opf = (
            "<?xml version='1.0' encoding='utf-8'?>\n"
            "<package xmlns='http://www.idpf.org/2007/opf' unique-identifier='bookid' version='3.0'>\n"
            "  <metadata xmlns:dc='http://purl.org/dc/elements/1.1/'>\n"
            f"    <dc:identifier id='bookid'>{uid}</dc:identifier>\n"
            f"    <dc:title>{html.escape(title)}</dc:title>\n"
            f"    <dc:creator>{html.escape(author)}</dc:creator>\n"
            "    <dc:language>zh</dc:language>\n"
            f"    <meta property='dcterms:modified'>{modified}</meta>\n"
            "  </metadata>\n"
            "  <manifest>\n"
            f"    {manifest_str}\n"
            "  </manifest>\n"
            "  <spine toc='ncx'>\n"
            f"    {spine_str}\n"
            "  </spine>\n"
            "</package>"
        )
        epub.writestr("OEBPS/content.opf", opf)
    os.replace(tmp_path, out_path)


def safe_move(src: str, dst_dir: str, *, overwrite: bool = False) -> None:
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, os.path.basename(src))
    if overwrite:
        if os.path.exists(dst):
            os.remove(dst)
    else:
        if os.path.exists(dst):
            base, ext = os.path.splitext(dst)
            count = 1
            while os.path.exists(f"{base}_{count}{ext}"):
                count += 1
            dst = f"{base}_{count}{ext}"
    shutil.move(src, dst)


def convert_all() -> None:
    os.makedirs(TEMP_DIR, exist_ok=True)
    patterns = compile_patterns()
    files = sorted(f for f in os.listdir(INPUT_DIR)
                   if f.lower().endswith(".txt"))
    if not files:
        logging.info("No TXT files found under %s", INPUT_DIR)
        return
    for filename in files:
        txt_path = os.path.join(INPUT_DIR, filename)
        title = os.path.splitext(filename)[0]
        author = "未知"
        chapters = parse_chapters(txt_path, patterns)
        chapters = ensure_chapter_text(chapters)
        epub_path = os.path.join(TEMP_DIR, f"{title}.epub")
        create_epub(title, author, chapters, epub_path)
        safe_move(epub_path, FINAL_EPUB_DIR)
        safe_move(txt_path, HISTORY_DIR)
        logging.info("Converted %s → %s", filename,
                     os.path.basename(epub_path))


if __name__ == "__main__":
    convert_all()
