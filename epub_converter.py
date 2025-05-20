#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse
import os
import re
import uuid
import zipfile
import functools
import logging
import html
import codecs
from datetime import datetime, timezone
from typing import List, Tuple

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

@functools.lru_cache(maxsize=None)
def detect_encoding(file_path: str, sample_size: int = 4096) -> str:
    with open(file_path, 'rb') as f:
        raw = f.read(sample_size)

    boms = {
        codecs.BOM_UTF8: 'utf-8-sig',
        codecs.BOM_UTF16_LE: 'utf-16-le',
        codecs.BOM_UTF16_BE: 'utf-16-be',
        codecs.BOM_UTF32_LE: 'utf-32-le',
        codecs.BOM_UTF32_BE: 'utf-32-be',
    }
    for bom, enc in boms.items():
        if raw.startswith(bom):
            return enc

    if len(raw) >= 4:
        null_even = raw[::2].count(0)
        null_odd = raw[1::2].count(0)
        if null_even > len(raw) * 0.3 and null_odd < len(raw) * 0.05:
            return 'utf-16-le'
        if null_odd > len(raw) * 0.3 and null_even < len(raw) * 0.05:
            return 'utf-16-be'

    for enc in ('utf-8', 'gb18030', 'gbk', 'big5'):
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue

    try:
        import chardet
        guess = chardet.detect(raw)
        if guess and guess['confidence'] > 0.5 and guess['encoding']:
            return guess['encoding']
    except ImportError:
        pass

    logging.warning(f"{file_path} encoding not detected, falling back to utf-8 (ignore)")
    return 'utf-8'


def clean_text(text: str) -> str:
    allowed_chars = []
    for ch in text:
        code = ord(ch)
        if ch in ('\t', '\n', '\r', ' ') or 0x20 <= code <= 0xD7FF or 0xE000 <= code <= 0xFFFD:
            allowed_chars.append(ch)
    return ''.join(allowed_chars)


def is_chapter_heading(line: str) -> bool:
    """Return True if *line* looks like a chapter heading."""
    text = clean_text(line).strip()
    if not text or len(text) > 50:
        return False

    norm = re.sub(r"[\s:：.-]+", "", text)
    if norm in ("序", "序章", "楔子"):
        return True

    pat1 = (
        r"^第[0-9零一二三四五六七八九十百千万〇两]+(?:卷|季|集|部|册)?"
        r"(?:第[0-9零一二三四五六七八九十百千万〇两]+)?(?:章|回|篇|节|话).*"
    )
    pat2 = r"^[0-9一二三四五六七八九十百千万〇两]{1,4}[、.．]\S+"
    return bool(re.match(pat1, norm) or re.match(pat2, text))


def detect_author(file_path: str, max_lines: int = 20) -> str:
    enc = detect_encoding(file_path)
    pat = re.compile(r"作者[:：]\s*(.+)")
    with open(file_path, 'r', encoding=enc, errors='ignore') as f:
        for _ in range(max_lines):
            line = clean_text(f.readline())
            if not line:
                break
            m = pat.search(line)
            if m:
                return m.group(1).strip()
    return ''


def parse_chapters(file_path: str) -> List[Tuple[str, str]]:
    enc = detect_encoding(file_path)
    chapters: List[Tuple[str, str]] = []
    title = None
    buffer: List[str] = []
    with open(file_path, 'r', encoding=enc, errors='ignore') as f:
        for raw in f:
            line = clean_text(raw.rstrip('\n'))
            if is_chapter_heading(line):
                if title or buffer:
                    chapters.append((title or '前言', '\n'.join(buffer)))
                title = line
                buffer = []
            else:
                buffer.append(line)
    if title or buffer:
        chapters.append((title or '正文', '\n'.join(buffer)))
    return chapters


def chapter_to_xhtml(idx: int, title: str, text: str) -> str:
    esc_title = html.escape(clean_text(title))
    paras = [f"    <p>{html.escape(clean_text(p.strip()))}</p>" for p in text.splitlines() if p.strip()]
    return (
        "<?xml version='1.0' encoding='utf-8'?>\n"
        "<html xmlns='http://www.w3.org/1999/xhtml'>\n<head>\n  <title>" + esc_title + "</title>\n  <link rel='stylesheet' type='text/css' href='style.css'/>\n</head>\n<body>\n  <h2 id='chap" + str(idx) + "'>" + esc_title + "</h2>\n" +
        '\n'.join(paras) + "\n</body>\n</html>"
    )


def create_epub(title: str, author: str, chapters: List[Tuple[str, str]], out_path: str, lang: str = 'zh'):
    tmp_path = out_path + '.tmp'
    uid = str(uuid.uuid4())
    modified = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    css = 'body { font-family: SimSun, serif; line-height:1.5; text-indent:2em; }'

    with zipfile.ZipFile(tmp_path, 'w') as epub:
        epub.writestr('mimetype', 'application/epub+zip', compress_type=zipfile.ZIP_STORED)
        epub.writestr(
            'META-INF/container.xml',
            """<?xml version='1.0' encoding='UTF-8'?>
<container version='1.0' xmlns='urn:oasis:names:tc:opendocument:xmlns:container'>
  <rootfiles>
    <rootfile full-path='OEBPS/content.opf' media-type='application/oebps-package+xml'/>
  </rootfiles>
</container>""",
        )
        epub.writestr('OEBPS/style.css', css)

        manifest = [
            "<item id='nav' href='nav.xhtml' properties='nav' media-type='application/xhtml+xml'/>",
            "<item id='css' href='style.css' media-type='text/css'/>",
        ]
        spine = []
        nav_list = []

        for i, (ch_title, ch_text) in enumerate(chapters, 1):
            fname = f'chapter{i}.xhtml'
            epub.writestr(f'OEBPS/{fname}', chapter_to_xhtml(i, ch_title, ch_text))
            manifest.append(f"<item id='c{i}' href='{fname}' media-type='application/xhtml+xml'/>")
            spine.append(f"<itemref idref='c{i}'/>")
            nav_list.append(f"      <li><a href='{fname}#chap{i}'>{html.escape(ch_title)}</a></li>")

        epub.writestr(
            'OEBPS/nav.xhtml',
            """<?xml version='1.0' encoding='utf-8'?>
<html xmlns='http://www.w3.org/1999/xhtml'>
<head><title>目录</title><link rel='stylesheet' type='text/css' href='style.css'/></head>
<body><nav epub:type='toc' id='toc'><h1>目录</h1><ol>
"""
            + ''.join(nav_list)
            + "\n</ol></nav></body></html>",
        )

        nav_points = [
            f"    <navPoint id='navPoint-{i}' playOrder='{i}'>\n      <navLabel><text>{html.escape(ch_title)}</text></navLabel>\n      <content src='chapter{i}.xhtml'/>\n    </navPoint>"
            for i, (ch_title, _) in enumerate(chapters, 1)
        ]
        epub.writestr(
            'OEBPS/toc.ncx',
            """<?xml version='1.0' encoding='utf-8'?>
<!DOCTYPE ncx PUBLIC '-//NISO//DTD ncx 2005-1//EN' 'http://www.daisy.org/z3986/2005/ncx-2005-1.dtd'>
<ncx xmlns='http://www.daisy.org/z3986/2005/ncx/' version='2005-1'>
  <head>
    <meta name='dtb:uid' content='"""
            + uid
            + "'/>\n    <meta name='dtb:depth' content='1'/>\n    <meta name='dtb:totalPageCount' content='0'/>\n    <meta name='dtb:maxPageNumber' content='0'/>\n  </head>\n  <docTitle><text>"
            + html.escape(title)
            + "</text></docTitle>\n  <navMap>\n"
            + '\n'.join(nav_points)
            + "\n  </navMap>\n</ncx>",
        )

        manifest.append("<item id='ncx' href='toc.ncx' media-type='application/x-dtbncx+xml'/>")
        manifest_str = '\n    '.join(manifest)
        spine_str = '\n    '.join(spine)

        opf = (
            f"""<?xml version='1.0' encoding='utf-8'?>
<package xmlns='http://www.idpf.org/2007/opf' unique-identifier='bookid' version='3.0'>
  <metadata xmlns:dc='http://purl.org/dc/elements/1.1/'>
    <dc:identifier id='bookid'>{uid}</dc:identifier>
    <dc:title>{html.escape(title)}</dc:title>
    <dc:creator>{html.escape(author)}</dc:creator>
    <dc:language>{lang}</dc:language>
    <meta property='dcterms:modified'>{modified}</meta>
  </metadata>
  <manifest>
    {manifest_str}
  </manifest>
  <spine toc='ncx'>
    {spine_str}
  </spine>
</package>"""
        )
        epub.writestr('OEBPS/content.opf', opf)
    os.replace(tmp_path, out_path)
    logging.info(f"EPUB generated: {out_path}")


def convert_txt_file(in_file: str, out_dir: str, lang: str):
    if not os.path.isfile(in_file):
        logging.error(f"File not found: {in_file}")
        return
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(in_file))[0]
    author = detect_author(in_file)
    chapters = parse_chapters(in_file)
    out_file = os.path.join(out_dir, f"{base}.epub")
    create_epub(base, author, chapters, out_file, lang)


def batch_convert(input_dir: str, output_dir: str, lang: str):
    if not os.path.isdir(input_dir):
        logging.error(f"Invalid input directory: {input_dir}")
        return
    texts = [f for f in os.listdir(input_dir) if f.lower().endswith('.txt')]
    if not texts:
        logging.info("No .txt files found in input directory")
        return
    for name in texts:
        convert_txt_file(os.path.join(input_dir, name), output_dir, lang)


def main():
    default_in = r'C:\Users\User\Desktop\txt to epub\txt file'
    default_out = r'C:\Users\User\Desktop\txt to epub\epub file'
    parser = argparse.ArgumentParser(description="Batch convert Chinese txt to EPUB")
    parser.add_argument('-i', '--input', default=default_in, help='input directory')
    parser.add_argument('-o', '--output', default=default_out, help='output directory')
    parser.add_argument('--lang', default='zh', help='language code')
    args = parser.parse_args()
    batch_convert(args.input, args.output, args.lang)

if __name__ == '__main__':
    main()
