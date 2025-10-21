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
import shutil
from datetime import datetime, timezone
from typing import List, Tuple

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# Default folders used after conversion
HISTORY_DIR = r"C:\Users\User\Desktop\txt to epub\history"
FINAL_EPUB_DIR = r"C:\Users\User\iCloudDrive\Downloads\中文小说"

# -------------------- Configuration --------------------
# Users may tweak the following constants directly instead of
# providing command line arguments.
MAX_CHAPTER_CHARS = 30000
FALLBACK_SPLIT_CHARS = 12000
KEEP_BLANK_LINES = True
COVER_FILE = None  # path to image file (jpg/png) if provided
DRY_RUN = False
MOVE_TO_HISTORY = True
MAX_TITLE_LEN = 50
# --------------------------------------------------------


NBSP = "\u00A0"

NOISE_PATTERNS = [re.compile(p, re.IGNORECASE) for p in [
    r"手机访问",
    r"请记住本站域名",
    r"（未完待续）",
]]

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


ENTITY_RE = re.compile(r"&(?:amp;)?(nbsp|ldquo|rdquo|lsquo|rsquo|mdash);")
ENTITY_MAP = {
    "nbsp": NBSP,
    "ldquo": "“",
    "rdquo": "”",
    "lsquo": "‘",
    "rsquo": "’",
    "mdash": "—",
}


def fix_named_entities(s: str) -> str:
    return ENTITY_RE.sub(lambda m: ENTITY_MAP.get(m.group(1), m.group(0)), s)


DEFAULT_PATTERN_STRS = [
    r"第[0-9零一二三四五六七八九十百千万〇两]+(?:[卷部集])?(?:第[0-9零一二三四五六七八九十百千万〇两]+)?[章节回篇节话]",
    r"[卷部集][0-9零一二三四五六七八九十百千万〇两]+第[0-9零一二三四五六七八九十百千万〇两]+[章节回篇节话]",
    r"楔子",
    r"序章?",
    r"番外.*",
    r"后记.*",
    r"前言.*",
    r"Chapter\s*\d+.*",
]

def compile_patterns(extra: List[str]) -> List[re.Pattern]:
    patterns = [re.compile(p, re.IGNORECASE) for p in DEFAULT_PATTERN_STRS]
    for pat in extra:
        try:
            patterns.append(re.compile(pat, re.IGNORECASE))
        except re.error as e:
            logging.warning(f"Invalid pattern '{pat}': {e}")
    return patterns


def is_chapter_heading(line: str, patterns: List[re.Pattern]) -> bool:
    """Return True if *line* looks like a chapter heading."""
    text = clean_text(line).strip()
    if not text or len(text) > MAX_TITLE_LEN:
        return False

    # convert full-width digits to half-width
    text = "".join(
        chr(ord(ch) - 0xFEE0) if "０" <= ch <= "９" else ch for ch in text
    )
    # strip leading punctuation to allow patterns like "【第1章】"
    text = re.sub(r"^[^\w\u4e00-\u9fff]+", "", text)
    norm = re.sub(r"[\s:：.-]+", "", text)
    for pat in patterns:
        if pat.match(norm):
            return True
    return False


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


def detect_title_author(file_path: str, max_lines: int = 10) -> Tuple[str, str]:
    """Try to detect book title and author from filename or the first few lines."""
    base = os.path.splitext(os.path.basename(file_path))[0]
    title = ''
    author = ''

    m = re.match(r"《(.+?)》(.+)", base)
    if m:
        title, author = m.group(1).strip(), m.group(2).strip()
    else:
        m = re.match(r"(.+?)[-_](.+)", base)
        if m:
            title, author = m.group(1).strip(), m.group(2).strip()

    enc = detect_encoding(file_path)
    try:
        with open(file_path, 'r', encoding=enc, errors='ignore') as f:
            lines = [clean_text(f.readline()).strip() for _ in range(max_lines)]
    except Exception:
        lines = []

    for line in lines:
        if not title:
            m = re.search(r"《([^》]+)》", line)
            if m:
                title = m.group(1).strip()
        if not author:
            m = re.search(r"作者[:：]?\s*(\S+)", line)
            if m:
                author = m.group(1).strip()
        if title and author:
            break

    if not title:
        title = base
    return title, author


def safe_move(src: str, dst_dir: str, *, overwrite: bool = False) -> str:
    """Move *src* into *dst_dir*.

    If *overwrite* is True and a file with the same name exists in
    *dst_dir*, it will be replaced. Otherwise a numerical suffix is added
    to avoid conflicts."""
    os.makedirs(dst_dir, exist_ok=True)
    base = os.path.basename(src)
    dst = os.path.join(dst_dir, base)
    if overwrite:
        if os.path.exists(dst):
            os.remove(dst)
    else:
        name, ext = os.path.splitext(base)
        count = 1
        while os.path.exists(dst):
            dst = os.path.join(dst_dir, f"{name}_{count}{ext}")
            count += 1
    shutil.move(src, dst)
    return dst


def parse_chapters(file_path: str, patterns: List[re.Pattern]) -> List[Tuple[str, str]]:
    enc = detect_encoding(file_path)
    chapters: List[Tuple[str, str]] = []
    title = None
    buffer: List[str] = []
    with open(file_path, 'r', encoding=enc, errors='ignore') as f:
        for raw in f:
            line = clean_text(raw.rstrip('\n'))
            if is_chapter_heading(line, patterns):
                if title or buffer:
                    chapters.append((title or '前言', '\n'.join(buffer)))
                title = line
                buffer = []
            else:
                buffer.append(line)
    if title or buffer:
        chapters.append((title or '正文', '\n'.join(buffer)))
    return chapters


def _split_long_chapter(title: str, text: str, limit: int) -> List[Tuple[str, str]]:
    segments = []
    while len(text) > limit:
        cut = text.rfind('\n', 0, limit)
        if cut == -1:
            cut = text.rfind('。', 0, limit)
        if cut == -1:
            cut = limit
        segments.append(text[:cut].strip())
        text = text[cut:].lstrip()
    segments.append(text.strip())
    if len(segments) == 1:
        return [(title, segments[0])]
    return [(f"{title}（{i}）", seg) for i, seg in enumerate(segments, 1)]


def ensure_reasonable_chapters(chapters: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    if not chapters:
        return chapters

    if (
        len(chapters) == 1
        and chapters[0][0] in ('正文', '前言')
        and len(chapters[0][1]) > FALLBACK_SPLIT_CHARS
    ):
        text = chapters[0][1]
        segs = []
        while len(text) > FALLBACK_SPLIT_CHARS:
            cut = text.rfind('\n', 0, FALLBACK_SPLIT_CHARS)
            if cut == -1:
                cut = text.rfind('。', 0, FALLBACK_SPLIT_CHARS)
            if cut == -1:
                cut = FALLBACK_SPLIT_CHARS
            segs.append(text[:cut].strip())
            text = text[cut:].lstrip()
        segs.append(text.strip())
        return [(f"第{i}章", seg) for i, seg in enumerate(segs, 1)]

    result: List[Tuple[str, str]] = []
    for title, text in chapters:
        if len(text) > MAX_CHAPTER_CHARS:
            result.extend(_split_long_chapter(title, text, MAX_CHAPTER_CHARS))
        else:
            result.append((title, text))
    return result


def _normalize_paragraphs(text: str) -> List[str]:
    lines = text.splitlines()
    result: List[str] = []
    for line in lines:
        line = clean_text(line).strip()
        if any(pat.search(line) for pat in NOISE_PATTERNS):
            continue
        if not line:
            if KEEP_BLANK_LINES and (not result or result[-1] != ''):
                result.append('')
            elif not KEEP_BLANK_LINES:
                continue
        else:
            result.append(line)
    return result

def chapter_to_xhtml(idx: int, title: str, text: str) -> str:
    # 标题：先清洗→替换命名实体→再 HTML 转义
    title_clean = clean_text(title)
    title_fixed = fix_named_entities(title_clean)
    esc_title = html.escape(title_fixed)

    paras = []
    for p in _normalize_paragraphs(text):
        if p:
            # 正文段落：同样先清洗→替换命名实体→再转义
            body_clean = clean_text(p)
            body_fixed = fix_named_entities(body_clean)
            paras.append(f"    <p>{html.escape(body_fixed)}</p>")
        else:
            # 空行占位：用 NBSP（U+00A0 或 &#160;），不要再 escape
            paras.append(f"    <p class='blank'>{NBSP}</p>")

    return (
        "<?xml version='1.0' encoding='utf-8'?>\n"
        "<html xmlns='http://www.w3.org/1999/xhtml' xml:lang='zh' lang='zh'>\n"
        "<head>\n"
        "  <meta charset='utf-8'/>\n"
        f"  <title>{esc_title}</title>\n"
        "  <link rel='stylesheet' type='text/css' href='style.css'/>\n"
        "</head>\n<body>\n"
        f"  <h2 id='chap{idx}'>{esc_title}</h2>\n" +
        "\n".join(paras) + "\n</body>\n</html>"
    )



def create_epub(title: str, author: str, chapters: List[Tuple[str, str]], out_path: str, lang: str = 'zh'):
    tmp_path = out_path + '.tmp'
    uid = str(uuid.uuid4())
    modified = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    css = (
        "body{margin:0 0.8em;font-family:'PingFang SC','Hiragino Sans GB','Heiti SC','STSong','Songti SC',serif;"
        "line-height:1.8;text-align:justify;text-justify:inter-ideograph;}"
        "p{margin:0.6em 0;text-indent:2em;}"
        "img{max-width:100%;height:auto;display:block;margin:0 auto;}"
        ".blank{text-indent:0;height:0.9em;}"
    )

    cover_info = None
    if COVER_FILE and os.path.isfile(COVER_FILE):
        ext = os.path.splitext(COVER_FILE)[1].lower()
        if ext in ('.jpg', '.jpeg', '.png'):
            media_type = 'image/jpeg' if ext in ('.jpg', '.jpeg') else 'image/png'
            cover_info = (ext, media_type)

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
        nav_points = []

        if cover_info:
            ext, media_type = cover_info
            cover_name = 'cover' + ext
            with open(COVER_FILE, 'rb') as cf:
                epub.writestr(f'OEBPS/{cover_name}', cf.read())
            manifest.append(
                f"<item id='coverimg' href='{cover_name}' media-type='{media_type}' properties='cover-image'/>"
            )
            cover_xhtml = (
                f"<?xml version='1.0' encoding='utf-8'?>\n<html xmlns='http://www.w3.org/1999/xhtml' xml:lang='{lang}' lang='{lang}'>\n"
                "<head>\n  <meta charset='utf-8'/><title>封面</title><link rel='stylesheet' type='text/css' href='style.css'/></head>\n"
                f"<body>\n  <img src='{cover_name}' alt='cover'/>\n</body>\n</html>"
            )
        else:
            title_clean = clean_text(title)
            author_clean = clean_text(author)
            esc_title = html.escape(fix_named_entities(title_clean))
            esc_author = html.escape(fix_named_entities(author_clean))
            cover_xhtml = (
                f"<?xml version='1.0' encoding='utf-8'?>\n<html xmlns='http://www.w3.org/1999/xhtml' xml:lang='{lang}' lang='{lang}'>\n"
                "<head>\n  <meta charset='utf-8'/><title>封面</title><link rel='stylesheet' type='text/css' href='style.css'/></head>\n"
                f"<body>\n  <h1>{esc_title}</h1>\n  <p>{esc_author}</p>\n</body>\n</html>"
            )
        epub.writestr('OEBPS/cover.xhtml', cover_xhtml)
        manifest.append("<item id='cover' href='cover.xhtml' media-type='application/xhtml+xml'/>")
        spine.append("<itemref idref='cover' linear='yes'/>")
        spine.append("<itemref idref='nav' linear='no'/>")

        for i, (ch_title, ch_text) in enumerate(chapters, 1):
            fname = f'chapter{i}.xhtml'
            epub.writestr(f'OEBPS/{fname}', chapter_to_xhtml(i, ch_title, ch_text))
            manifest.append(f"<item id='c{i}' href='{fname}' media-type='application/xhtml+xml'/>")
            spine.append(f"<itemref idref='c{i}'/>")
            title_clean = clean_text(ch_title)
            esc_title = html.escape(fix_named_entities(title_clean))
            nav_list.append(f"      <li><a href='{fname}#chap{i}'>{esc_title}</a></li>")
            nav_points.append(
                f"    <navPoint id='navPoint-{i}' playOrder='{i}'>\n      <navLabel><text>{esc_title}</text></navLabel>\n      <content src='{fname}#chap{i}'/>\n    </navPoint>"
            )

        epub.writestr(
            'OEBPS/nav.xhtml',
            f"""<?xml version='1.0' encoding='utf-8'?>
<html xmlns='http://www.w3.org/1999/xhtml' xmlns:epub='http://www.idpf.org/2007/ops' xml:lang='{lang}' lang='{lang}'>
<head><meta charset='utf-8'/><title>目录</title><link rel='stylesheet' type='text/css' href='style.css'/></head>
<body><nav epub:type='toc' id='toc'><h1>目录</h1><ol>
"""
            + ''.join(nav_list)
            + "\n</ol></nav></body></html>",
        )

        epub.writestr(
            'OEBPS/toc.ncx',
            """<?xml version='1.0' encoding='utf-8'?>
<!DOCTYPE ncx PUBLIC '-//NISO//DTD ncx 2005-1//EN' 'http://www.daisy.org/z3986/2005/ncx-2005-1.dtd'>
<ncx xmlns='http://www.daisy.org/z3986/2005/ncx/' version='2005-1'>
  <head>
    <meta name='dtb:uid' content='"""
            + uid
            + "'/>\n    <meta name='dtb:depth' content='1'/>\n    <meta name='dtb:totalPageCount' content='0'/>\n    <meta name='dtb:maxPageNumber' content='0'/>\n  </head>\n  <docTitle><text>"
            + html.escape(fix_named_entities(clean_text(title)))
            + "</text></docTitle>\n  <navMap>\n"
            + '\n'.join(nav_points)
            + "\n  </navMap>\n</ncx>",
        )

        manifest.append("<item id='ncx' href='toc.ncx' media-type='application/x-dtbncx+xml'/>")
        manifest_str = '\n    '.join(manifest)
        spine_str = '\n    '.join(spine)

        # EPUB2 compatibility bits for wider reader support
        extra_meta = "<meta name='cover' content='coverimg'/>" if cover_info else ""
        guide = (
            "<guide>\n"
            "    <reference type='cover' title='Cover' href='cover.xhtml'/>\n"
            "    <reference type='toc' title='Table of Contents' href='nav.xhtml'/>\n"
            "  </guide>"
        )

        opf = (
            f"""<?xml version='1.0' encoding='utf-8'?>
<package xmlns='http://www.idpf.org/2007/opf' unique-identifier='bookid' version='3.0'>
  <metadata xmlns:dc='http://purl.org/dc/elements/1.1/'>
    <dc:identifier id='bookid'>{uid}</dc:identifier>
    <dc:title>{html.escape(fix_named_entities(clean_text(title)))}</dc:title>
    <dc:creator>{html.escape(fix_named_entities(clean_text(author)))}</dc:creator>
    <dc:language>{lang}</dc:language>
    <meta property='dcterms:modified'>{modified}</meta>
    {extra_meta}
  </metadata>
  <manifest>
    {manifest_str}
  </manifest>
  <spine toc='ncx'>
    {spine_str}
  </spine>
  {guide}
</package>"""
        )
        epub.writestr('OEBPS/content.opf', opf)
    os.replace(tmp_path, out_path)
    logging.info(f"EPUB generated: {out_path}")


def convert_txt_file(
    in_file: str,
    out_dir: str,
    lang: str,
    patterns: List[re.Pattern],
    history_dir: str,
    final_dir: str,
):
    if not os.path.isfile(in_file):
        logging.error(f"File not found: {in_file}")
        return
    os.makedirs(out_dir, exist_ok=True)
    title, author = detect_title_author(in_file)
    chapters = parse_chapters(in_file, patterns)
    chapters = ensure_reasonable_chapters(chapters)
    if DRY_RUN:
        print(f"[DRY RUN] Chapters detected: {len(chapters)}")
        for i, (ch_title, _) in enumerate(chapters[:5], 1):
            print(f"  {i}. {ch_title}")
        return
    base = os.path.splitext(os.path.basename(in_file))[0]
    out_file = os.path.join(out_dir, f"{base}.epub")
    create_epub(title, author, chapters, out_file, lang)
    if final_dir:
        final_path = safe_move(out_file, final_dir, overwrite=True)
        logging.info(f"EPUB moved to {final_path}")
    if MOVE_TO_HISTORY and history_dir:
        hist_path = safe_move(in_file, history_dir)
        logging.info(f"Source moved to {hist_path}")


def batch_convert(
    input_dir: str,
    output_dir: str,
    lang: str,
    patterns: List[re.Pattern],
    history_dir: str,
    final_dir: str,
):
    if not os.path.isdir(input_dir):
        logging.error(f"Invalid input directory: {input_dir}")
        return
    texts = [f for f in os.listdir(input_dir) if f.lower().endswith('.txt')]
    if not texts:
        logging.info("No .txt files found in input directory")
        return
    for name in texts:
        convert_txt_file(
            os.path.join(input_dir, name),
            output_dir,
            lang,
            patterns,
            history_dir,
            final_dir,
        )


def main():
    default_in = r'C:\Users\User\Desktop\txt to epub\txt file'
    default_out = r'C:\Users\User\Desktop\txt to epub\epub file'
    parser = argparse.ArgumentParser(description="Batch convert Chinese txt to EPUB")
    parser.add_argument('-i', '--input', default=default_in, help='input directory')
    parser.add_argument('-o', '--output', default=default_out, help='temporary output directory')
    parser.add_argument('--history', default=HISTORY_DIR, help='directory to move processed txt files')
    parser.add_argument('--dest', default=FINAL_EPUB_DIR, help='directory to move generated EPUB files')
    parser.add_argument('--lang', default='zh', help='language code')
    parser.add_argument('-p', '--pattern', action='append', default=[], help='additional chapter regex, can be used multiple times')
    args = parser.parse_args()
    patterns = compile_patterns(args.pattern)
    logging.info(
        "Config: MAX_CHAPTER_CHARS=%s FALLBACK_SPLIT_CHARS=%s KEEP_BLANK_LINES=%s COVER_FILE=%s DRY_RUN=%s MOVE_TO_HISTORY=%s MAX_TITLE_LEN=%s",
        MAX_CHAPTER_CHARS,
        FALLBACK_SPLIT_CHARS,
        KEEP_BLANK_LINES,
        COVER_FILE,
        DRY_RUN,
        MOVE_TO_HISTORY,
        MAX_TITLE_LEN,
    )
    batch_convert(args.input, args.output, args.lang, patterns, args.history, args.dest)

if __name__ == '__main__':
    main()
