import os
import re
import sys
from ebooklib import epub
from bs4 import BeautifulSoup
import chardet
from PIL import Image, ImageDraw, ImageFont

def txt_to_epub(txt_file, epub_output_dir):
    # 检查文件是否存在
    if not os.path.isfile(txt_file):
        print("文件不存在：{}".format(txt_file))
        return

    # 检测文件编码
    with open(txt_file, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    confidence = result['confidence']
    print("处理文件：{}".format(txt_file))
    print("检测到的文件编码为：{}，置信度：{}".format(encoding, confidence))

    # 如果检测结果为 None，默认使用 gb18030
    if encoding is None:
        encoding = 'gb18030'
        print("无法检测到编码，默认使用 gb18030")

    # 读取 txt 文件内容
    try:
        with open(txt_file, 'r', encoding=encoding, errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print("读取文件时出错：{}".format(e))
        return

    # 从文件名中提取书名和作者
    base_name = os.path.basename(txt_file)
    match = re.match(r'《(.+?)》(.+?)\.txt$', base_name)
    if not match:
        print("文件名格式不正确，应为：《书名》作者.txt，跳过该文件。")
        return
    title, author = match.groups()

    # 创建新的 epub 书籍
    book = epub.EpubBook()
    book.set_identifier('id_{}'.format(title))
    book.set_title(title)
    book.set_language('zh')
    book.add_author(author)

    # 生成封面图片
    cover_image = generate_cover_image(title, author)
    if cover_image:
        book.set_cover("cover.jpg", cover_image)
        print("已生成并添加封面图片。")
    else:
        print("未能生成封面图片，跳过添加封面。")

    # 解析章节，假设章节标题格式为 "第X章"、"第X节"等
    pattern = r'(第[\d一二三四五六七八九十百千万零两]+[章节卷回篇集部].*?)\n'
    parts = re.split(pattern, content)
    chapters = []

    if len(parts) > 1:
        # 文本按照 [内容, 章节标题, 章节内容, ...] 的方式分割
        for i in range(1, len(parts), 2):
            chap_title = parts[i].strip()
            chap_content = parts[i + 1].strip()
            chapters.append((chap_title, chap_content))
    else:
        # 如果未能分割章节，则将全文作为一个章节
        chapters.append((title, content))

    # 添加章节到书籍
    epub_chapters = []
    for idx, (chap_title, chap_content) in enumerate(chapters):
        c = create_chapter(chap_title, chap_content, idx + 1)
        book.add_item(c)
        epub_chapters.append(c)

    # 设置书籍目录
    book.toc = (epub_chapters)

    # 添加默认的 NCX 和导航文件
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())

    # 定义 CSS 样式
    style = '''
    @namespace epub "http://www.idpf.org/2007/ops";
    body {
        font-family: "微软雅黑", Arial, sans-serif;
        line-height: 1.6;
        margin: 0;
        padding: 1em;
    }
    h1 {
        text-align: center;
        margin: 1em 0;
    }
    p {
        text-indent: 2em;
        margin: 1em 0;
    }
    '''

    # 添加 CSS 文件
    nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=style)
    book.add_item(nav_css)

    # 设置书籍的 spine（阅读顺序）
    book.spine = ['nav'] + epub_chapters

    # 生成 epub 文件
    epub_name = '{} - {}.epub'.format(title, author)
    epub_path = os.path.join(epub_output_dir, epub_name)

    # 检查输出目录是否存在，不存在则创建
    if not os.path.exists(epub_output_dir):
        os.makedirs(epub_output_dir)

    # 检查 epub 文件是否已经存在
    if os.path.isfile(epub_path):
        print("epub 文件已存在，跳过转换：{}".format(epub_name))
        return

    epub.write_epub(epub_path, book, {})

    print("成功生成电子书：{}".format(epub_path))

def create_chapter(title, content, index):
    # 创建章节内容
    soup = BeautifulSoup('', 'html.parser')

    # 添加章节标题
    h1 = soup.new_tag('h1')
    h1.string = title
    soup.append(h1)

    # 将内容按段落分割
    paragraphs = content.split('\n')
    for para in paragraphs:
        if para.strip():
            p = soup.new_tag('p')
            p.string = para.strip()
            soup.append(p)

    # 创建章节对象
    c = epub.EpubHtml(title=title, file_name='chap_{}.xhtml'.format(index), lang='zh')
    c.content = str(soup)
    c.add_item(epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css"))
    return c

def generate_cover_image(title, author):
    # 封面图片尺寸
    width = 800
    height = 1200

    # 背景颜色
    background_color = (255, 255, 255)  # 白色

    # 字体颜色
    text_color = (0, 0, 0)  # 黑色

    # 创建空白图片
    image = Image.new('RGB', (width, height), background_color)

    draw = ImageDraw.Draw(image)

    try:
        # 尝试使用微软雅黑字体
        font_title = ImageFont.truetype('msyh.ttc', 60)
        font_author = ImageFont.truetype('msyh.ttc', 40)
    except IOError:
        # 如果系统中没有该字体，使用默认字体
        font_title = ImageFont.load_default()
        font_author = ImageFont.load_default()
        print("警告：系统中未找到指定字体，将使用默认字体。")

    # 检查 Pillow 版本，选择合适的方法
    if hasattr(draw, 'textbbox'):
        # Pillow 10.0.0 及以上版本
        title_bbox = draw.textbbox((0, 0), title, font=font_title)
        title_w = title_bbox[2] - title_bbox[0]
        title_h = title_bbox[3] - title_bbox[1]

        author_bbox = draw.textbbox((0, 0), author, font=font_author)
        author_w = author_bbox[2] - author_bbox[0]
        author_h = author_bbox[3] - author_bbox[1]
    else:
        # Pillow 10.0.0 以下版本
        title_w, title_h = draw.textsize(title, font=font_title)
        author_w, author_h = draw.textsize(author, font=font_author)

    title_x = (width - title_w) / 2
    title_y = (height - title_h) / 2 - 50  # 向上偏移一些

    author_x = (width - author_w) / 2
    author_y = title_y + title_h + 20  # 标题下方

    # 绘制标题和作者
    draw.text((title_x, title_y), title, fill=text_color, font=font_title)
    draw.text((author_x, author_y), author, fill=text_color, font=font_author)

    # 将图片保存到内存中
    from io import BytesIO
    image_bytes = BytesIO()
    image.save(image_bytes, format='JPEG')
    cover_content = image_bytes.getvalue()

    return cover_content

if __name__ == '__main__':
    # 输入和输出目录
    txt_input_dir = r'C:\Users\User\Desktop\txt to epub\txt file'
    epub_output_dir = r'C:\Users\User\Desktop\txt to epub\epub file'

    # 处理路径中的中文字符
    txt_input_dir = os.path.abspath(txt_input_dir)
    epub_output_dir = os.path.abspath(epub_output_dir)

    # 检查输入目录是否存在
    if not os.path.exists(txt_input_dir):
        print("输入目录不存在：{}".format(txt_input_dir))
        sys.exit(1)

    # 获取输入目录中的所有 txt 文件
    txt_files = [f for f in os.listdir(txt_input_dir) if f.endswith('.txt')]

    if not txt_files:
        print("在输入目录中未找到任何 txt 文件。")
        sys.exit(0)

    for txt_file in txt_files:
        txt_file_path = os.path.join(txt_input_dir, txt_file)
        txt_to_epub(txt_file_path, epub_output_dir)

    print("所有文件处理完毕。")
