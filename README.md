# EPUB Converter

A small utility to convert Chinese novel `.txt` files to EPUB format.

## Features

- 自动检测文本编码（UTF-8 / GBK / GB18030）
- 智能识别章节标题（例如 “第四卷 第一章” 或纯数字开头）
- 生成带目录的 EPUB，并使用 SimSun（宋体）字体、良好排版
- 支持批量处理一个文件夹下的所有 `.txt` 文件
- 可选设置作者名，EPUB 元数据中包含作者信息

## 依赖

- Python 3
- 仅使用标准库

## 安装

```bash
# 如果你尚未安装 Python 环境，请先安装 Python 3.x
git clone https://github.com/yourrepo/epub-converter.git
cd epub-converter
