- 👋 Hi, I’m gpec
- 👀 I’m interested in sleep
- 🌱 I’m currently learning python,c and other languages
- 📫 How to reach me:gmail:eunwol1991@gmail.com & line:kingdom1464

## EPUB Converter

Use `epub_converter.py` to turn Chinese novel `.txt` files into EPUB books.

Run without arguments to be prompted for the folder containing your `.txt` files and the destination folder for the EPUBs:

```
python3 epub_converter.py
```

You can also provide paths on the command line:

```
python3 epub_converter.py -i "C:/path/to/txt folder" -o "C:/path/to/output folder"
```

All `.txt` files in the input folder are converted in batch. Chapter headings are detected automatically (including patterns like `第四卷 第一章`) and the resulting EPUBs use the **SimSun** font so they display nicely in the iPhone Books app.
