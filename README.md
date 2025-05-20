## EPUB Converter

Use `epub_converter.py` to batch convert Chinese novel `.txt` files into EPUB books.

```
python3 epub_converter.py /path/to/txt_folder /path/to/output_folder [-a "Author"]
```

All `.txt` files found in the input folder will be converted. The script automatically detects common encodings, handles different chapter heading styles, and sets the font to **SimSun** so novels display nicely in the iPhone Books app.