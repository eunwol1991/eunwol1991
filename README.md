# EPUB Converter

A small utility to convert Chinese novel `.txt` files to EPUB format.

## Usage

```bash
python3 epub_converter.py /path/to/txt_folder /path/to/output_folder [-a "Author"]
```

All `.txt` files in the input directory will be processed. Output EPUB files will be saved in the specified output directory. Text encoding is detected automatically and common chapter heading styles are recognized.

### Single File
To convert a single text file:

```bash
python3 epub_converter.py path/to/file.txt path/to/output_dir -a "Author"
```

### Batch Conversion
Place all `.txt` files in one folder and run the script as shown in the basic usage example above. Each text file becomes an EPUB using the file name as the title.

## Dependencies
- Python 3
- Standard library only

## License
See [LICENSE](LICENSE) for license information.

