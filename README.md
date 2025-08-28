# InvoiceTranslator

A simple Python application that replaces words in Word documents based on
mappings defined in `words.txt` while preserving the original layout and
formatting.

## Usage
1. Place the `.docx` files to translate in the `invoice/` directory.
2. Add word mappings to `words.txt` using the format `original=replacement`.
3. Install dependencies and run the translator:
   ```bash
   pip install python-docx
   python translate.py
   ```
4. Translated files will be saved in the `translated/` directory with the
   suffix `_translated` added to the original filename.

