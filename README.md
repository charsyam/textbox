# textbox

Extract readable text and lightweight layout information from HWP, DOCX, PPTX,
XLSX, and PDF files.

```bash
pip install -r requirements.txt
python textbox/hwp.py document.hwp
python textbox/docx.py document.docx
python textbox/pptx.py slides.pptx
python textbox/xlsx.py workbook.xlsx
python textbox/pdf.py document.pdf
# Automatically detect the format and extract through a virtual-file wrapper:
python extract.py evidence.bin
```

The common CLI reads the signature rather than trusting the extension. Use
`--structure` for JSON layout data, `--provenance` for evidence metadata, or
`--format hwp` to override detection for a damaged file.

Each extractor keeps the original `get_text()` API and also exposes structured
content:

```python
from textbox.xlsx import XLSXExtractor

document = XLSXExtractor("workbook.xlsx")
print(document.get_text())
print(document.get_structure())
```

## Virtual and forensic files

The common `extract()` entry point detects the format from file signatures and
container members, not from the filename. It accepts paths, bytes, binary
streams, and forensic VFS objects exposing `size` and
`read_random(offset, size)`:

```python
from textbox import TextboxExtractor, extract, extract_text

document = extract(file_bytes, name="deleted-document.bin")
document = extract(io.BytesIO(file_bytes), name="memory-object")
document = extract(vfs_file)  # vfs_file.size + vfs_file.read_random(...)

# A stable facade when callers should not depend on format-specific classes:
document = TextboxExtractor(vfs_file)
text = extract_text(vfs_file)

print(document.get_text())
print(document.provenance)
# {"name": ..., "size": ..., "sha256": ..., "detectedFormat": ...}
```

Supported signatures are HWP/CFB, DOCX/PPTX/XLSX/ZIP, and PDF. Caller-owned
streams are not closed and their original position is restored. Random-access
VFS objects are read directly through a seekable adapter.

- HWP: paragraphs, distribution-document decryption, numbering, and tables
- DOCX: paragraphs and tables in document order, lists, headers, and footers
- PPTX: position-based reading order, tables, charts, groups, and speaker notes
- XLSX: worksheets, cell values/formulas, merged ranges, and column widths
- PDF: position-based text blocks, tables, pages, and optional Tesseract OCR

PDF OCR is automatic only for pages without usable text or with strongly
corrupted Unicode mappings. Tesseract is optional and must be installed
separately (including the `kor` language data for Korean):

```python
from textbox.pdf import PDFExtractor

document = PDFExtractor("scan.pdf", ocr="auto", ocr_language="kor+eng")
# Force OCR when a legacy font produces plausible-looking but incorrect text:
document = PDFExtractor("legacy-font.pdf", ocr=True)
```
