# textbox

Extract readable text and lightweight layout information from HWP, DOCX, PPTX,
and XLSX files.

```bash
pip install -r requirements.txt
python textbox/hwp.py document.hwp
python textbox/docx.py document.docx
python textbox/pptx.py slides.pptx
python textbox/xlsx.py workbook.xlsx
```

Each extractor keeps the original `get_text()` API and also exposes structured
content:

```python
from textbox.xlsx import XLSXExtractor

document = XLSXExtractor("workbook.xlsx")
print(document.get_text())
print(document.get_structure())
```

- HWP: paragraphs, distribution-document decryption, numbering, and tables
- DOCX: paragraphs and tables in document order, lists, headers, and footers
- PPTX: position-based reading order, tables, charts, groups, and speaker notes
- XLSX: worksheets, cell values/formulas, merged ranges, and column widths
