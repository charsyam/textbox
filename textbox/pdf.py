import unicodedata

import pymupdf

try:
    from ._layout import render_grid
    from .source import as_source
except ImportError:
    from _layout import render_grid
    from source import as_source


class PDFExtractor(object):
    def __init__(
        self,
        filename,
        ocr="auto",
        ocr_language="kor+eng",
        ocr_dpi=200,
        extract_tables=True,
        name=None,
    ):
        self.source = as_source(filename, name)
        self.filename = self.source.name
        self.ocr = ocr
        self.ocr_language = ocr_language
        self.ocr_dpi = ocr_dpi
        self.extract_tables = extract_tables
        self.pages = []
        self.tables = []
        self._extract()
        self.text = self._render()
        self.provenance = self.source.provenance("pdf")

    def get_text(self):
        return self.text

    def get_structure(self):
        return {
            "type": "pdf",
            "provenance": self.provenance,
            "pages": self.pages,
            "tables": self.tables,
        }

    def get_pages(self):
        return self.pages

    def get_tables(self):
        return self.tables

    def _extract(self):
        with pymupdf.open(stream=self.source.read_bytes(), filetype="pdf") as document:
            metadata = document.metadata or {}
            self.metadata = metadata
            for page_index, page in enumerate(document):
                normal_text = page.get_text("text", sort=True)
                needs_ocr = self._needs_ocr(page, normal_text)
                text_page = None
                used_ocr = False
                ocr_error = None
                if self.ocr is True or (self.ocr == "auto" and needs_ocr):
                    try:
                        text_page = page.get_textpage_ocr(
                            language=self.ocr_language,
                            dpi=self.ocr_dpi,
                            full=True,
                        )
                        used_ocr = True
                    except Exception as error:
                        ocr_error = str(error)

                tables = self._extract_page_tables(page, page_index)
                table_rects = [pymupdf.Rect(table["bbox"]) for table in tables]
                raw_blocks = page.get_text(
                    "blocks", sort=True, textpage=text_page
                )
                blocks = []
                for block in raw_blocks:
                    if len(block) < 7 or block[6] != 0 or not block[4].strip():
                        continue
                    rect = pymupdf.Rect(block[:4])
                    center = (rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2
                    if any(table_rect.contains(center) for table_rect in table_rects):
                        continue
                    blocks.append(
                        {
                            "kind": "text",
                            "bbox": [round(value, 3) for value in block[:4]],
                            "text": block[4].strip(),
                            "order": int(block[5]),
                        }
                    )

                items = list(blocks)
                items.extend(
                    {
                        "kind": "table",
                        "bbox": table["bbox"],
                        "table": table,
                        "order": 100000 + index,
                    }
                    for index, table in enumerate(tables)
                )
                items.sort(key=lambda item: (item["bbox"][1], item["bbox"][0], item["order"]))
                page_info = {
                    "number": page_index + 1,
                    "width": round(page.rect.width, 3),
                    "height": round(page.rect.height, 3),
                    "rotation": page.rotation,
                    "usedOCR": used_ocr,
                    "needsOCR": needs_ocr,
                    "ocrError": ocr_error,
                    "items": items,
                }
                self.pages.append(page_info)

    def _needs_ocr(self, page, text):
        if self.ocr is False:
            return False
        visible = "".join(char for char in text if not char.isspace())
        if not visible:
            return bool(page.get_images(full=True)) or len(page.get_drawings()) > 20
        bad = sum(
            char == "\ufffd" or unicodedata.category(char) == "Co"
            for char in visible
        )
        return bad / len(visible) > 0.15

    def _extract_page_tables(self, page, page_index):
        if not self.extract_tables:
            return []
        try:
            found = page.find_tables()
        except Exception:
            return []
        tables = []
        for table_index, table in enumerate(found.tables):
            rows = table.extract()
            normalized = [
                ["" if value is None else str(value).strip() for value in row]
                for row in rows
            ]
            item = {
                "page": page_index + 1,
                "index": table_index,
                "bbox": [round(value, 3) for value in table.bbox],
                "rows": len(normalized),
                "cols": max((len(row) for row in normalized), default=0),
                "data": normalized,
            }
            tables.append(item)
            self.tables.append(item)
        return tables

    @staticmethod
    def _render_table(table):
        return render_grid(
            table["data"],
            title="[표 {}×{}]".format(table["rows"], table["cols"]),
        )

    def _render(self):
        output = []
        for page in self.pages:
            header = "[페이지 {}]".format(page["number"])
            if page["usedOCR"]:
                header += " [OCR]"
            elif page["needsOCR"] and page["ocrError"]:
                header += " [OCR 실패: {}]".format(page["ocrError"])
            lines = [header]
            for item in page["items"]:
                if item["kind"] == "table":
                    lines.append(self._render_table(item["table"]))
                else:
                    lines.append(item["text"])
            output.append("\n\n".join(lines))
        return "\n\n".join(output)


def get_text(filename):
    print(PDFExtractor(filename).get_text())


if __name__ == "__main__":
    import sys

    get_text(sys.argv[1])
