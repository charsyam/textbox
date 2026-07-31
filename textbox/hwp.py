import io
import struct
import unicodedata
import zlib
from collections import deque

import olefile

try:
    from .source import as_source
except ImportError:
    from source import as_source

try:
    from Crypto.Cipher import AES
except ImportError:  # pragma: no cover - exercised only with an incomplete install
    AES = None


class HWPExtractor(object):
    FILE_HEADER_SECTION = "FileHeader"
    HWP_SUMMARY_SECTION = "\x05HwpSummaryInformation"
    SECTION_NAME_LENGTH = len("Section")
    BODYTEXT_SECTION = "BodyText"
    VIEWTEXT_SECTION = "ViewText"
    HWP_TEXT_TAGS = [67]
    DISTRIBUTE_DOC_DATA_TAG = 28

    def __init__(self, filename, name=None):
        self.source = as_source(filename, name)
        self.filename = self.source.name
        self._ole = self.load(self.source)
        self._dirs = self._ole.listdir()

        if not self.is_valid(self._dirs):
            raise ValueError("Not Valid HwpFile")

        header_data = self._ole.openstream(self.FILE_HEADER_SECTION).read()
        flags = struct.unpack_from("<I", header_data, 36)[0]
        self._compressed = bool(flags & 0x01)
        self._encrypted = bool(flags & 0x02)
        self._distribution = bool(flags & 0x04)

        # Password-encrypted HWP and distribution/read-only HWP use different
        # schemes.  rhwp only decrypts the latter, whose key is embedded in
        # each ViewText section.
        if self._encrypted:
            raise ValueError("Password-encrypted HWP files are not supported")

        self._para_shapes = []
        self._numberings = []
        self._bullets = []
        self._numbering_counters = {}
        self._auto_number_counters = {}
        self.tables = []
        self._load_doc_info()
        self.text = self._get_text()
        self.provenance = self.source.provenance("hwp")

    def load(self, filename):
        return olefile.OleFileIO(io.BytesIO(filename.read_bytes()))

    def is_valid(self, dirs):
        return (
            [self.FILE_HEADER_SECTION] in dirs
            and [self.HWP_SUMMARY_SECTION] in dirs
        )

    def is_compressed(self, ole=None):
        if hasattr(self, "_compressed"):
            return self._compressed
        stream = (ole or self._ole).openstream(self.FILE_HEADER_SECTION)
        header_data = stream.read()
        return bool(struct.unpack_from("<I", header_data, 36)[0] & 0x01)

    def get_body_sections(self, dirs):
        storage = self.VIEWTEXT_SECTION if self._distribution else self.BODYTEXT_SECTION
        sections = []
        for path in dirs:
            if (
                len(path) == 2
                and path[0] == storage
                and path[1].startswith("Section")
            ):
                sections.append(int(path[1][self.SECTION_NAME_LENGTH :]))

        return ["{}/Section{}".format(storage, number) for number in sorted(sections)]

    def get_text(self):
        return self.text

    def get_structure(self):
        return {
            "type": "hwp",
            "provenance": self.provenance,
            "tables": self.tables,
            "text": self.text,
        }

    def _get_text(self):
        text = ""
        for section in self.get_body_sections(self._dirs):
            text += self.get_text_from_section(section)
            text += "\n"
        return text

    def get_text_from_section(self, section):
        data = self._ole.openstream(section).read()

        if self._distribution:
            unpacked_data = self._decrypt_viewtext_section(data)
        else:
            unpacked_data = self._decompress(data) if self._compressed else data

        records = self._read_records(unpacked_data)
        table_blocks, table_child_paragraphs = self._extract_tables(records)
        text = []
        for index, record in enumerate(records):
            rec_type, level, payload = record
            if rec_type != 66 or index in table_child_paragraphs:
                continue

            end = index + 1
            while end < len(records):
                next_type, next_level, _ = records[end]
                if next_type == 66 and next_level <= level:
                    break
                end += 1

            para_text = next(
                (
                    child_payload
                    for child_type, child_level, child_payload in records[index + 1 : end]
                    if child_type == 67 and child_level == level + 1
                ),
                None,
            )
            if para_text is None:
                continue

            para_shape_id = (
                struct.unpack_from("<H", payload, 8)[0] if len(payload) >= 10 else 0
            )
            prefix = self._paragraph_head(para_shape_id)
            auto_numbers = self._paragraph_auto_numbers(
                records[index + 1 : end], level + 1
            )
            paragraph_tables = list(table_blocks.get(index, ()))
            table_replacements = deque(
                "\n" + self._format_table(table) + "\n" for table in paragraph_tables
            )
            paragraph_text = self._parse_para_text(
                para_text,
                auto_number_replacements=iter(auto_numbers),
                control_replacements={(0x000B, b" lbt"): table_replacements},
            )
            # Damaged/legacy files can contain a table CTRL_HEADER without its
            # matching PARA_TEXT marker. Preserve such tables after the text.
            while table_replacements:
                paragraph_text += table_replacements.popleft()
            text.append(prefix + paragraph_text + "\n")
            for table in paragraph_tables:
                self.tables.append(table)

        return "".join(text)

    def get_tables(self):
        """Return tables with rows, columns, cell text, and merge spans."""
        return self.tables

    def _extract_tables(self, records):
        attached = {}
        child_paragraphs = set()
        for ctrl_index, (tag, ctrl_level, payload) in enumerate(records):
            if tag != 71 or len(payload) < 4 or payload[:4] != b" lbt":
                continue

            end = ctrl_index + 1
            while end < len(records) and records[end][1] > ctrl_level:
                end += 1
            table_records = records[ctrl_index + 1 : end]
            table_record_index = next(
                (i for i, record in enumerate(table_records) if record[0] == 77),
                None,
            )
            if table_record_index is None:
                continue

            table_payload = table_records[table_record_index][2]
            if len(table_payload) < 8:
                continue
            row_count, col_count = struct.unpack_from("<HH", table_payload, 4)
            if not row_count or not col_count or row_count * col_count > 100000:
                continue

            cells = []
            i = table_record_index + 1
            while i < len(table_records):
                cell_tag, cell_level, cell_payload = table_records[i]
                if cell_tag != 72 or len(cell_payload) < 34:
                    i += 1
                    continue

                cell_end = i + 1
                while cell_end < len(table_records):
                    next_tag, next_level, _ = table_records[cell_end]
                    if next_level < cell_level or (
                        next_level == cell_level and next_tag in (72, 77)
                    ):
                        break
                    cell_end += 1

                col, row, col_span, row_span = struct.unpack_from("<HHHH", cell_payload, 8)
                width, height = struct.unpack_from("<II", cell_payload, 16)
                paragraphs = []
                for relative_index in range(i + 1, cell_end):
                    child_tag, child_level, _ = table_records[relative_index]
                    if child_tag != 66:
                        continue
                    absolute_index = ctrl_index + 1 + relative_index
                    child_paragraphs.add(absolute_index)
                    next_para = relative_index + 1
                    while next_para < cell_end:
                        next_type, next_level, next_payload = table_records[next_para]
                        if next_type == 66 and next_level <= child_level:
                            break
                        if next_type == 67 and next_level == child_level + 1:
                            paragraphs.append(self._parse_para_text(next_payload))
                            break
                        next_para += 1

                cells.append(
                    {
                        "row": row,
                        "col": col,
                        "rowSpan": max(row_span, 1),
                        "colSpan": max(col_span, 1),
                        "width": width,
                        "height": height,
                        "text": "\n".join(paragraphs).strip(),
                    }
                )
                i = cell_end

            table = {
                "rows": row_count,
                "cols": col_count,
                "cells": cells,
            }
            parent = next(
                (
                    index
                    for index in range(ctrl_index - 1, -1, -1)
                    if records[index][0] == 66 and records[index][1] < ctrl_level
                ),
                None,
            )
            if parent is not None:
                attached.setdefault(parent, []).append(table)
        return attached, child_paragraphs

    @staticmethod
    def _format_table(table):
        rows, cols = table["rows"], table["cols"]
        grid = [["" for _ in range(cols)] for _ in range(rows)]
        for cell in table["cells"]:
            row, col = cell["row"], cell["col"]
            if row >= rows or col >= cols:
                continue
            value = cell["text"].replace("\n", "<br>")
            if cell["rowSpan"] > 1 or cell["colSpan"] > 1:
                value += " (병합 {}×{})".format(cell["rowSpan"], cell["colSpan"])
            grid[row][col] = value

        natural_widths = []
        for col in range(cols):
            natural_widths.append(max(
                [HWPExtractor._display_width(grid[row][col]) for row in range(rows)]
                + [3]
            ))

        # Derive each column's physical width from HWPUNIT cell measurements.
        # Unmerged cells are the strongest evidence. A merged cell contributes
        # an equal share only where no direct measurement exists.
        physical_widths = [0.0] * cols
        for cell in table["cells"]:
            col = cell["col"]
            span = max(cell["colSpan"], 1)
            width = cell.get("width", 0)
            if width <= 0 or col >= cols:
                continue
            if span == 1:
                physical_widths[col] = max(physical_widths[col], float(width))
        for cell in table["cells"]:
            col = cell["col"]
            span = max(cell["colSpan"], 1)
            width = cell.get("width", 0)
            if width <= 0 or col >= cols:
                continue
            share = float(width) / span
            for target in range(col, min(col + span, cols)):
                if physical_widths[target] == 0:
                    physical_widths[target] = share

        known = [width for width in physical_widths if width > 0]
        fallback = sum(known) / len(known) if known else 1.0
        physical_widths = [width or fallback for width in physical_widths]

        # Roughly 500 HWPUNIT per terminal column gives an A4-width table about
        # 80 columns wide. Clamp pathological documents and retain at least
        # three columns per cell.
        physical_total = sum(physical_widths)
        target_total = round(physical_total / 500)
        target_total = max(cols * 3, min(target_total, 100))
        widths = HWPExtractor._allocate_column_widths(
            physical_widths, natural_widths, target_total
        )

        horizontal = {
            "top": ("┌", "┬", "┐"),
            "middle": ("├", "┼", "┤"),
            "bottom": ("└", "┴", "┘"),
        }

        def border(kind):
            left, joint, right = horizontal[kind]
            return left + joint.join("─" * (width + 2) for width in widths) + right

        lines = ["[표 {}×{}]".format(rows, cols), border("top")]
        for row_index, row in enumerate(grid):
            wrapped = [
                HWPExtractor._wrap_display(value, widths[col])
                for col, value in enumerate(row)
            ]
            height = max(len(cell_lines) for cell_lines in wrapped)
            for line_index in range(height):
                cells = []
                for col, cell_lines in enumerate(wrapped):
                    value = cell_lines[line_index] if line_index < len(cell_lines) else ""
                    cells.append(
                        " "
                        + value
                        + " " * (widths[col] - HWPExtractor._display_width(value))
                        + " "
                    )
                lines.append("│" + "│".join(cells) + "│")
            if row_index + 1 < rows:
                lines.append(border("middle"))
        lines.append(border("bottom"))
        return "\n".join(lines)

    @staticmethod
    def _allocate_column_widths(physical, natural, target_total):
        widths = [3] * len(physical)
        remaining = max(target_total - sum(widths), 0)
        total = sum(physical) or len(physical)
        exact = [remaining * value / total for value in physical]
        additions = [int(value) for value in exact]
        for index, addition in enumerate(additions):
            widths[index] += addition
        leftover = remaining - sum(additions)
        order = sorted(
            range(len(exact)), key=lambda index: exact[index] - additions[index], reverse=True
        )
        for index in order[:leftover]:
            widths[index] += 1

        # Content may justify a wider table, but never let one column dominate.
        for index, content_width in enumerate(natural):
            widths[index] = min(max(widths[index], min(content_width, 12)), 30)
        return widths

    @staticmethod
    def _display_width(text):
        return sum(
            0
            if unicodedata.combining(char)
            else 2
            if unicodedata.east_asian_width(char) in ("W", "F")
            else 1
            for char in text
        )

    @staticmethod
    def _wrap_display(text, width):
        logical_lines = text.replace("<br>", "\n").split("\n")
        result = []
        for logical_line in logical_lines:
            current = []
            current_width = 0
            for char in logical_line:
                char_width = HWPExtractor._display_width(char)
                if current and current_width + char_width > width:
                    result.append("".join(current))
                    current = []
                    current_width = 0
                current.append(char)
                current_width += char_width
            result.append("".join(current))
        return result or [""]

    def _load_doc_info(self):
        if not self._ole.exists("DocInfo"):
            return
        data = self._ole.openstream("DocInfo").read()
        if self._compressed:
            data = self._decompress(data)

        for tag, _level, payload in self._read_records(data):
            if tag == 23:  # HWPTAG_NUMBERING
                self._numberings.append(self._parse_numbering(payload))
            elif tag == 24:  # HWPTAG_BULLET
                self._bullets.append(self._parse_bullet(payload))
            elif tag == 25:  # HWPTAG_PARA_SHAPE
                self._para_shapes.append(self._parse_para_shape(payload))

    @classmethod
    def _read_records(cls, data):
        records = []
        offset = 0
        while offset + 4 <= len(data):
            header = struct.unpack_from("<I", data, offset)[0]
            level = (header >> 10) & 0x3FF
            rec_type, rec_len, header_size = cls._record_header(data, offset)
            start = offset + header_size
            end = start + rec_len
            if end > len(data):
                raise ValueError("Truncated HWP record")
            records.append((rec_type, level, data[start:end]))
            offset = end
        return records

    @staticmethod
    def _parse_para_shape(data):
        if len(data) < 32:
            return {"head_type": 0, "level": 0, "numbering_id": 0}
        attr = struct.unpack_from("<I", data)[0]
        level = (attr >> 25) & 0x07
        if len(data) >= 58:
            level = max(level, min(struct.unpack_from("<I", data, 54)[0], 9))
        return {
            "head_type": (attr >> 23) & 0x03,
            "level": level,
            "numbering_id": struct.unpack_from("<H", data, 30)[0],
        }

    @staticmethod
    def _parse_bullet(data):
        if len(data) < 14:
            return {"char": "•", "text_distance": 0}
        return {
            "char": chr(struct.unpack_from("<H", data, 12)[0]),
            "text_distance": struct.unpack_from("<h", data, 6)[0],
        }

    @staticmethod
    def _parse_numbering(data):
        pos = 0
        heads = []
        formats = []
        for _ in range(7):
            if pos + 14 > len(data):
                return {"heads": heads, "formats": formats, "starts": [1] * 7}
            attr = struct.unpack_from("<I", data, pos)[0]
            text_distance = struct.unpack_from("<h", data, pos + 6)[0]
            heads.append(
                {
                    "format": (attr >> 5) & 0x0F,
                    "text_distance": text_distance,
                }
            )
            length = struct.unpack_from("<H", data, pos + 12)[0]
            pos += 14
            byte_len = length * 2
            formats.append(
                data[pos : pos + byte_len].decode("utf-16-le", errors="replace")
            )
            pos += byte_len

        if pos + 2 <= len(data):
            pos += 2  # legacy start number
        starts = []
        for _ in range(7):
            starts.append(
                struct.unpack_from("<I", data, pos)[0] if pos + 4 <= len(data) else 1
            )
            pos += 4
        return {"heads": heads, "formats": formats, "starts": starts}

    def _paragraph_head(self, para_shape_id):
        if para_shape_id >= len(self._para_shapes):
            return ""
        shape = self._para_shapes[para_shape_id]
        head_type = shape["head_type"]
        ref_id = shape["numbering_id"]
        if not ref_id:
            return ""

        if head_type == 3 and ref_id <= len(self._bullets):
            bullet = self._bullets[ref_id - 1]
            char = self._map_bullet_char(bullet["char"])
            return char + (" " if bullet["text_distance"] > 0 else "")

        if head_type not in (1, 2) or ref_id > len(self._numberings):
            return ""
        numbering = self._numberings[ref_id - 1]
        level = min(shape["level"], 6)
        counters = self._numbering_counters.setdefault(ref_id, [0] * 7)
        counters[level] += 1
        for deeper in range(level + 1, 7):
            counters[deeper] = 0
        result = self._expand_numbering_format(
            numbering["formats"][level], counters, numbering, level
        )
        distance = numbering["heads"][level]["text_distance"]
        return result + (" " if result and distance > 0 else "")

    def _expand_numbering_format(self, pattern, counters, numbering, level):
        def level_text(index):
            count = counters[index] or 1
            start = numbering["starts"][index] or 1
            value = start - 1 + count
            return self._format_number(value, numbering["heads"][index]["format"])

        result = []
        i = 0
        while i < len(pattern):
            if pattern[i] == "^" and i + 1 < len(pattern):
                code = pattern[i + 1]
                if code in "1234567":
                    result.append(level_text(int(code) - 1))
                    i += 2
                    continue
                if code in "nN":
                    result.append(".".join(level_text(n) for n in range(level + 1)))
                    if code == "N":
                        result.append(".")
                    i += 2
                    continue
            result.append(pattern[i])
            i += 1
        return "".join(result)

    @staticmethod
    def _format_number(number, format_code):
        if format_code == 1 and 1 <= number <= 20:
            return chr(0x2460 + number - 1)
        if format_code in (2, 3):
            values = (
                (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
            )
            result = []
            remaining = number
            for value, symbol in values:
                while remaining >= value:
                    result.append(symbol)
                    remaining -= value
            value = "".join(result)
            return value.lower() if format_code == 3 else value
        if format_code in (4, 5) and number > 0:
            result = []
            while number:
                number -= 1
                result.append(chr(ord("A") + number % 26))
                number //= 26
            value = "".join(reversed(result))
            return value.lower() if format_code == 5 else value
        if format_code == 8 and 1 <= number <= 14:
            return "가나다라마바사아자차카타파하"[number - 1]
        return str(number)

    @staticmethod
    def _map_bullet_char(char):
        mappings = {
            "\uf06c": "●",
            "\uf06d": "●",
            "\uf06e": "■",
            "\uf06f": "□",
            "\U000F02EF": "·",
        }
        return mappings.get(char, char)

    def _paragraph_auto_numbers(self, records, direct_level):
        replacements = []
        for tag, level, payload in records:
            if tag != 71 or level != direct_level or len(payload) < 8:
                continue
            control_id = payload[:4]
            if control_id == b"onwn" and len(payload) >= 10:
                kind = struct.unpack_from("<I", payload, 4)[0] & 0x0F
                self._auto_number_counters[kind] = struct.unpack_from(
                    "<H", payload, 8
                )[0] - 1
            elif control_id == b"onta":
                attr = struct.unpack_from("<I", payload, 4)[0]
                kind = attr & 0x0F
                self._auto_number_counters[kind] = (
                    self._auto_number_counters.get(kind, 0) + 1
                )
                number = self._auto_number_counters[kind]
                prefix = chr(struct.unpack_from("<H", payload, 12)[0]) if len(payload) >= 14 else ""
                suffix = chr(struct.unpack_from("<H", payload, 14)[0]) if len(payload) >= 16 else ""
                replacements.append(
                    prefix.replace("\x00", "")
                    + self._format_number(number, (attr >> 4) & 0xFF)
                    + suffix.replace("\x00", "")
                )
            elif control_id in (b"  nf", b"  ne"):
                kind = 1 if control_id == b"  nf" else 2
                self._auto_number_counters[kind] = (
                    self._auto_number_counters.get(kind, 0) + 1
                )
                replacements.append(str(self._auto_number_counters[kind]))
        return replacements


    @staticmethod
    def _parse_para_text(
        data, auto_number_replacements=None, control_replacements=None
    ):
        """Decode an HWP PARA_TEXT payload without exposing control payloads.

        HWP control characters are not ordinary UTF-16 characters. Most
        inline/extended controls occupy eight UTF-16 code units: one marker
        followed by seven units of binary metadata. This follows rhwp's
        parser/body_text.rs rules.
        """
        text = []
        pos = 0
        auto_number_replacements = auto_number_replacements or iter(())
        control_replacements = control_replacements or {}

        while pos + 1 < len(data):
            code_unit = struct.unpack_from("<H", data, pos)[0]

            if code_unit == 0:
                pos += 2
            elif code_unit == 0x0009:
                # TAB is an eight-code-unit inline control.
                text.append("\t")
                pos += 16
            elif code_unit == 0x000A:
                text.append("\n")
                pos += 2
            elif code_unit == 0x000D:
                # Paragraph terminator is not visible text.
                break
            elif HWPExtractor._is_extended_control(code_unit):
                # Auto/new-number controls reserve a visible position. The
                # actual number is resolved from surrounding control records;
                # retain a space so adjacent words do not run together.
                if code_unit in (0x0011, 0x0012):
                    text.append(next(auto_number_replacements, " "))
                else:
                    control_id = data[pos + 2 : pos + 6]
                    replacements = control_replacements.get(
                        (code_unit, control_id),
                        control_replacements.get(code_unit),
                    )
                    if replacements:
                        text.append(replacements.popleft())
                pos += 16
            elif code_unit < 0x0020:
                replacements = {
                    0x0018: "-",       # hyphen
                    0x0019: " ",       # reserved spacing character
                    0x001E: "\u00A0",  # no-break space
                    0x001F: "\u2007",  # fixed-width/figure space
                }
                replacement = replacements.get(code_unit)
                if replacement is not None:
                    text.append(replacement)
                pos += 2
            else:
                # Decode one Unicode scalar, preserving valid surrogate pairs.
                unit_count = 1
                if 0xD800 <= code_unit <= 0xDBFF and pos + 3 < len(data):
                    low = struct.unpack_from("<H", data, pos + 2)[0]
                    if 0xDC00 <= low <= 0xDFFF:
                        unit_count = 2
                raw = data[pos : pos + unit_count * 2]
                text.append(raw.decode("utf-16-le", errors="replace"))
                pos += unit_count * 2

        return "".join(text)

    @staticmethod
    def _is_extended_control(code_unit):
        # HWP 5.0 table 6: inline/extended controls occupy 16 bytes.
        # TAB(9), line break(10), and paragraph break(13) are handled above.
        return (
            0x0001 <= code_unit <= 0x0008
            or 0x000B <= code_unit <= 0x000C
            or 0x000E <= code_unit <= 0x0017
        )

    @staticmethod
    def _record_header(data, offset=0):
        if offset + 4 > len(data):
            raise ValueError("HWP record header is shorter than 4 bytes")
        header = struct.unpack_from("<I", data, offset)[0]
        rec_type = header & 0x3FF
        rec_len = (header >> 20) & 0xFFF
        header_size = 4
        if rec_len == 0xFFF:
            if offset + 8 > len(data):
                raise ValueError("Extended HWP record header is truncated")
            rec_len = struct.unpack_from("<I", data, offset + 4)[0]
            header_size = 8
        return rec_type, rec_len, header_size

    @staticmethod
    def _decompress(data):
        # HWP streams use raw DEFLATE. AES block padding/trailing bytes are
        # tolerated by zlib after the end of the compressed stream.
        return zlib.decompress(data, -15)

    @staticmethod
    def _decrypt_distribute_doc_data(data):
        if len(data) != 256:
            raise ValueError(
                "DISTRIBUTE_DOC_DATA must be 256 bytes (got {})".format(len(data))
            )

        result = bytearray(data)
        seed = struct.unpack_from("<I", result)[0]

        def rand():
            nonlocal seed
            seed = (seed * 214013 + 2531011) & 0xFFFFFFFF
            return (seed >> 16) & 0x7FFF

        remaining = 0
        key = 0
        for i in range(256):
            if remaining == 0:
                key = rand() & 0xFF
                remaining = (rand() & 0x0F) + 1
            if i >= 4:
                result[i] ^= key
            remaining -= 1
        return bytes(result)

    def _decrypt_viewtext_section(self, section_data):
        rec_type, rec_len, header_size = self._record_header(section_data)
        if rec_type != self.DISTRIBUTE_DOC_DATA_TAG:
            raise ValueError("ViewText does not start with DISTRIBUTE_DOC_DATA")
        if rec_len != 256:
            raise ValueError(
                "DISTRIBUTE_DOC_DATA must be 256 bytes (got {})".format(rec_len)
            )

        payload_end = header_size + rec_len
        decrypted_header = self._decrypt_distribute_doc_data(
            section_data[header_size:payload_end]
        )
        key_offset = 4 + (decrypted_header[0] & 0x0F)
        aes_key = decrypted_header[key_offset : key_offset + 16]

        encrypted_body = section_data[payload_end:]
        if not encrypted_body:
            raise ValueError("ViewText contains no encrypted body")
        if len(encrypted_body) % 16:
            raise ValueError("ViewText encrypted body is not AES block-aligned")
        if AES is None:
            raise ImportError(
                "pycryptodome is required to read distribution HWP files"
            )

        decrypted_body = AES.new(aes_key, AES.MODE_ECB).decrypt(encrypted_body)
        return self._decompress(decrypted_body) if self._compressed else decrypted_body


def get_text(filename):
    hwp = HWPExtractor(filename)
    print(hwp.get_text())


if __name__ == "__main__":
    import sys

    get_text(sys.argv[1])
