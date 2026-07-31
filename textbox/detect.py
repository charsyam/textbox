import zipfile

import olefile

from .source import as_source


class UnsupportedFormatError(ValueError):
    pass


def detect_format(source, name=None):
    document = as_source(source, name)
    signature = document.read_prefix(16)
    if signature.startswith(b"%PDF-"):
        return "pdf"
    if signature.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        with document.open() as stream:
            try:
                ole = olefile.OleFileIO(stream)
                try:
                    if ole.exists("FileHeader"):
                        header = ole.openstream("FileHeader").read(32)
                        if header.startswith(b"HWP Document File"):
                            return "hwp"
                finally:
                    ole.close()
            except (OSError, IOError, ValueError):
                pass
        return "ole"
    if signature.startswith((b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")):
        with document.open() as stream:
            try:
                with zipfile.ZipFile(stream) as archive:
                    names = set(archive.namelist())
            except (OSError, zipfile.BadZipFile):
                return "zip"
        if "word/document.xml" in names:
            return "docx"
        if "ppt/presentation.xml" in names:
            return "pptx"
        if "xl/workbook.xml" in names:
            return "xlsx"
        return "zip"
    return "unknown"
