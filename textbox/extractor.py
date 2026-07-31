from .detect import UnsupportedFormatError, detect_format
from .source import as_source


def extract(source, name=None, format=None, **options):
    """Detect and extract a document from a path or forensic virtual file."""
    document = as_source(source, name)
    detected = (format or detect_format(document)).lower()
    if detected == "hwp":
        from .hwp import HWPExtractor

        extractor = HWPExtractor(document, **options)
    elif detected == "docx":
        from .docx import DOCXExtractor

        extractor = DOCXExtractor(document, **options)
    elif detected == "pptx":
        from .pptx import PPTXExtractor

        extractor = PPTXExtractor(document, **options)
    elif detected == "xlsx":
        from .xlsx import XLSXExtractor

        extractor = XLSXExtractor(document, **options)
    elif detected == "pdf":
        from .pdf import PDFExtractor

        extractor = PDFExtractor(document, **options)
    else:
        raise UnsupportedFormatError(
            "unsupported document format: {} ({})".format(detected, document.name)
        )
    extractor.provenance = document.provenance(detected)
    return extractor


class TextboxExtractor(object):
    """Format-independent facade selected from the document signature."""

    def __init__(self, source, name=None, format=None, **options):
        self.document = extract(
            source, name=name, format=format, **options
        )
        self.provenance = self.document.provenance
        self.detected_format = self.provenance["detectedFormat"]

    def get_text(self):
        return self.document.get_text()

    def get_structure(self):
        return self.document.get_structure()

    def get_tables(self):
        getter = getattr(self.document, "get_tables", None)
        return getter() if getter else []

    def get_pages(self):
        getter = getattr(self.document, "get_pages", None)
        return getter() if getter else []

    def get_sheets(self):
        getter = getattr(self.document, "get_sheets", None)
        return getter() if getter else []

    def get_slides(self):
        getter = getattr(self.document, "get_slides", None)
        return getter() if getter else []


def extract_text(source, name=None, format=None, **options):
    return extract(source, name=name, format=format, **options).get_text()
