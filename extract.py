#!/usr/bin/env python3
"""Command-line entry point for signature-based document extraction."""

import argparse
import json
import os
import sys

from textbox import TextboxExtractor


class ExternalFile(object):
    """Expose a normal filesystem file through the forensic VFS interface."""

    def __init__(self, path):
        self.path = os.path.abspath(os.fspath(path))
        self.name = os.path.basename(self.path)
        self.size = os.path.getsize(self.path)

    def read_random(self, offset, size):
        if offset < 0 or size < 0:
            raise ValueError("offset and size must be non-negative")
        with open(self.path, "rb") as stream:
            stream.seek(offset)
            return stream.read(size)


def build_parser():
    parser = argparse.ArgumentParser(
        description=(
            "Wrap a filesystem file as a virtual file, detect its format from "
            "the signature, and extract its contents."
        )
    )
    parser.add_argument("file", help="document file to extract")
    parser.add_argument(
        "--format",
        choices=("hwp", "docx", "pptx", "xlsx", "pdf"),
        help="force a format instead of signature detection",
    )
    output = parser.add_mutually_exclusive_group()
    output.add_argument(
        "--structure", action="store_true", help="print structured data as JSON"
    )
    output.add_argument(
        "--provenance", action="store_true", help="print evidence metadata as JSON"
    )
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    virtual_file = ExternalFile(args.file)
    document = TextboxExtractor(
        virtual_file,
        name=virtual_file.name,
        format=args.format,
    )

    if args.structure:
        result = document.get_structure()
        result.setdefault("provenance", document.provenance)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.provenance:
        print(json.dumps(document.provenance, ensure_ascii=False, indent=2))
    else:
        print(document.get_text())
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, ValueError) as error:
        print("extract.py: {}".format(error), file=sys.stderr)
        sys.exit(1)
