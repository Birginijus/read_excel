#!/usr/bin/env python3
"""Print an Excel workbook's worksheet contents to standard output."""

import argparse
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    openpyxl = None


def print_worksheet(workbook, sheet_name):
    worksheet = workbook[sheet_name]
    print(f"[Sheet: {sheet_name}]")

    for row in worksheet.iter_rows(values_only=True):
        print("\t".join("" if value is None else str(value) for value in row))


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Read an Excel workbook and print its contents as tab-separated rows."
    )
    parser.add_argument("excel_file", type=Path, help="Path to an .xlsx or .xlsm workbook")
    parser.add_argument(
        "--sheet",
        action="append",
        dest="sheets",
        metavar="NAME",
        help="Worksheet to print (may be repeated; default: all worksheets)",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    path = args.excel_file

    if not path.exists():
        print(f"Error: file not found: {path}", file=sys.stderr)
        return 1
    if not path.is_file():
        print(f"Error: not a file: {path}", file=sys.stderr)
        return 1
    if openpyxl is None:
        print(
            "Error: openpyxl is not installed. Install dependencies with "
            "'python3 -m pip install -r requirements.txt'",
            file=sys.stderr,
        )
        return 1

    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            sheet_names = set(args.sheets) if args.sheets else None
            if sheet_names is not None:
                unknown = sheet_names - set(workbook.sheetnames)
                if unknown:
                    available = ", ".join(workbook.sheetnames)
                    print(
                        f"Error: unknown sheet(s): {', '.join(sorted(unknown))}. "
                        f"Available sheets: {available}",
                        file=sys.stderr,
                    )
                    return 1

            for sheet_name in workbook.sheetnames:
                if sheet_names is None or sheet_name in sheet_names:
                    print_worksheet(workbook, sheet_name)
        finally:
            workbook.close()
    except (OSError, openpyxl.exceptions.InvalidFileException) as exc:
        print(f"Error: unsupported or invalid Excel file: {path}: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
