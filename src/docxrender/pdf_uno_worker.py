"""Subprocess entrypoint for LibreOffice UNO PDF conversion."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import cast

from docxrender.pdf_uno import (
    convert_docx_to_pdf_with_uno,
    create_docx_to_pdf_result,
    create_docx_to_pdf_state,
    deserialize_docx_to_pdf_options,
    validate_docx_input,
)


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docxrender.pdf_uno_worker",
        description="Run DOCX-to-PDF conversion in a Python process with UNO access.",
    )
    parser.add_argument(
        "--options-json",
        type=Path,
        required=True,
        help="Path to serialized DocxToPdfOptions JSON.",
    )
    return parser


def run_options_file(file_options: Path) -> None:
    payload_raw = json.loads(file_options.read_text(encoding="utf-8"))
    if not isinstance(payload_raw, dict):
        raise ValueError("DOCX-to-PDF options JSON must be an object.")
    payload = cast(dict[str, object], payload_raw)
    options = deserialize_docx_to_pdf_options(payload)
    state = create_docx_to_pdf_state(options)
    validate_docx_input(state)
    convert_docx_to_pdf_with_uno(state)
    create_docx_to_pdf_result(state)


def main() -> int:
    args = create_parser().parse_args()
    try:
        run_options_file(args.options_json)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
