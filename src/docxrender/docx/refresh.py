"""Optional DOCX field refresh helpers."""

from __future__ import annotations

import re
import time
import zipfile
from html import unescape
from pathlib import Path

from docxrender.contracts import DocxFieldRefreshOptions
from docxrender.docx.fields import DOCX_FIELD_PART_PATTERN, write_frozen_docx_fields

TOC_FIELD_PATTERN = re.compile(
    (
        r"<w:fldChar\b[^>]*\bw:fldCharType=\"begin\"[^>]*/>"
        r"(?:(?!<w:fldChar\b[^>]*\bw:fldCharType=\"end\").)*?"
        r"<w:instrText\b[^>]*>[^<]*\bTOC\b[^<]*</w:instrText>"
        r"(?:(?!<w:fldChar\b[^>]*\bw:fldCharType=\"end\").)*?"
        r"<w:fldChar\b[^>]*\bw:fldCharType=\"separate\"[^>]*/>"
        r"(?P<result>.*?)"
        r"<w:fldChar\b[^>]*\bw:fldCharType=\"end\"[^>]*/>"
    ),
    re.S,
)
TEXT_RUN_PATTERN = re.compile(r"<w:t\b[^>]*>(?P<text>.*?)</w:t>", re.S)


def refresh_docx_fields(
    file_docx: Path,
    *,
    options: DocxFieldRefreshOptions | None,
) -> None:
    if options is None:
        return

    file_refreshed = options.file_out_docx_refreshed or file_docx
    from docxrender.pdf_uno import refresh_docx_with_uno

    refresh_docx_with_uno(
        file_in_docx=file_docx,
        file_out_docx=file_refreshed,
        options=options,
    )
    wait_for_refreshed_docx(file_refreshed, options=options)
    if options.should_require_toc:
        validate_docx_toc_result(file_refreshed)
    if options.should_freeze_fields:
        write_frozen_docx_fields(file_refreshed)


def wait_for_refreshed_docx(
    file_docx: Path,
    *,
    options: DocxFieldRefreshOptions,
) -> None:
    deadline = time.monotonic() + options.timeout_seconds
    stable_checks_required = max(options.stable_checks, 1)
    stable_checks_seen = 0
    stat_previous: tuple[int, int] | None = None

    while time.monotonic() <= deadline:
        if file_docx.exists() and file_docx.is_file() and file_docx.stat().st_size > 0:
            stat_current = (file_docx.stat().st_size, file_docx.stat().st_mtime_ns)
            if stat_current == stat_previous:
                stable_checks_seen += 1
            else:
                stable_checks_seen = 1
                stat_previous = stat_current
            if stable_checks_seen >= stable_checks_required:
                return
        time.sleep(max(options.poll_interval_seconds, 0.0))

    raise TimeoutError(
        "Refreshed DOCX did not become stable before timeout: "
        f"file_docx={file_docx.resolve()} "
        f"timeout_seconds={options.timeout_seconds} "
        f"stable_checks={options.stable_checks}"
    )


def validate_docx_toc_result(file_docx: Path) -> None:
    if has_materialized_toc_result(file_docx):
        return
    raise RuntimeError(
        "DOCX TOC result was not materialized after field refresh: "
        f"file_docx={file_docx.resolve()}"
    )


def has_materialized_toc_result(file_docx: Path) -> bool:
    for text_part in read_docx_field_parts(file_docx):
        for match in TOC_FIELD_PATTERN.finditer(text_part):
            if extract_text_from_field_result(match.group("result")).strip():
                return True
    return False


def read_docx_field_parts(file_docx: Path) -> tuple[str, ...]:
    parts: list[str] = []
    with zipfile.ZipFile(file_docx, "r") as zip_file:
        for name in zip_file.namelist():
            if DOCX_FIELD_PART_PATTERN.fullmatch(name):
                parts.append(zip_file.read(name).decode("utf-8"))
    return tuple(parts)


def extract_text_from_field_result(text_result_xml: str) -> str:
    texts = [
        unescape(match.group("text"))
        for match in TEXT_RUN_PATTERN.finditer(text_result_xml)
    ]
    return "".join(texts)
