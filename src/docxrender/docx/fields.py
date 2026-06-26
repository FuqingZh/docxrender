"""DOCX field preparation helpers."""

from __future__ import annotations

import re
import tempfile
import zipfile
from collections.abc import Callable
from pathlib import Path

DOCX_FIELD_PART_PATTERN = re.compile(r"word/(?:document|header\d+|footer\d+)\.xml$")
DOCX_SETTING_PART_PATTERN = re.compile(r"word/settings\.xml$")
DOCX_PARAGRAPH_PATTERN = re.compile(r"<w:p\b.*?</w:p>", re.S)
DOCX_FIELD_BEGIN_RUN_PATTERN = re.compile(
    (
        r"<w:r\b[^>]*>"
        r"(?:(?!</w:r>).)*?"
        r"<w:fldChar\b[^>]*\bw:fldCharType=\"begin\"[^>]*/>"
        r"(?:(?!</w:r>).)*?"
        r"</w:r>"
    ),
    re.S,
)
DOCX_FIELD_SEPARATE_RUN_PATTERN = re.compile(
    (
        r"<w:r\b[^>]*>"
        r"(?:(?!</w:r>).)*?"
        r"<w:fldChar\b[^>]*\bw:fldCharType=\"separate\"[^>]*/>"
        r"(?:(?!</w:r>).)*?"
        r"</w:r>"
    ),
    re.S,
)
DOCX_FIELD_END_RUN_PATTERN = re.compile(
    (
        r"<w:r\b[^>]*>"
        r"(?:(?!</w:r>).)*?"
        r"<w:fldChar\b[^>]*\bw:fldCharType=\"end\"[^>]*/>"
        r"(?:(?!</w:r>).)*?"
        r"</w:r>"
    ),
    re.S,
)


def write_docx_field_update_markers(file_docx: Path) -> None:
    def edit_part(filename: str, data: bytes) -> bytes:
        if DOCX_SETTING_PART_PATTERN.fullmatch(filename):
            return _ensure_update_fields_setting(data.decode("utf-8")).encode("utf-8")
        if DOCX_FIELD_PART_PATTERN.fullmatch(filename):
            return _ensure_field_dirty_attrs(data.decode("utf-8")).encode("utf-8")
        return data

    _rewrite_docx_zip(file_docx, edit_part)


def write_frozen_docx_fields(file_docx: Path) -> None:
    def edit_part(filename: str, data: bytes) -> bytes:
        if DOCX_SETTING_PART_PATTERN.fullmatch(filename):
            return _strip_docx_update_fields_setting(data.decode("utf-8")).encode(
                "utf-8"
            )
        if DOCX_FIELD_PART_PATTERN.fullmatch(filename):
            text_part = data.decode("utf-8")
            text_part = _strip_docx_field_dirty_attrs(text_part)
            text_part = _freeze_docx_field_result_runs(text_part)
            return text_part.encode("utf-8")
        return data

    _rewrite_docx_zip(file_docx, edit_part)


def _rewrite_docx_zip(
    file_docx: Path,
    edit_part: Callable[[str, bytes], bytes],
) -> None:
    with tempfile.NamedTemporaryFile(
        suffix=".docx",
        delete=False,
        dir=file_docx.parent,
    ) as file_tmp:
        path_tmp = Path(file_tmp.name)
    try:
        with (
            zipfile.ZipFile(file_docx, "r") as zip_in,
            zipfile.ZipFile(path_tmp, "w", compression=zipfile.ZIP_DEFLATED) as zip_out,
        ):
            for item in zip_in.infolist():
                zip_out.writestr(item, edit_part(item.filename, zip_in.read(item)))
        path_tmp.replace(file_docx)
    finally:
        if path_tmp.exists():
            path_tmp.unlink()


def _ensure_update_fields_setting(text_settings_xml: str) -> str:
    if "<w:updateFields" in text_settings_xml:
        return re.sub(
            r"<w:updateFields\b[^>]*/>",
            '<w:updateFields w:val="true"/>',
            text_settings_xml,
        )
    return text_settings_xml.replace(
        "</w:settings>",
        '<w:updateFields w:val="true"/></w:settings>',
    )


def _strip_docx_update_fields_setting(text_settings_xml: str) -> str:
    return re.sub(r"<w:updateFields\b[^>]*/>", "", text_settings_xml)


def _ensure_field_dirty_attrs(text_part_xml: str) -> str:
    return re.sub(
        r"<w:fldChar\b(?![^>]*/?w:dirty=)",
        '<w:fldChar w:dirty="true"',
        text_part_xml,
    )


def _strip_docx_field_dirty_attrs(text_part_xml: str) -> str:
    return re.sub(r"\s+w:dirty=\"[^\"]*\"", "", text_part_xml)


def _freeze_docx_field_result_runs(text_part_xml: str) -> str:
    return DOCX_PARAGRAPH_PATTERN.sub(_freeze_field_paragraph_match, text_part_xml)


def _freeze_field_paragraph_match(match_paragraph: re.Match[str]) -> str:
    text_paragraph = match_paragraph.group(0)
    if "<w:fldChar" not in text_paragraph:
        return text_paragraph
    text_paragraph = DOCX_FIELD_BEGIN_RUN_PATTERN.sub("", text_paragraph)
    text_paragraph = DOCX_FIELD_SEPARATE_RUN_PATTERN.sub("", text_paragraph)
    text_paragraph = DOCX_FIELD_END_RUN_PATTERN.sub("", text_paragraph)
    text_paragraph = re.sub(
        r"<w:instrText\b[^>]*>.*?</w:instrText>",
        "",
        text_paragraph,
    )
    return text_paragraph
