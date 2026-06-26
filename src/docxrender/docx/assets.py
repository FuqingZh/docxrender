from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from docxrender.contracts import DocxHeaderFooterImageOptions

IMAGE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


def apply_header_footer_images(
    file_docx: Path,
    *,
    options: DocxHeaderFooterImageOptions | None,
) -> None:
    """Apply configured header and footer image replacements.

    Args:
        file_docx (Path): DOCX file to edit in place.
        options (DocxHeaderFooterImageOptions | None): Header/footer image options.
    """

    if options is None:
        return
    _validate_image_file(options.file_header_image)
    _validate_image_file(options.file_footer_image)

    needs_insert = False
    if options.file_header_image is not None:
        paths_header = _read_media_paths_from_rels(file_docx, part_prefix="header")
        if paths_header and options.should_replace_existing:
            _replace_docx_media(
                file_docx=file_docx,
                paths_media=paths_header,
                file_replacement=options.file_header_image,
            )
        elif not paths_header and options.should_insert_when_missing:
            needs_insert = True
    if options.file_footer_image is not None:
        paths_footer = _read_media_paths_from_rels(file_docx, part_prefix="footer")
        if paths_footer and options.should_replace_existing:
            _replace_docx_media(
                file_docx=file_docx,
                paths_media=paths_footer,
                file_replacement=options.file_footer_image,
            )
        elif not paths_footer and options.should_insert_when_missing:
            needs_insert = True

    if needs_insert:
        _insert_missing_header_footer_images(file_docx, options=options)


def _validate_image_file(file_image: Path | None) -> None:
    if file_image is None:
        return
    if not file_image.is_file():
        raise FileNotFoundError(f"Header/footer image does not exist: {file_image}")


def _read_media_paths_from_rels(file_docx: Path, *, part_prefix: str) -> set[str]:
    paths_media: set[str] = set()
    with zipfile.ZipFile(file_docx) as zip_file:
        for name in zip_file.namelist():
            if not (
                name.startswith(f"word/_rels/{part_prefix}")
                and name.endswith(".xml.rels")
            ):
                continue
            root = ET.fromstring(zip_file.read(name))
            for rel in root:
                if rel.attrib.get("Type") != IMAGE_REL_TYPE:
                    continue
                target = rel.attrib.get("Target", "")
                if target.startswith("../"):
                    target = target[3:]
                if not target.startswith("word/"):
                    target = f"word/{target}"
                paths_media.add(target)
    return paths_media


def _replace_docx_media(
    *,
    file_docx: Path,
    paths_media: set[str],
    file_replacement: Path,
) -> None:
    data_replacement = file_replacement.read_bytes()
    with tempfile.NamedTemporaryFile(
        delete=False,
        dir=file_docx.parent,
        suffix=".docx",
    ) as handle_tmp:
        file_tmp = Path(handle_tmp.name)
    try:
        with (
            zipfile.ZipFile(file_docx) as zip_in,
            zipfile.ZipFile(file_tmp, "w", compression=zipfile.ZIP_DEFLATED) as zip_out,
        ):
            for item in zip_in.infolist():
                data = (
                    data_replacement
                    if item.filename in paths_media
                    else zip_in.read(item.filename)
                )
                zip_out.writestr(item, data)
        shutil.move(str(file_tmp), str(file_docx))
    finally:
        if file_tmp.exists():
            file_tmp.unlink()


def _insert_missing_header_footer_images(
    file_docx: Path,
    *,
    options: DocxHeaderFooterImageOptions,
) -> None:
    document = Document(str(file_docx))
    sections = list(document.sections)
    target_sections = sections[max(options.idx_section_start, 0) :]
    for section in target_sections:
        if options.file_header_image is not None and not section.header.part.rels:
            paragraph = section.header.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run().add_picture(
                str(options.file_header_image),
                width=Cm(options.width_cm),
            )
        if options.file_footer_image is not None and not section.footer.part.rels:
            paragraph = section.footer.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run().add_picture(
                str(options.file_footer_image),
                width=Cm(options.width_cm),
            )
    document.save(str(file_docx))
