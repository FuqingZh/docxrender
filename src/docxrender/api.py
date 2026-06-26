from __future__ import annotations

from typing import Any, cast

from docx import Document
from docxtpl import DocxTemplate

from docxrender.contracts import (
    DocxToPdfOptions,
    DocxToPdfResult,
    DocxWriteOptions,
    DocxWriteResult,
)
from docxrender.docx.body import insert_markdown_blocks
from docxrender.docx.fields import (
    write_docx_field_update_markers,
    write_frozen_docx_fields,
)
from docxrender.docx.refresh import refresh_docx_fields
from docxrender.markdown import parse_markdown_blocks


def write_docx(options: DocxWriteOptions) -> DocxWriteResult:
    """Write a DOCX file from a template, context, markdown body, and style.

    Args:
        options (DocxWriteOptions): DOCX writing options.

    Returns:
        DocxWriteResult: Result containing the written DOCX path.

    Raises:
        FileNotFoundError: The template or a referenced image does not exist.
        RuntimeError: The rendered DOCX cannot be opened or written.
    """

    _write_template_docx(options)
    markdown_blocks = parse_markdown_blocks(options.markdown_body)
    document = Document(str(options.file_out_docx))
    insert_markdown_blocks(
        document,
        markdown_blocks,
        anchor_token=options.anchor_token,
        dir_base=options.dir_base,
        style=options.style,
    )
    document.save(str(options.file_out_docx))
    if options.should_update_fields:
        write_docx_field_update_markers(options.file_out_docx)
    if options.field_refresh is None and options.should_freeze_fields:
        write_frozen_docx_fields(options.file_out_docx)
    refresh_docx_fields(options.file_out_docx, options=options.field_refresh)
    return DocxWriteResult(file_docx=options.file_out_docx)


def convert_docx_to_pdf(options: DocxToPdfOptions) -> DocxToPdfResult:
    """Convert a DOCX file to PDF through LibreOffice.

    Args:
        options (DocxToPdfOptions): DOCX-to-PDF conversion options.

    Returns:
        DocxToPdfResult: Result containing the written PDF path and optional refreshed
            DOCX path.

    Raises:
        FileNotFoundError: The input DOCX does not exist.
        RuntimeError: LibreOffice or UNO cannot load or convert the document.
    """

    from docxrender.pdf_uno import run_docx_to_pdf_pipeline

    return run_docx_to_pdf_pipeline(options)


def _write_template_docx(options: DocxWriteOptions) -> None:
    options.file_out_docx.parent.mkdir(parents=True, exist_ok=True)
    template = cast(Any, DocxTemplate(str(options.file_template)))
    context = dict(options.context)
    context.setdefault("body_anchor", options.anchor_token)
    template.render(context)
    template.save(str(options.file_out_docx))
