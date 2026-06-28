from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document

from docxrender.contracts import (
    DocxFieldRefreshOptions,
    DocxHeaderFooterImageOptions,
    DocxTemplateRenderOptions,
    DocxToPdfOptions,
    DocxToPdfResult,
    DocxWriteOptions,
    DocxWriteResult,
)
from docxrender.docx.assets import apply_header_footer_images
from docxrender.docx.body import insert_markdown_blocks
from docxrender.docx.fields import (
    write_docx_field_update_markers,
    write_frozen_docx_fields,
)
from docxrender.docx.refresh import refresh_docx_fields
from docxrender.markdown import parse_markdown_blocks
from docxrender.template import write_docx_template


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

    write_docx_template(
        DocxTemplateRenderOptions(
            file_template=options.file_template,
            file_out_docx=options.file_out_docx,
            context=options.context,
            context_defaults={"body_anchor": options.body_anchor.anchor_token},
            inline_images=options.template_inline_images,
            context_policy=options.template_context_policy,
        )
    )
    markdown_blocks = parse_markdown_blocks(
        options.markdown_body,
        options=options.markdown,
    )
    document = Document(str(options.file_out_docx))
    insert_markdown_blocks(
        document,
        markdown_blocks,
        body_anchor=options.body_anchor,
        dir_base=options.dir_base,
        style=options.style,
        body_render_policy=options.body_render_policy,
    )
    document.save(str(options.file_out_docx))
    postprocess_docx(
        file_docx=options.file_out_docx,
        should_update_fields=options.should_update_fields,
        should_freeze_fields=options.should_freeze_fields,
        field_refresh=options.field_refresh,
        header_footer_images=options.header_footer_images,
    )
    return DocxWriteResult(file_docx=options.file_out_docx)


def write_existing_docx(
    *,
    file_in_docx: Path,
    file_out_docx: Path | None = None,
    should_update_fields: bool = True,
    should_freeze_fields: bool = False,
    field_refresh: DocxFieldRefreshOptions | None = None,
    header_footer_images: DocxHeaderFooterImageOptions | None = None,
) -> DocxWriteResult:
    """Apply DOCX post-processing to an existing DOCX file.

    Args:
        file_in_docx (Path): Existing DOCX path.
        file_out_docx (Path | None): Optional output path. When omitted, the input
            DOCX is edited in place.
        should_update_fields (bool): Whether DOCX fields should be marked for update.
        should_freeze_fields (bool): Whether DOCX fields should be frozen.
        field_refresh (DocxFieldRefreshOptions | None): Optional UNO refresh options.
        header_footer_images (DocxHeaderFooterImageOptions | None): Optional
            header/footer image options.

    Returns:
        DocxWriteResult: Result containing the edited DOCX path.
    """

    path_in = file_in_docx
    path_out = path_in if file_out_docx is None else file_out_docx
    if path_out != path_in:
        path_out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path_in, path_out)
    postprocess_docx(
        file_docx=path_out,
        should_update_fields=should_update_fields,
        should_freeze_fields=should_freeze_fields,
        field_refresh=field_refresh,
        header_footer_images=header_footer_images,
    )
    return DocxWriteResult(file_docx=path_out)


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


def postprocess_docx(
    *,
    file_docx: Path,
    should_update_fields: bool,
    should_freeze_fields: bool,
    field_refresh: DocxFieldRefreshOptions | None,
    header_footer_images: DocxHeaderFooterImageOptions | None,
) -> None:
    apply_header_footer_images(file_docx, options=header_footer_images)
    if should_update_fields:
        write_docx_field_update_markers(file_docx)
    if field_refresh is None and should_freeze_fields:
        write_frozen_docx_fields(file_docx)
    refresh_docx_fields(file_docx, options=field_refresh)
