from __future__ import annotations

from pathlib import Path
from typing import Any, cast

from docx.document import Document as DocxDocument
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement  # pyright: ignore[reportUnknownVariableType]
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docxrender.contracts import DocxStyle
from docxrender.markdown import (
    MarkdownBlock,
    MarkdownHeading,
    MarkdownImage,
    MarkdownOrderedList,
    MarkdownPageBreak,
    MarkdownParagraph,
    MarkdownSpacer,
    MarkdownTable,
)


def insert_markdown_blocks(
    document: DocxDocument,
    markdown_blocks: tuple[MarkdownBlock, ...],
    *,
    anchor_token: str,
    dir_base: Path,
    style: DocxStyle,
) -> None:
    anchor = _find_anchor_paragraph(document, anchor_token)
    if anchor is None:
        for block in markdown_blocks:
            _append_block(
                document,
                block,
                dir_base=dir_base,
                style=style,
            )
    else:
        for block in markdown_blocks:
            _insert_block_before_anchor(
                document,
                anchor,
                block,
                dir_base=dir_base,
                style=style,
            )
        _remove_paragraph(anchor)


def _find_anchor_paragraph(
    document: DocxDocument,
    anchor_token: str,
) -> Paragraph | None:
    for paragraph in document.paragraphs:
        if paragraph.text.strip() == anchor_token:
            return paragraph
    return None


def _append_block(
    document: DocxDocument,
    block: MarkdownBlock,
    *,
    dir_base: Path,
    style: DocxStyle,
) -> None:
    match block:
        case MarkdownHeading(level=level, text=text):
            paragraph = document.add_heading(level=level)
            _write_heading(paragraph, text, level=level, style=style)
        case MarkdownParagraph(text=text):
            paragraph = document.add_paragraph()
            _apply_paragraph_style(paragraph, text=text, style=style)
            _write_text_with_line_breaks(
                paragraph,
                text,
                style=style,
                size_pt=_paragraph_text_size(text, style=style),
            )
        case MarkdownOrderedList(items=items):
            for item in items:
                paragraph = document.add_paragraph(style="List Number")
                _apply_ordered_list_style(paragraph)
                _write_text_with_line_breaks(paragraph, item, style=style)
        case MarkdownTable(rows=rows):
            _append_table(document, rows, style=style)
        case MarkdownImage(path=path, caption=caption, width_pct=width_pct):
            _append_image(
                document,
                dir_base / path,
                caption=caption,
                width_pct=width_pct,
                style=style,
            )
        case MarkdownPageBreak():
            document.add_page_break()
        case MarkdownSpacer():
            document.add_paragraph()


def _insert_block_before_anchor(
    document: DocxDocument,
    anchor: Paragraph,
    block: MarkdownBlock,
    *,
    dir_base: Path,
    style: DocxStyle,
) -> None:
    match block:
        case MarkdownHeading(level=level, text=text):
            paragraph = anchor.insert_paragraph_before(style=f"Heading {level}")
            _write_heading(paragraph, text, level=level, style=style)
        case MarkdownParagraph(text=text):
            paragraph = anchor.insert_paragraph_before()
            _apply_paragraph_style(paragraph, text=text, style=style)
            _write_text_with_line_breaks(
                paragraph,
                text,
                style=style,
                size_pt=_paragraph_text_size(text, style=style),
            )
        case MarkdownOrderedList(items=items):
            for item in items:
                paragraph = anchor.insert_paragraph_before(style="List Number")
                _apply_ordered_list_style(paragraph)
                _write_text_with_line_breaks(paragraph, item, style=style)
        case MarkdownTable(rows=rows):
            table = _append_table(document, rows, style=style)
            _insert_table_before_anchor(anchor, table)
        case MarkdownImage(path=path, caption=caption, width_pct=width_pct):
            paragraph_image = anchor.insert_paragraph_before()
            _add_picture(paragraph_image, dir_base / path, width_pct=width_pct)
            if caption:
                paragraph_caption = anchor.insert_paragraph_before()
                paragraph_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _write_text_with_line_breaks(
                    paragraph_caption,
                    caption,
                    style=style,
                    size_pt=style.sizes.pt_caption,
                )
        case MarkdownPageBreak():
            anchor.insert_paragraph_before().add_run().add_break(WD_BREAK.PAGE)
        case MarkdownSpacer():
            anchor.insert_paragraph_before()


def _write_heading(
    paragraph: Paragraph,
    text: str,
    *,
    level: int,
    style: DocxStyle,
) -> None:
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.first_line_indent = None
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run(text)
    _apply_run_style(
        run,
        style=style,
        size_pt=style.sizes.pt_heading_by_level.get(level, style.sizes.pt_body),
        east_asia_font=style.fonts.font_name_heading_east_asia,
    )
    run.bold = True


def _write_text_with_line_breaks(
    paragraph: Paragraph,
    text: str,
    *,
    style: DocxStyle,
    size_pt: float | None = None,
) -> None:
    size_effective = style.sizes.pt_body if size_pt is None else size_pt
    for idx, line in enumerate(text.split("\n")):
        if idx > 0:
            paragraph.add_run().add_break(WD_BREAK.LINE)
        run = paragraph.add_run(line)
        _apply_run_style(run, style=style, size_pt=size_effective)


def _apply_run_style(
    run: Run,
    *,
    style: DocxStyle,
    size_pt: float,
    east_asia_font: str | None = None,
) -> None:
    run.font.name = style.fonts.font_name_latin
    run.font.size = Pt(size_pt)
    run_properties = cast(Any, run)._element.get_or_add_rPr()
    run_fonts = run_properties.get_or_add_rFonts()
    run_fonts.set(
        qn("w:eastAsia"),
        east_asia_font or style.fonts.font_name_body_east_asia,
    )


def _apply_paragraph_style(
    paragraph: Paragraph,
    *,
    text: str,
    style: DocxStyle,
) -> None:
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    if _is_note_text(text, style=style):
        paragraph.paragraph_format.first_line_indent = None
        paragraph.paragraph_format.line_spacing = style.paragraph.line_spacing_note
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return
    paragraph.paragraph_format.first_line_indent = Cm(
        style.paragraph.first_line_indent_cm
    )
    paragraph.paragraph_format.line_spacing = style.paragraph.line_spacing_body


def _apply_ordered_list_style(paragraph: Paragraph) -> None:
    paragraph.paragraph_format.left_indent = Cm(0.74)
    paragraph.paragraph_format.first_line_indent = Cm(-0.74)


def _is_note_text(text: str, *, style: DocxStyle) -> bool:
    return text.strip().startswith(style.paragraph.note_prefixes)


def _paragraph_text_size(text: str, *, style: DocxStyle) -> float:
    if _is_note_text(text, style=style):
        return style.sizes.pt_caption
    return style.sizes.pt_body


def _append_table(
    document: DocxDocument,
    rows: tuple[tuple[str, ...], ...],
    *,
    style: DocxStyle,
) -> Table:
    if not rows:
        return document.add_table(rows=0, cols=0)
    count_cols = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=count_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    _apply_three_line_table_borders(table, style=style)
    for row_idx, row in enumerate(rows):
        for col_idx, value in enumerate(row):
            paragraph = table.cell(row_idx, col_idx).paragraphs[0]
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.line_spacing = style.table.line_spacing
            _write_text_with_line_breaks(
                paragraph,
                value,
                style=style,
                size_pt=style.sizes.pt_table,
            )
            if row_idx == 0:
                for run in paragraph.runs:
                    run.bold = True
    return table


def _append_image(
    document: DocxDocument,
    path_image: Path,
    *,
    caption: str,
    width_pct: float,
    style: DocxStyle,
) -> None:
    paragraph = document.add_paragraph()
    _add_picture(paragraph, path_image, width_pct=width_pct)
    if caption:
        paragraph_caption = document.add_paragraph()
        paragraph_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _write_text_with_line_breaks(
            paragraph_caption,
            caption,
            style=style,
            size_pt=style.sizes.pt_caption,
        )


def _add_picture(paragraph: Paragraph, path_image: Path, *, width_pct: float) -> None:
    width_inches = 6.0 * max(0.1, min(width_pct / 100.0, 1.0))
    paragraph.add_run().add_picture(str(path_image), width=Inches(width_inches))


def _insert_table_before_anchor(anchor: Paragraph, table: Table) -> None:
    cast(Any, anchor)._p.addprevious(cast(Any, table)._tbl)


def _remove_paragraph(paragraph: Paragraph) -> None:
    element = cast(Any, paragraph)._element
    element.getparent().remove(element)


def _apply_three_line_table_borders(table: Table, *, style: DocxStyle) -> None:
    for row in table.rows:
        for cell in row.cells:
            _set_cell_border(cell, edge_name="top", value="nil", size="0", style=style)
            _set_cell_border(
                cell,
                edge_name="bottom",
                value="nil",
                size="0",
                style=style,
            )
    if not table.rows:
        return
    for cell in table.rows[0].cells:
        _set_cell_border(
            cell,
            edge_name="top",
            value="single",
            size=style.table.border_size_main,
            style=style,
        )
        _set_cell_border(
            cell,
            edge_name="bottom",
            value="single",
            size=style.table.border_size_header,
            style=style,
        )
    if len(table.rows) > 1:
        for cell in table.rows[-1].cells:
            _set_cell_border(
                cell,
                edge_name="bottom",
                value="single",
                size=style.table.border_size_main,
                style=style,
            )


def _set_cell_border(
    cell: Any,
    *,
    edge_name: str,
    value: str,
    size: str,
    style: DocxStyle,
) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()  # noqa: SLF001
    borders: Any = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = cast(Any, OxmlElement("w:tcBorders"))
        tc_pr.append(borders)
    edge: Any = borders.find(qn(f"w:{edge_name}"))
    if edge is None:
        edge = cast(Any, OxmlElement(f"w:{edge_name}"))
        borders.append(edge)
    _set_xml_border(edge, value=value, size=size, color=style.table.border_color)


def _set_xml_border(edge: Any, *, value: str, size: str, color: str) -> None:
    edge.set(qn("w:val"), value)
    edge.set(qn("w:sz"), size)
    edge.set(qn("w:space"), "0")
    edge.set(qn("w:color"), color)
