from __future__ import annotations

import re
from pathlib import Path
from typing import Any, cast

from docx.document import Document as DocxDocument
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement  # pyright: ignore[reportUnknownVariableType]
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docxrender.contracts import (
    DocxBodyAnchorOptions,
    DocxBodyRenderPolicy,
    DocxStyle,
)
from docxrender.markdown import (
    MarkdownBlock,
    MarkdownHeading,
    MarkdownImage,
    MarkdownOrderedList,
    MarkdownPageBreak,
    MarkdownParagraph,
    MarkdownSpacer,
    MarkdownTable,
    MarkdownText,
    MarkdownTextSegment,
    MarkdownUnorderedList,
)


def insert_markdown_blocks(
    document: DocxDocument,
    markdown_blocks: tuple[MarkdownBlock, ...],
    *,
    body_anchor: DocxBodyAnchorOptions,
    dir_base: Path,
    style: DocxStyle,
    body_render_policy: DocxBodyRenderPolicy = DocxBodyRenderPolicy(),
) -> None:
    anchor = find_body_anchor_paragraph(document, body_anchor)
    blocks_effective = _apply_heading_numbers(
        markdown_blocks,
        policy=body_render_policy,
    )
    if anchor is None:
        for block in blocks_effective:
            _append_block(
                document,
                block,
                dir_base=dir_base,
                style=style,
                policy=body_render_policy,
            )
    else:
        for block in blocks_effective:
            _insert_block_before_anchor(
                document,
                anchor,
                block,
                dir_base=dir_base,
                style=style,
                policy=body_render_policy,
            )
        if (
            body_anchor.should_preserve_section_properties
            and has_section_properties(anchor)
        ):
            clear_paragraph_runs(anchor)
        else:
            remove_paragraph(anchor)


def find_body_anchor_paragraphs(
    document: DocxDocument,
    options: DocxBodyAnchorOptions,
) -> tuple[Paragraph, ...]:
    """Find body anchor paragraphs in top-level document paragraphs.

    Args:
        document (DocxDocument): DOCX document to search.
        options (DocxBodyAnchorOptions): Anchor search options.

    Returns:
        tuple[Paragraph, ...]: Matching paragraphs in document order.
    """

    matches: list[Paragraph] = []
    for paragraph in document.paragraphs:
        if _is_body_anchor_paragraph(paragraph, options):
            matches.append(paragraph)
    return tuple(matches)


def find_body_anchor_paragraph(
    document: DocxDocument,
    options: DocxBodyAnchorOptions,
) -> Paragraph | None:
    """Find the unique body anchor paragraph according to options.

    Args:
        document (DocxDocument): DOCX document to search.
        options (DocxBodyAnchorOptions): Anchor search options.

    Returns:
        Paragraph | None: Unique anchor paragraph, or `None` when missing and
            `rule_missing="append"`.

    Raises:
        ValueError: The anchor is missing with `rule_missing="raise"`, the anchor
            is duplicated, or an unsupported rule is supplied.
    """

    matches = find_body_anchor_paragraphs(document, options)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError(
            "Body anchor token matched multiple paragraphs: "
            f"anchor_token={options.anchor_token!r} count={len(matches)}"
        )
    if options.rule_missing == "append":
        return None
    if options.rule_missing == "raise":
        raise ValueError(
            "Could not locate body anchor token in document: "
            f"anchor_token={options.anchor_token!r} rule_match={options.rule_match!r}"
        )
    raise ValueError(f"Unsupported body anchor missing rule: {options.rule_missing!r}")


def _is_body_anchor_paragraph(
    paragraph: Paragraph,
    options: DocxBodyAnchorOptions,
) -> bool:
    if options.rule_match == "equals":
        return paragraph.text.strip() == options.anchor_token
    if options.rule_match == "contains":
        return options.anchor_token in paragraph.text
    raise ValueError(f"Unsupported body anchor match rule: {options.rule_match!r}")


def _append_block(
    document: DocxDocument,
    block: MarkdownBlock,
    *,
    dir_base: Path,
    style: DocxStyle,
    policy: DocxBodyRenderPolicy,
) -> None:
    match block:
        case MarkdownHeading(level=level, text=text):
            paragraph = document.add_heading(level=level)
            _write_heading(paragraph, text, level=level, style=style)
        case MarkdownParagraph(text=text):
            paragraph = document.add_paragraph()
            _apply_paragraph_style(paragraph, text=text, style=style)
            _write_markdown_text(
                paragraph,
                text,
                style=style,
                size_pt=_paragraph_text_size(text, style=style),
            )
        case MarkdownUnorderedList(items=items):
            for item in items:
                paragraph = _add_unordered_list_paragraph(document, item, policy=policy)
                _write_markdown_text(paragraph, item, style=style)
        case MarkdownOrderedList(items=items):
            for item_idx, item in enumerate(items, start=1):
                paragraph = _add_ordered_list_paragraph(
                    document,
                    item,
                    item_idx=item_idx,
                    policy=policy,
                )
                _write_markdown_text(paragraph, item, style=style)
        case MarkdownTable(rows=rows):
            _append_table(document, rows, style=style, policy=policy)
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
    policy: DocxBodyRenderPolicy,
) -> None:
    match block:
        case MarkdownHeading(level=level, text=text):
            paragraph = anchor.insert_paragraph_before(style=f"Heading {level}")
            _write_heading(paragraph, text, level=level, style=style)
        case MarkdownParagraph(text=text):
            paragraph = anchor.insert_paragraph_before()
            _apply_paragraph_style(paragraph, text=text, style=style)
            _write_markdown_text(
                paragraph,
                text,
                style=style,
                size_pt=_paragraph_text_size(text, style=style),
            )
        case MarkdownUnorderedList(items=items):
            for item in items:
                paragraph = _insert_unordered_list_paragraph_before(
                    anchor,
                    item,
                    policy=policy,
                )
                _write_markdown_text(paragraph, item, style=style)
        case MarkdownOrderedList(items=items):
            for item_idx, item in enumerate(items, start=1):
                paragraph = _insert_ordered_list_paragraph_before(
                    anchor,
                    item,
                    item_idx=item_idx,
                    policy=policy,
                )
                _write_markdown_text(paragraph, item, style=style)
        case MarkdownTable(rows=rows):
            table = _append_table(document, rows, style=style, policy=policy)
            _insert_table_before_anchor(anchor, table)
        case MarkdownImage(path=path, caption=caption, width_pct=width_pct):
            paragraph_image = anchor.insert_paragraph_before()
            paragraph_image.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _add_picture(paragraph_image, dir_base / path, width_pct=width_pct)
            if caption:
                paragraph_caption = anchor.insert_paragraph_before()
                _apply_image_caption_paragraph_style(paragraph_caption)
                _write_markdown_text(
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
    text: MarkdownText,
    *,
    level: int,
    style: DocxStyle,
) -> None:
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.first_line_indent = None
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _write_markdown_text(
        paragraph,
        text,
        style=style,
        size_pt=style.sizes.pt_heading_by_level.get(level, style.sizes.pt_body),
        east_asia_font=style.fonts.font_name_heading_east_asia,
        should_force_bold=True,
    )


def _write_markdown_text(
    paragraph: Paragraph,
    text: MarkdownText,
    *,
    style: DocxStyle,
    size_pt: float | None = None,
    east_asia_font: str | None = None,
    should_force_bold: bool = False,
) -> None:
    size_effective = style.sizes.pt_body if size_pt is None else size_pt
    is_first_run = True
    for segment in text:
        is_first_line = True
        for line in segment.text.split("\n"):
            if not is_first_run and is_first_line:
                pass
            if not is_first_line:
                paragraph.add_run().add_break(WD_BREAK.LINE)
            run = paragraph.add_run(line)
            _apply_run_style(
                run,
                style=style,
                size_pt=size_effective,
                east_asia_font=east_asia_font,
            )
            run.bold = should_force_bold or segment.is_bold
            is_first_run = False
            is_first_line = False


def _plain_text(text: MarkdownText) -> str:
    return "".join(segment.text for segment in text)


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
    text: MarkdownText,
    style: DocxStyle,
) -> None:
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    text_plain = _plain_text(text)
    if _is_note_text(text_plain, style=style):
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


def _add_unordered_list_paragraph(
    document: DocxDocument,
    item: MarkdownText,
    *,
    policy: DocxBodyRenderPolicy,
) -> Paragraph:
    if policy.rule_unordered_list == "word_style":
        paragraph = document.add_paragraph(style="List Bullet")
        _apply_ordered_list_style(paragraph)
        return paragraph
    if policy.rule_unordered_list == "plain_text":
        return document.add_paragraph("- ")
    raise ValueError(f"Unsupported unordered list rule: {policy.rule_unordered_list!r}")


def _insert_unordered_list_paragraph_before(
    anchor: Paragraph,
    item: MarkdownText,
    *,
    policy: DocxBodyRenderPolicy,
) -> Paragraph:
    if policy.rule_unordered_list == "word_style":
        paragraph = anchor.insert_paragraph_before(style="List Bullet")
        _apply_ordered_list_style(paragraph)
        return paragraph
    if policy.rule_unordered_list == "plain_text":
        return anchor.insert_paragraph_before("- ")
    raise ValueError(f"Unsupported unordered list rule: {policy.rule_unordered_list!r}")


def _add_ordered_list_paragraph(
    document: DocxDocument,
    item: MarkdownText,
    *,
    item_idx: int,
    policy: DocxBodyRenderPolicy,
) -> Paragraph:
    if policy.rule_ordered_list == "word_style":
        paragraph = document.add_paragraph(style="List Number")
        _apply_ordered_list_style(paragraph)
        return paragraph
    if policy.rule_ordered_list == "plain_text":
        return document.add_paragraph(f"{item_idx}. ")
    raise ValueError(f"Unsupported ordered list rule: {policy.rule_ordered_list!r}")


def _insert_ordered_list_paragraph_before(
    anchor: Paragraph,
    item: MarkdownText,
    *,
    item_idx: int,
    policy: DocxBodyRenderPolicy,
) -> Paragraph:
    if policy.rule_ordered_list == "word_style":
        paragraph = anchor.insert_paragraph_before(style="List Number")
        _apply_ordered_list_style(paragraph)
        return paragraph
    if policy.rule_ordered_list == "plain_text":
        return anchor.insert_paragraph_before(f"{item_idx}. ")
    raise ValueError(f"Unsupported ordered list rule: {policy.rule_ordered_list!r}")


def _is_note_text(text: str, *, style: DocxStyle) -> bool:
    return text.strip().startswith(style.paragraph.note_prefixes)


def _paragraph_text_size(text: MarkdownText, *, style: DocxStyle) -> float:
    if _is_note_text(_plain_text(text), style=style):
        return style.sizes.pt_caption
    return style.sizes.pt_body


def _append_table(
    document: DocxDocument,
    rows: tuple[tuple[MarkdownText, ...], ...],
    *,
    style: DocxStyle,
    policy: DocxBodyRenderPolicy,
) -> Table:
    if not rows:
        return document.add_table(rows=0, cols=0)
    count_cols = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=count_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    _apply_content_autofit_table_layout(table)
    _apply_three_line_table_borders(table, style=style)
    if policy.should_stripe_table_rows:
        _apply_table_body_stripes(table, style=style)
    for row_idx, row in enumerate(rows):
        for col_idx, value in enumerate(row):
            cell = table.cell(row_idx, col_idx)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = style.table.line_spacing
            _write_markdown_text(
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
    caption: MarkdownText,
    width_pct: float,
    style: DocxStyle,
) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_picture(paragraph, path_image, width_pct=width_pct)
    if caption:
        paragraph_caption = document.add_paragraph()
        _apply_image_caption_paragraph_style(paragraph_caption)
        _write_markdown_text(
            paragraph_caption,
            caption,
            style=style,
            size_pt=style.sizes.pt_caption,
        )


def _add_picture(paragraph: Paragraph, path_image: Path, *, width_pct: float) -> None:
    width_inches = 6.0 * max(0.1, min(width_pct / 100.0, 1.0))
    paragraph.add_run().add_picture(str(path_image), width=Inches(width_inches))


def _apply_content_autofit_table_layout(table: Table) -> None:
    table.autofit = True
    table_element = cast(Any, table)._tbl
    tbl_pr = table_element.tblPr
    tbl_layout: Any = tbl_pr.first_child_found_in("w:tblLayout")
    if tbl_layout is None:
        tbl_layout = cast(Any, OxmlElement("w:tblLayout"))
        tbl_pr.append(tbl_layout)
    tbl_layout.set(qn("w:type"), "autofit")
    tbl_width: Any = tbl_pr.first_child_found_in("w:tblW")
    if tbl_width is not None:
        tbl_width.set(qn("w:type"), "auto")
        tbl_width.set(qn("w:w"), "0")
    for grid_col in table_element.tblGrid.gridCol_lst:
        if qn("w:w") in grid_col.attrib:
            del grid_col.attrib[qn("w:w")]
    for cell in table_element.iter_tcs():
        tc_pr = cell.get_or_add_tcPr()
        tc_width: Any = tc_pr.first_child_found_in("w:tcW")
        if tc_width is not None:
            tc_pr.remove(tc_width)


def _insert_table_before_anchor(anchor: Paragraph, table: Table) -> None:
    cast(Any, anchor)._p.addprevious(cast(Any, table)._tbl)


def _apply_image_caption_paragraph_style(paragraph: Paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.right_indent = Cm(0)
    paragraph.paragraph_format.first_line_indent = None
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.2


def has_section_properties(paragraph: Paragraph) -> bool:
    """Return whether a paragraph carries DOCX section properties.

    Args:
        paragraph (Paragraph): Paragraph to inspect.

    Returns:
        bool: Whether the paragraph contains `w:sectPr`.
    """

    paragraph_element = cast(Any, paragraph)._p
    return (
        paragraph_element.pPr is not None
        and paragraph_element.pPr.sectPr is not None
    )


def clear_paragraph_runs(paragraph: Paragraph) -> None:
    """Remove all runs from a paragraph while keeping paragraph properties.

    Args:
        paragraph (Paragraph): Paragraph to clear.
    """

    for run in list(paragraph.runs):
        element = cast(Any, run)._element
        element.getparent().remove(element)


def remove_paragraph(paragraph: Paragraph) -> None:
    """Remove a paragraph element from its parent.

    Args:
        paragraph (Paragraph): Paragraph to remove.
    """

    element = cast(Any, paragraph)._element
    element.getparent().remove(element)


def _apply_three_line_table_borders(table: Table, *, style: DocxStyle) -> None:
    for edge_name in ("top", "bottom", "left", "right", "insideV", "insideH"):
        _set_table_border(
            table,
            edge_name=edge_name,
            value="nil",
            size="0",
            style=style,
        )
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


def _set_table_border(
    table: Table,
    *,
    edge_name: str,
    value: str,
    size: str,
    style: DocxStyle,
) -> None:
    tbl_pr = cast(Any, table)._tbl.tblPr
    borders: Any = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = cast(Any, OxmlElement("w:tblBorders"))
        tbl_pr.append(borders)
    edge: Any = borders.find(qn(f"w:{edge_name}"))
    if edge is None:
        edge = cast(Any, OxmlElement(f"w:{edge_name}"))
        borders.append(edge)
    _set_xml_border(edge, value=value, size=size, color=style.table.border_color)


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


def _apply_table_body_stripes(table: Table, *, style: DocxStyle) -> None:
    for row_idx, row in enumerate(table.rows[1:], start=1):
        if row_idx % 2 == 0:
            continue
        for cell in row.cells:
            _set_cell_shading(cell, fill=style.table.stripe_fill_color)


def _set_cell_shading(cell: Any, *, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()  # noqa: SLF001
    shading: Any = tc_pr.first_child_found_in("w:shd")
    if shading is None:
        shading = cast(Any, OxmlElement("w:shd"))
        tc_pr.append(shading)
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)


def _apply_heading_numbers(
    blocks: tuple[MarkdownBlock, ...],
    *,
    policy: DocxBodyRenderPolicy,
) -> tuple[MarkdownBlock, ...]:
    if not policy.should_number_headings:
        return blocks
    counts = [0, 0, 0, 0, 0, 0]
    blocks_out: list[MarkdownBlock] = []
    for block in blocks:
        if not isinstance(block, MarkdownHeading):
            blocks_out.append(block)
            continue
        level = min(max(block.level, 1), len(counts))
        counts[level - 1] += 1
        for idx in range(level, len(counts)):
            counts[idx] = 0
        text_plain = _plain_text(block.text)
        if re.match(r"^\d+(?:\.\d+)*\.?\s+", text_plain):
            blocks_out.append(block)
            continue
        prefix = ".".join(str(value) for value in counts[:level] if value > 0)
        text_prefixed = (
            f"{prefix}. {text_plain}" if level == 1 else f"{prefix} {text_plain}"
        )
        blocks_out.append(
            MarkdownHeading(
                level=block.level,
                text=(MarkdownTextSegment(text_prefixed),),
            )
        )
    return tuple(blocks_out)
