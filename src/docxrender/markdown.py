from __future__ import annotations

import re
from dataclasses import dataclass

from docxrender.contracts import DocxMarkdownOptions


@dataclass(frozen=True, slots=True)
class MarkdownTextSegment:
    text: str
    is_bold: bool = False


MarkdownText = tuple[MarkdownTextSegment, ...]


@dataclass(frozen=True, slots=True)
class MarkdownHeading:
    level: int
    text: MarkdownText


@dataclass(frozen=True, slots=True)
class MarkdownParagraph:
    text: MarkdownText


@dataclass(frozen=True, slots=True)
class MarkdownUnorderedList:
    items: tuple[MarkdownText, ...]


@dataclass(frozen=True, slots=True)
class MarkdownOrderedList:
    items: tuple[MarkdownText, ...]


@dataclass(frozen=True, slots=True)
class MarkdownTable:
    rows: tuple[tuple[MarkdownText, ...], ...]


@dataclass(frozen=True, slots=True)
class MarkdownImage:
    caption: MarkdownText
    path: str
    width_pct: float


@dataclass(frozen=True, slots=True)
class MarkdownPageBreak:
    pass


@dataclass(frozen=True, slots=True)
class MarkdownSpacer:
    pass


MarkdownBlock = (
    MarkdownHeading
    | MarkdownParagraph
    | MarkdownUnorderedList
    | MarkdownOrderedList
    | MarkdownTable
    | MarkdownImage
    | MarkdownPageBreak
    | MarkdownSpacer
)

RE_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
RE_UNORDERED_LIST_ITEM = re.compile(r"^[-*]\s+(.*)$")
RE_ORDERED_LIST_ITEM = re.compile(r"^\d+\.\s+(.*)$")
RE_IMAGE = re.compile(
    r"^!\[(?P<caption>.*?)\]\((?P<path>.*?)\)"
    r"(?:\{[^}]*width=(?P<width>\d+)%[^}]*\})?\s*$"
)
RE_TABLE_DELIMITER = re.compile(r"^:?-{3,}:?$")
RE_INLINE_BOLD = re.compile(r"\*\*(.+?)\*\*")
RE_INLINE_CODE = re.compile(r"`([^`]+)`")
RE_INLINE_LINK = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")


def parse_markdown_blocks(
    markdown_body: str,
    *,
    options: DocxMarkdownOptions | None = None,
) -> tuple[MarkdownBlock, ...]:
    options_effective = DocxMarkdownOptions() if options is None else options
    lines = markdown_body.splitlines()
    blocks: list[MarkdownBlock] = []
    idx = 0
    while idx < len(lines):
        line = lines[idx]
        text = line.strip()
        if not text:
            idx += 1
            continue
        if text == r"\newpage":
            blocks.append(MarkdownPageBreak())
            idx += 1
            continue
        if text == r"\vspace":
            blocks.append(MarkdownSpacer())
            idx += 1
            continue

        match_heading = RE_HEADING.match(text)
        if match_heading is not None:
            blocks.append(
                MarkdownHeading(
                    level=len(match_heading.group(1)),
                    text=parse_markdown_text(
                        match_heading.group(2).strip(),
                        options=options_effective,
                    ),
                )
            )
            idx += 1
            continue

        match_image = RE_IMAGE.match(text)
        if match_image is not None:
            width_raw = match_image.group("width")
            blocks.append(
                MarkdownImage(
                    caption=parse_markdown_text(
                        match_image.group("caption"),
                        options=options_effective,
                    ),
                    path=match_image.group("path").strip(),
                    width_pct=_parse_image_width_pct(
                        width_raw,
                        options=options_effective,
                    ),
                )
            )
            idx += 1
            continue

        match_unordered_item = RE_UNORDERED_LIST_ITEM.match(text)
        if match_unordered_item is not None:
            items: list[MarkdownText] = []
            while idx < len(lines):
                match_current = RE_UNORDERED_LIST_ITEM.match(lines[idx].strip())
                if match_current is None:
                    break
                items.append(
                    parse_markdown_text(
                        match_current.group(1).strip(),
                        options=options_effective,
                    )
                )
                idx += 1
            blocks.append(MarkdownUnorderedList(items=tuple(items)))
            continue

        match_list_item = RE_ORDERED_LIST_ITEM.match(text)
        if match_list_item is not None:
            items: list[MarkdownText] = []
            while idx < len(lines):
                match_current = RE_ORDERED_LIST_ITEM.match(lines[idx].strip())
                if match_current is None:
                    break
                items.append(
                    parse_markdown_text(
                        match_current.group(1).strip(),
                        options=options_effective,
                    )
                )
                idx += 1
            blocks.append(MarkdownOrderedList(items=tuple(items)))
            continue

        if _is_table_start(lines, idx):
            rows, idx = _parse_table_rows(lines, idx, options=options_effective)
            blocks.append(MarkdownTable(rows=rows))
            continue

        paragraph_parts: list[tuple[str, bool]] = []
        while idx < len(lines):
            line_current = lines[idx]
            text_current = line_current.strip()
            if not text_current:
                break
            if _is_table_start(lines, idx):
                break
            if _is_special_line(text_current):
                break
            line_text, has_hard_break = _strip_hard_break(line_current)
            paragraph_parts.append((line_text.strip(), has_hard_break))
            idx += 1
            if idx < len(lines) and lines[idx].strip() and not _is_special_line(
                lines[idx].strip()
            ):
                continue
            break
        blocks.append(
            MarkdownParagraph(
                text=parse_markdown_text(
                    _join_paragraph_parts(paragraph_parts),
                    options=options_effective,
                )
            )
        )
    return tuple(blocks)


def parse_markdown_text(
    text: str,
    *,
    options: DocxMarkdownOptions,
) -> MarkdownText:
    cleaned = _clean_inline_text(text, options=options)
    if not options.should_parse_inline_bold:
        return (MarkdownTextSegment(cleaned),)
    segments: list[MarkdownTextSegment] = []
    idx = 0
    for match_bold in RE_INLINE_BOLD.finditer(cleaned):
        if match_bold.start() > idx:
            segments.append(MarkdownTextSegment(cleaned[idx : match_bold.start()]))
        segments.append(MarkdownTextSegment(match_bold.group(1), is_bold=True))
        idx = match_bold.end()
    if idx < len(cleaned):
        segments.append(MarkdownTextSegment(cleaned[idx:]))
    if not segments:
        return (MarkdownTextSegment(""),)
    return tuple(segment for segment in segments if segment.text or segment.is_bold)


def _strip_hard_break(line: str) -> tuple[str, bool]:
    has_hard_break = line.endswith("  ")
    return line.rstrip(), has_hard_break


def _join_paragraph_parts(parts: list[tuple[str, bool]]) -> str:
    if not parts:
        return ""
    text = parts[0][0]
    for previous, current in zip(parts, parts[1:], strict=False):
        text += ("\n" if previous[1] else " ") + current[0]
    return text


def _is_special_line(text: str) -> bool:
    return (
        text == r"\newpage"
        or text == r"\vspace"
        or RE_UNORDERED_LIST_ITEM.match(text) is not None
        or RE_HEADING.match(text) is not None
        or RE_IMAGE.match(text) is not None
        or RE_ORDERED_LIST_ITEM.match(text) is not None
    )


def _is_table_start(lines: list[str], idx: int) -> bool:
    if idx + 1 >= len(lines):
        return False
    line_header = lines[idx].strip()
    line_delimiter = lines[idx + 1].strip()
    return (
        line_header.startswith("|")
        and line_delimiter.startswith("|")
        and _is_table_delimiter_row(line_delimiter)
    )


def _parse_table_rows(
    lines: list[str],
    idx: int,
    *,
    options: DocxMarkdownOptions,
) -> tuple[tuple[tuple[MarkdownText, ...], ...], int]:
    rows: list[tuple[MarkdownText, ...]] = []
    idx_current = idx
    while idx_current < len(lines) and lines[idx_current].strip().startswith("|"):
        row_text = lines[idx_current].strip().strip("|")
        row = tuple(
            parse_markdown_text(cell.strip(), options=options)
            for cell in row_text.split("|")
        )
        if not (idx_current == idx + 1 and _is_table_delimiter_row(lines[idx_current])):
            rows.append(row)
        idx_current += 1
    return tuple(rows), idx_current


def _is_table_delimiter_row(line: str) -> bool:
    row = [cell.strip() for cell in line.strip().strip("|").split("|")]
    return bool(row) and all(RE_TABLE_DELIMITER.match(cell) for cell in row)


def _clean_inline_text(text: str, *, options: DocxMarkdownOptions) -> str:
    text_clean = text
    if options.should_parse_inline_code:
        text_clean = RE_INLINE_CODE.sub(r"\1", text_clean)
    if options.should_parse_links_as_text:
        text_clean = RE_INLINE_LINK.sub(r"\1", text_clean)
    return text_clean


def _parse_image_width_pct(
    width_raw: str | None,
    *,
    options: DocxMarkdownOptions,
) -> float:
    if width_raw is None or not options.should_parse_image_width_attr:
        return options.default_image_width_pct
    return float(width_raw)
