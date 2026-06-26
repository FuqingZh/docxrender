from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class MarkdownHeading:
    level: int
    text: str


@dataclass(frozen=True, slots=True)
class MarkdownParagraph:
    text: str


@dataclass(frozen=True, slots=True)
class MarkdownOrderedList:
    items: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class MarkdownTable:
    rows: tuple[tuple[str, ...], ...]


@dataclass(frozen=True, slots=True)
class MarkdownImage:
    caption: str
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
    | MarkdownOrderedList
    | MarkdownTable
    | MarkdownImage
    | MarkdownPageBreak
    | MarkdownSpacer
)

RE_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
RE_ORDERED_LIST_ITEM = re.compile(r"^\d+\.\s+(.*)$")
RE_IMAGE = re.compile(
    r"^!\[(?P<caption>.*?)\]\((?P<path>.*?)\)"
    r"(?:\{[^}]*width=(?P<width>\d+)%[^}]*\})?\s*$"
)


def parse_markdown_blocks(markdown_body: str) -> tuple[MarkdownBlock, ...]:
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
                    text=match_heading.group(2).strip(),
                )
            )
            idx += 1
            continue

        match_image = RE_IMAGE.match(text)
        if match_image is not None:
            width_raw = match_image.group("width")
            blocks.append(
                MarkdownImage(
                    caption=match_image.group("caption"),
                    path=match_image.group("path").strip(),
                    width_pct=float(width_raw) if width_raw is not None else 90.0,
                )
            )
            idx += 1
            continue

        match_list_item = RE_ORDERED_LIST_ITEM.match(text)
        if match_list_item is not None:
            items: list[str] = []
            while idx < len(lines):
                match_current = RE_ORDERED_LIST_ITEM.match(lines[idx].strip())
                if match_current is None:
                    break
                items.append(match_current.group(1).strip())
                idx += 1
            blocks.append(MarkdownOrderedList(items=tuple(items)))
            continue

        if text.startswith("|"):
            rows: list[tuple[str, ...]] = []
            idx_table = 0
            while idx < len(lines) and lines[idx].strip().startswith("|"):
                row = tuple(
                    cell.strip() for cell in lines[idx].strip().strip("|").split("|")
                )
                is_header_separator = idx_table == 1 and all(
                    set(cell) <= {"-", ":"} for cell in row
                )
                if not is_header_separator:
                    rows.append(row)
                idx += 1
                idx_table += 1
            blocks.append(MarkdownTable(rows=tuple(rows)))
            continue

        paragraph_parts: list[tuple[str, bool]] = []
        while idx < len(lines):
            line_current = lines[idx]
            text_current = line_current.strip()
            if not text_current:
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
        blocks.append(MarkdownParagraph(text=_join_paragraph_parts(paragraph_parts)))
    return tuple(blocks)


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
        or text.startswith("|")
        or RE_HEADING.match(text) is not None
        or RE_IMAGE.match(text) is not None
        or RE_ORDERED_LIST_ITEM.match(text) is not None
    )
