from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True, slots=True)
class DocxFontStyle:
    """Font names used when writing DOCX content.

    Attributes:
        font_name_latin (str): Latin font name applied to runs.
        font_name_body_east_asia (str): East Asian font name applied to body text.
        font_name_heading_east_asia (str): East Asian font name applied to headings.
    """

    font_name_latin: str
    font_name_body_east_asia: str
    font_name_heading_east_asia: str


@dataclass(frozen=True, slots=True)
class DocxSizeStyle:
    """Point sizes used by the DOCX writer.

    Attributes:
        pt_title_page_title (float): Title-page report title size.
        pt_title_page_meta (float): Title-page metadata size.
        pt_title_page_compiler (float): Title-page compiler or organization text size.
        pt_body (float): Body paragraph text size.
        pt_caption (float): Figure caption and note text size.
        pt_table (float): Markdown table body text size.
        pt_heading_by_level (Mapping[int, float]): Heading text size by heading level.
    """

    pt_title_page_title: float
    pt_title_page_meta: float
    pt_title_page_compiler: float
    pt_body: float
    pt_caption: float
    pt_table: float
    pt_heading_by_level: Mapping[int, float]

    def with_overrides(
        self,
        *,
        pt_title_page_title: float | None = None,
        pt_title_page_meta: float | None = None,
        pt_title_page_compiler: float | None = None,
        pt_body: float | None = None,
        pt_caption: float | None = None,
        pt_table: float | None = None,
        pt_heading_by_level: Mapping[int, float] | None = None,
    ) -> DocxSizeStyle:
        """Create a copy with selected size values overridden.

        Args:
            pt_title_page_title (float | None): Title-page report title size.
            pt_title_page_meta (float | None): Title-page metadata size.
            pt_title_page_compiler (float | None): Compiler or organization size.
            pt_body (float | None): Body paragraph text size.
            pt_caption (float | None): Caption and note text size.
            pt_table (float | None): Markdown table text size.
            pt_heading_by_level (Mapping[int, float] | None): Heading sizes by level.

        Returns:
            DocxSizeStyle: New size style with overrides applied.
        """

        return DocxSizeStyle(
            pt_title_page_title=(
                pt_title_page_title
                if pt_title_page_title is not None
                else self.pt_title_page_title
            ),
            pt_title_page_meta=(
                pt_title_page_meta
                if pt_title_page_meta is not None
                else self.pt_title_page_meta
            ),
            pt_title_page_compiler=(
                pt_title_page_compiler
                if pt_title_page_compiler is not None
                else self.pt_title_page_compiler
            ),
            pt_body=pt_body if pt_body is not None else self.pt_body,
            pt_caption=pt_caption if pt_caption is not None else self.pt_caption,
            pt_table=pt_table if pt_table is not None else self.pt_table,
            pt_heading_by_level=(
                dict(pt_heading_by_level)
                if pt_heading_by_level is not None
                else self.pt_heading_by_level
            ),
        )


@dataclass(frozen=True, slots=True)
class DocxTableStyle:
    """Table style values used by markdown table rendering.

    Attributes:
        border_color (str): WordprocessingML border color as a hex RGB string.
        stripe_fill_color (str): Body-row stripe fill color as a hex RGB string.
        border_size_main (str): Main top/bottom border size in Word units.
        border_size_header (str): Header separator border size in Word units.
        line_spacing (float): Line spacing applied to table cell paragraphs.
    """

    border_color: str
    stripe_fill_color: str
    border_size_main: str
    border_size_header: str
    line_spacing: float


@dataclass(frozen=True, slots=True)
class DocxParagraphStyle:
    """Paragraph style values used by body and note paragraphs.

    Attributes:
        line_spacing_body (float): Line spacing for ordinary body paragraphs.
        line_spacing_note (float): Line spacing for note paragraphs.
        first_line_indent_cm (float): Body paragraph first-line indent in centimeters.
        note_prefixes (tuple[str, ...]): Text prefixes classified as note paragraphs.
    """

    line_spacing_body: float
    line_spacing_note: float
    first_line_indent_cm: float
    note_prefixes: tuple[str, ...] = ("注：", "注:")


@dataclass(frozen=True, slots=True)
class DocxStyle:
    """Complete style bundle for DOCX writing.

    `docxkit` accepts this structured style object directly. Reading TOML,
    JSON, YAML, or another configuration format is the caller's responsibility.

    Attributes:
        fonts (DocxFontStyle): Font names for Latin and East Asian text.
        sizes (DocxSizeStyle): Point sizes for title, body, caption, table, and
            headings.
        table (DocxTableStyle): Table border, fill, and spacing settings.
        paragraph (DocxParagraphStyle): Body and note paragraph settings.
    """

    fonts: DocxFontStyle
    sizes: DocxSizeStyle
    table: DocxTableStyle
    paragraph: DocxParagraphStyle


@dataclass(frozen=True, slots=True)
class DocxFieldRefreshOptions:
    """Options for refreshing DOCX fields through LibreOffice UNO.

    Attributes:
        exe_libreoffice (Path): LibreOffice executable path.
        dir_user_profile (Path): Isolated LibreOffice user profile directory.
        file_out_docx_refreshed (Path | None): Optional refreshed DOCX output path.
        file_listener_log (Path | None): Optional LibreOffice listener log path.
        should_require_toc (bool): Whether refreshed DOCX must contain TOC results.
        should_freeze_fields (bool): Whether refreshed field results should be frozen.
        timeout_seconds (float): Maximum wait time for refreshed DOCX validation.
        poll_interval_seconds (float): Poll interval for refreshed DOCX validation.
        stable_checks (int): Consecutive stable file-stat checks required.
    """

    exe_libreoffice: Path
    dir_user_profile: Path
    file_out_docx_refreshed: Path | None = None
    file_listener_log: Path | None = None
    should_require_toc: bool = False
    should_freeze_fields: bool = False
    timeout_seconds: float = 30.0
    poll_interval_seconds: float = 0.5
    stable_checks: int = 2


@dataclass(frozen=True, slots=True)
class DocxWriteOptions:
    """Inputs for writing a DOCX file.

    Attributes:
        file_template (Path): Input DOCX template path.
        file_out_docx (Path): Output DOCX path to write.
        context (Mapping[str, Any]): Template context passed to `docxtpl`.
        markdown_body (str): Already-rendered markdown body to insert into the DOCX.
        dir_base (Path): Base directory used to resolve relative image paths.
        style (DocxStyle): Structured DOCX style settings.
        anchor_token (str): Paragraph text marking where markdown body content is
            inserted.
        should_update_fields (bool): Whether DOCX fields should be prepared for update.
        should_freeze_fields (bool): Whether DOCX fields should be frozen after writing.
        field_refresh (DocxFieldRefreshOptions | None): Optional UNO field refresh
            settings.
    """

    file_template: Path
    file_out_docx: Path
    context: Mapping[str, Any]
    markdown_body: str
    dir_base: Path
    style: DocxStyle
    anchor_token: str = "__REPORT_BODY_ANCHOR__"
    should_update_fields: bool = True
    should_freeze_fields: bool = False
    field_refresh: DocxFieldRefreshOptions | None = None


@dataclass(frozen=True, slots=True)
class DocxWriteResult:
    """Result of a DOCX write operation.

    Attributes:
        file_docx (Path): Written DOCX path.
    """

    file_docx: Path


@dataclass(frozen=True, slots=True)
class DocxToPdfOptions:
    """Inputs for converting a DOCX file to PDF.

    Attributes:
        exe_libreoffice (Path): LibreOffice executable path.
        file_in_docx (Path): Input DOCX path.
        file_out_pdf (Path): Output PDF path.
        dir_user_profile (Path): Isolated LibreOffice user profile directory.
        file_out_docx_refreshed (Path | None): Optional refreshed DOCX output path.
        file_listener_log (Path | None): Optional LibreOffice listener log path.
    """

    exe_libreoffice: Path
    file_in_docx: Path
    file_out_pdf: Path
    dir_user_profile: Path
    file_out_docx_refreshed: Path | None = None
    file_listener_log: Path | None = None


@dataclass(frozen=True, slots=True)
class DocxToPdfResult:
    """Result of converting DOCX to PDF.

    Attributes:
        file_pdf (Path): Written PDF path.
        file_docx_refreshed (Path | None): Refreshed DOCX path when requested.
    """

    file_pdf: Path
    file_docx_refreshed: Path | None = None
