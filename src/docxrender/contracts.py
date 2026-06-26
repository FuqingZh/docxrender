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

    def with_overrides(
        self,
        *,
        font_name_latin: str | None = None,
        font_name_body_east_asia: str | None = None,
        font_name_heading_east_asia: str | None = None,
    ) -> DocxFontStyle:
        """Create a copy with selected font values overridden.

        Args:
            font_name_latin (str | None): Latin font name applied to runs.
            font_name_body_east_asia (str | None): East Asian body font name.
            font_name_heading_east_asia (str | None): East Asian heading font name.

        Returns:
            DocxFontStyle: New font style with overrides applied.
        """

        return DocxFontStyle(
            font_name_latin=(
                font_name_latin
                if font_name_latin is not None
                else self.font_name_latin
            ),
            font_name_body_east_asia=(
                font_name_body_east_asia
                if font_name_body_east_asia is not None
                else self.font_name_body_east_asia
            ),
            font_name_heading_east_asia=(
                font_name_heading_east_asia
                if font_name_heading_east_asia is not None
                else self.font_name_heading_east_asia
            ),
        )


@dataclass(frozen=True, slots=True)
class DocxSizeStyle:
    """Point sizes used by DOCX rendering.

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

    def with_overrides(
        self,
        *,
        border_color: str | None = None,
        stripe_fill_color: str | None = None,
        border_size_main: str | None = None,
        border_size_header: str | None = None,
        line_spacing: float | None = None,
    ) -> DocxTableStyle:
        """Create a copy with selected table style values overridden.

        Args:
            border_color (str | None): WordprocessingML border color.
            stripe_fill_color (str | None): Body-row stripe fill color.
            border_size_main (str | None): Main border size in Word units.
            border_size_header (str | None): Header border size in Word units.
            line_spacing (float | None): Table paragraph line spacing.

        Returns:
            DocxTableStyle: New table style with overrides applied.
        """

        return DocxTableStyle(
            border_color=(
                border_color if border_color is not None else self.border_color
            ),
            stripe_fill_color=(
                stripe_fill_color
                if stripe_fill_color is not None
                else self.stripe_fill_color
            ),
            border_size_main=(
                border_size_main
                if border_size_main is not None
                else self.border_size_main
            ),
            border_size_header=(
                border_size_header
                if border_size_header is not None
                else self.border_size_header
            ),
            line_spacing=(
                line_spacing if line_spacing is not None else self.line_spacing
            ),
        )


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

    def with_overrides(
        self,
        *,
        line_spacing_body: float | None = None,
        line_spacing_note: float | None = None,
        first_line_indent_cm: float | None = None,
        note_prefixes: tuple[str, ...] | None = None,
    ) -> DocxParagraphStyle:
        """Create a copy with selected paragraph style values overridden.

        Args:
            line_spacing_body (float | None): Body paragraph line spacing.
            line_spacing_note (float | None): Note paragraph line spacing.
            first_line_indent_cm (float | None): First-line indent in centimeters.
            note_prefixes (tuple[str, ...] | None): Prefixes classified as notes.

        Returns:
            DocxParagraphStyle: New paragraph style with overrides applied.
        """

        return DocxParagraphStyle(
            line_spacing_body=(
                line_spacing_body
                if line_spacing_body is not None
                else self.line_spacing_body
            ),
            line_spacing_note=(
                line_spacing_note
                if line_spacing_note is not None
                else self.line_spacing_note
            ),
            first_line_indent_cm=(
                first_line_indent_cm
                if first_line_indent_cm is not None
                else self.first_line_indent_cm
            ),
            note_prefixes=(
                note_prefixes if note_prefixes is not None else self.note_prefixes
            ),
        )


@dataclass(frozen=True, slots=True)
class DocxStyle:
    """Complete style bundle for DOCX writing.

    `docxrender` accepts this structured style object directly. Reading TOML,
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

    def with_overrides(
        self,
        *,
        fonts: DocxFontStyle | None = None,
        sizes: DocxSizeStyle | None = None,
        table: DocxTableStyle | None = None,
        paragraph: DocxParagraphStyle | None = None,
    ) -> DocxStyle:
        """Create a copy with selected style components replaced.

        Args:
            fonts (DocxFontStyle | None): Optional font style replacement.
            sizes (DocxSizeStyle | None): Optional size style replacement.
            table (DocxTableStyle | None): Optional table style replacement.
            paragraph (DocxParagraphStyle | None): Optional paragraph style
                replacement.

        Returns:
            DocxStyle: New complete style object with replacements applied.
        """

        return DocxStyle(
            fonts=fonts or self.fonts,
            sizes=sizes or self.sizes,
            table=table or self.table,
            paragraph=paragraph or self.paragraph,
        )


@dataclass(frozen=True, slots=True)
class DocxFieldMarkerOptions:
    """Options for DOCX field update markers and field freezing.

    These options operate directly on DOCX XML and do not require LibreOffice or
    UNO. Field refresh through LibreOffice is configured separately with
    `DocxFieldRefreshOptions`.

    Attributes:
        should_update_fields (bool): Whether fields should be marked for update.
        should_freeze_fields (bool): Whether field markup should be frozen while
            preserving current field result text.
    """

    should_update_fields: bool = True
    should_freeze_fields: bool = False


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
class DocxHeaderFooterImageOptions:
    """Header and footer image replacement options.

    Attributes:
        file_header_image (Path | None): Image file to place in headers.
        file_footer_image (Path | None): Image file to place in footers.
        width_cm (float): Inserted image width in centimeters when no existing
            image relationship can be replaced.
        should_replace_existing (bool): Whether existing header/footer image media
            should be replaced.
        should_insert_when_missing (bool): Whether images should be inserted when
            the target header/footer does not already contain image media.
        idx_section_start (int): Zero-based section index where missing image
            insertion starts.
    """

    file_header_image: Path | None = None
    file_footer_image: Path | None = None
    width_cm: float = 16.0
    should_replace_existing: bool = True
    should_insert_when_missing: bool = True
    idx_section_start: int = 0


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
        header_footer_images (DocxHeaderFooterImageOptions | None): Optional header
            and footer image replacement settings.
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
    header_footer_images: DocxHeaderFooterImageOptions | None = None


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
        should_update_fields (bool): Whether staged DOCX fields should be marked for
            update before LibreOffice loads the document.
        should_freeze_fields (bool): Whether refreshed field results should be frozen
            in `file_out_docx_refreshed` when that output is requested.
    """

    exe_libreoffice: Path
    file_in_docx: Path
    file_out_pdf: Path
    dir_user_profile: Path
    file_out_docx_refreshed: Path | None = None
    file_listener_log: Path | None = None
    should_update_fields: bool = True
    should_freeze_fields: bool = False


@dataclass(frozen=True, slots=True)
class DocxToPdfResult:
    """Result of converting DOCX to PDF.

    Attributes:
        file_pdf (Path): Written PDF path.
        file_docx_refreshed (Path | None): Refreshed DOCX path when requested.
    """

    file_pdf: Path
    file_docx_refreshed: Path | None = None
