from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, Self

from docxrender.api import (
    convert_docx_to_pdf,
    write_docx,
    write_existing_docx,
)
from docxrender.contracts import (
    DocxBodyAnchorOptions,
    DocxFieldMarkerOptions,
    DocxFieldRefreshOptions,
    DocxFontStyle,
    DocxHeaderFooterImageOptions,
    DocxParagraphStyle,
    DocxSizeStyle,
    DocxStyle,
    DocxTableStyle,
    DocxToPdfOptions,
    DocxToPdfResult,
    DocxWriteOptions,
    DocxWriteResult,
)


@dataclass(frozen=True, slots=True)
class _PdfConversionSettings:
    exe_libreoffice: Path
    dir_user_profile: Path
    file_out_pdf: Path
    file_in_docx: Path | None = None
    file_out_docx_refreshed: Path | None = None
    file_listener_log: Path | None = None
    should_update_fields: bool | None = None
    should_freeze_fields: bool | None = None


def create_default_docx_style() -> DocxStyle:
    """Create the default DOCX style used by the fluent renderer.

    Returns:
        DocxStyle: Complete style object modeled after the shared report style
            defaults.
    """

    return DocxStyle(
        fonts=DocxFontStyle(
            font_name_latin="Times New Roman",
            font_name_body_east_asia="宋体",
            font_name_heading_east_asia="宋体",
        ),
        sizes=DocxSizeStyle(
            pt_title_page_title=36.0,
            pt_title_page_meta=18.0,
            pt_title_page_compiler=15.0,
            pt_body=12.0,
            pt_caption=10.5,
            pt_table=12.0,
            pt_heading_by_level={
                1: 16.0,
                2: 14.0,
                3: 12.0,
                4: 12.0,
                5: 12.0,
                6: 12.0,
            },
        ),
        table=DocxTableStyle(
            border_color="000000",
            stripe_fill_color="D9D9D9",
            border_size_main="12",
            border_size_header="6",
            line_spacing=1.5,
        ),
        paragraph=DocxParagraphStyle(
            line_spacing_body=1.5,
            line_spacing_note=1.2,
            first_line_indent_cm=0.74,
        ),
    )


class DocxRenderer:
    """Primary fluent pipeline facade for DOCX rendering.

    `DocxRenderer` collects style, DOCX post-processing, and PDF conversion
    settings. Its `with_*` methods do not perform I/O; terminal methods such as
    `write_docx` and `write_pdf` execute the configured pipeline by delegating
    to the module-level public functions.
    """

    def __init__(
        self,
        renderer: DocxRenderer | None = None,
        *,
        file_docx: Path | None = None,
        style: DocxStyle | None = None,
    ) -> None:
        """Initialize a renderer or clone an existing renderer.

        Args:
            renderer (DocxRenderer | None): Optional renderer to copy.
            file_docx (Path | None): Optional current DOCX path override.
            style (DocxStyle | None): Optional style override.
        """

        if renderer is None:
            self._style = style or create_default_docx_style()
            self._body_anchor = DocxBodyAnchorOptions()
            self._field_markers = DocxFieldMarkerOptions()
            self._field_refresh: DocxFieldRefreshOptions | None = None
            self._header_footer_images: DocxHeaderFooterImageOptions | None = None
            self._pdf_settings: _PdfConversionSettings | None = None
            self._docx_options: DocxWriteOptions | None = None
            self._file_docx = file_docx
            return

        self._style = style or renderer.style
        self._body_anchor = renderer._body_anchor
        self._field_markers = renderer._field_markers
        self._field_refresh = renderer._field_refresh
        self._header_footer_images = renderer._header_footer_images
        self._pdf_settings = renderer._pdf_settings
        self._docx_options = renderer._docx_options
        self._file_docx = file_docx if file_docx is not None else renderer.file_docx

    @property
    def style(self) -> DocxStyle:
        """Current complete DOCX style.

        Returns:
            DocxStyle: Complete style after applying fluent overrides.
        """

        return self._style

    @property
    def file_docx(self) -> Path | None:
        """Current DOCX path known to this renderer.

        Returns:
            Path | None: Current DOCX path, when one has been supplied or written.
        """

        return self._file_docx

    @property
    def docx_options(self) -> DocxWriteOptions | None:
        """Last built DOCX write options.

        Returns:
            DocxWriteOptions | None: Most recent template-write options, when built.
        """

        return self._docx_options

    @property
    def body_anchor(self) -> DocxBodyAnchorOptions:
        """Current body anchor settings.

        Returns:
            DocxBodyAnchorOptions: Body insertion anchor settings.
        """

        return self._body_anchor

    @property
    def field_markers(self) -> DocxFieldMarkerOptions:
        """Current DOCX field marker and freeze settings.

        Returns:
            DocxFieldMarkerOptions: Field marker settings used by DOCX writes.
        """

        return self._field_markers

    @property
    def pdf_options(self) -> DocxToPdfOptions | None:
        """Current PDF conversion options when enough state is available.

        Returns:
            DocxToPdfOptions | None: PDF options, or `None` when no input DOCX is
            known yet.
        """

        if self._pdf_settings is None:
            return None
        file_in_docx = self._pdf_settings.file_in_docx or self._file_docx
        if file_in_docx is None:
            return None
        return self._create_pdf_options(file_in_docx=file_in_docx)

    def with_docx(self, file_docx: Path) -> Self:
        """Set the current DOCX path.

        Args:
            file_docx (Path): Existing or future DOCX path.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._file_docx = file_docx
        return self

    def with_style(self, style: DocxStyle) -> Self:
        """Replace the renderer style.

        Args:
            style (DocxStyle): Complete style object to use for future writes.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._style = style
        return self

    def with_fonts(
        self,
        *,
        font_name_latin: str | None = None,
        font_name_body_east_asia: str | None = None,
        font_name_heading_east_asia: str | None = None,
    ) -> Self:
        """Override font settings.

        Args:
            font_name_latin (str | None): Latin font name applied to runs.
            font_name_body_east_asia (str | None): East Asian body font name.
            font_name_heading_east_asia (str | None): East Asian heading font name.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._style = self._style.with_overrides(
            fonts=self._style.fonts.with_overrides(
                font_name_latin=font_name_latin,
                font_name_body_east_asia=font_name_body_east_asia,
                font_name_heading_east_asia=font_name_heading_east_asia,
            )
        )
        return self

    def with_sizes(
        self,
        *,
        pt_title_page_title: float | None = None,
        pt_title_page_meta: float | None = None,
        pt_title_page_compiler: float | None = None,
        pt_body: float | None = None,
        pt_caption: float | None = None,
        pt_table: float | None = None,
        pt_heading_by_level: Mapping[int, float] | None = None,
    ) -> Self:
        """Override size settings.

        Args:
            pt_title_page_title (float | None): Title-page report title size.
            pt_title_page_meta (float | None): Title-page metadata size.
            pt_title_page_compiler (float | None): Compiler or organization size.
            pt_body (float | None): Body paragraph text size.
            pt_caption (float | None): Caption and note text size.
            pt_table (float | None): Markdown table text size.
            pt_heading_by_level (Mapping[int, float] | None): Heading sizes by level.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._style = self._style.with_overrides(
            sizes=self._style.sizes.with_overrides(
                pt_title_page_title=pt_title_page_title,
                pt_title_page_meta=pt_title_page_meta,
                pt_title_page_compiler=pt_title_page_compiler,
                pt_body=pt_body,
                pt_caption=pt_caption,
                pt_table=pt_table,
                pt_heading_by_level=pt_heading_by_level,
            ),
        )
        return self

    def with_table(
        self,
        *,
        border_color: str | None = None,
        stripe_fill_color: str | None = None,
        border_size_main: str | None = None,
        border_size_header: str | None = None,
        line_spacing: float | None = None,
    ) -> Self:
        """Override table style settings.

        Args:
            border_color (str | None): WordprocessingML border color.
            stripe_fill_color (str | None): Body-row stripe fill color.
            border_size_main (str | None): Main border size in Word units.
            border_size_header (str | None): Header border size in Word units.
            line_spacing (float | None): Table paragraph line spacing.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._style = self._style.with_overrides(
            table=self._style.table.with_overrides(
                border_color=border_color,
                stripe_fill_color=stripe_fill_color,
                border_size_main=border_size_main,
                border_size_header=border_size_header,
                line_spacing=line_spacing,
            ),
        )
        return self

    def with_paragraph(
        self,
        *,
        line_spacing_body: float | None = None,
        line_spacing_note: float | None = None,
        first_line_indent_cm: float | None = None,
        note_prefixes: tuple[str, ...] | None = None,
    ) -> Self:
        """Override paragraph style settings.

        Args:
            line_spacing_body (float | None): Body paragraph line spacing.
            line_spacing_note (float | None): Note paragraph line spacing.
            first_line_indent_cm (float | None): First-line indent in centimeters.
            note_prefixes (tuple[str, ...] | None): Prefixes classified as notes.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._style = self._style.with_overrides(
            paragraph=self._style.paragraph.with_overrides(
                line_spacing_body=line_spacing_body,
                line_spacing_note=line_spacing_note,
                first_line_indent_cm=first_line_indent_cm,
                note_prefixes=note_prefixes,
            ),
        )
        return self

    def with_body_anchor(
        self,
        options: DocxBodyAnchorOptions | None = None,
        *,
        anchor_token: str = "__REPORT_BODY_ANCHOR__",
        rule_match: Literal["equals", "contains"] = "equals",
        rule_missing: Literal["append", "raise"] = "append",
        should_preserve_section_properties: bool = True,
    ) -> Self:
        """Configure DOCX body insertion anchor behavior.

        Args:
            options (DocxBodyAnchorOptions | None): Complete anchor options.
            anchor_token (str): Paragraph token marking the insertion point.
            rule_match (Literal["equals", "contains"]): Anchor match rule.
            rule_missing (Literal["append", "raise"]): Missing-anchor rule.
            should_preserve_section_properties (bool): Whether anchor paragraphs with
                section properties should be cleared instead of removed.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._body_anchor = options or DocxBodyAnchorOptions(
            anchor_token=anchor_token,
            rule_match=rule_match,
            rule_missing=rule_missing,
            should_preserve_section_properties=should_preserve_section_properties,
        )
        return self

    def with_field_update_markers(
        self,
        options: DocxFieldMarkerOptions | None = None,
        *,
        should_update_fields: bool = True,
        should_freeze_fields: bool = False,
    ) -> Self:
        """Configure DOCX field update marker and freeze behavior.

        This configuration edits DOCX XML directly and does not require
        LibreOffice or UNO. Use `with_field_refresh` for LibreOffice-backed TOC
        and page-number refresh.

        Args:
            options (DocxFieldMarkerOptions | None): Complete field marker options.
            should_update_fields (bool): Whether fields should be marked for update.
            should_freeze_fields (bool): Whether field markup should be frozen.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._field_markers = options or DocxFieldMarkerOptions(
            should_update_fields=should_update_fields,
            should_freeze_fields=should_freeze_fields,
        )
        return self

    def with_header_footer_images(
        self,
        options: DocxHeaderFooterImageOptions | None = None,
        *,
        file_header_image: Path | None = None,
        file_footer_image: Path | None = None,
        width_cm: float = 16.0,
        should_replace_existing: bool = True,
        should_insert_when_missing: bool = True,
        idx_section_start: int = 0,
    ) -> Self:
        """Configure optional header and footer image handling.

        Args:
            options (DocxHeaderFooterImageOptions | None): Complete image options.
            file_header_image (Path | None): Header image file.
            file_footer_image (Path | None): Footer image file.
            width_cm (float): Inserted image width in centimeters.
            should_replace_existing (bool): Whether existing media should be replaced.
            should_insert_when_missing (bool): Whether missing images should be
                inserted.
            idx_section_start (int): Zero-based section index where missing image
                insertion starts.

        Returns:
            Self: This renderer, for method chaining.
        """

        self._header_footer_images = options or DocxHeaderFooterImageOptions(
            file_header_image=file_header_image,
            file_footer_image=file_footer_image,
            width_cm=width_cm,
            should_replace_existing=should_replace_existing,
            should_insert_when_missing=should_insert_when_missing,
            idx_section_start=idx_section_start,
        )
        return self

    def with_field_refresh(
        self,
        options: DocxFieldRefreshOptions | None = None,
        *,
        exe_libreoffice: Path | None = None,
        dir_user_profile: Path | None = None,
        file_out_docx_refreshed: Path | None = None,
        file_listener_log: Path | None = None,
        should_require_toc: bool = False,
        should_freeze_fields: bool = False,
        timeout_seconds: float = 30.0,
        poll_interval_seconds: float = 0.5,
        stable_checks: int = 2,
    ) -> Self:
        """Configure optional LibreOffice UNO field refresh.

        Args:
            options (DocxFieldRefreshOptions | None): Complete refresh options.
            exe_libreoffice (Path | None): LibreOffice executable path.
            dir_user_profile (Path | None): Isolated LibreOffice profile directory.
            file_out_docx_refreshed (Path | None): Optional refreshed DOCX output.
            file_listener_log (Path | None): Optional listener log path.
            should_require_toc (bool): Whether refreshed DOCX must contain TOC results.
            should_freeze_fields (bool): Whether refreshed fields should be frozen.
            timeout_seconds (float): Maximum wait time for refreshed DOCX validation.
            poll_interval_seconds (float): Poll interval for validation.
            stable_checks (int): Consecutive stable file-stat checks required.

        Returns:
            Self: This renderer, for method chaining.

        Raises:
            ValueError: `exe_libreoffice` or `dir_user_profile` is missing when
                `options` is not provided.
        """

        if options is not None:
            self._field_refresh = options
            return self
        if exe_libreoffice is None or dir_user_profile is None:
            raise ValueError(
                "exe_libreoffice and dir_user_profile are required when "
                "DocxFieldRefreshOptions is not provided."
            )
        self._field_refresh = DocxFieldRefreshOptions(
            exe_libreoffice=exe_libreoffice,
            dir_user_profile=dir_user_profile,
            file_out_docx_refreshed=file_out_docx_refreshed,
            file_listener_log=file_listener_log,
            should_require_toc=should_require_toc,
            should_freeze_fields=should_freeze_fields,
            timeout_seconds=timeout_seconds,
            poll_interval_seconds=poll_interval_seconds,
            stable_checks=stable_checks,
        )
        return self

    def with_pdf_conversion(
        self,
        options: DocxToPdfOptions | None = None,
        *,
        exe_libreoffice: Path | None = None,
        dir_user_profile: Path | None = None,
        file_out_pdf: Path | None = None,
        file_in_docx: Path | None = None,
        file_out_docx_refreshed: Path | None = None,
        file_listener_log: Path | None = None,
        should_update_fields: bool | None = None,
        should_freeze_fields: bool | None = None,
    ) -> Self:
        """Configure optional DOCX-to-PDF conversion.

        Args:
            options (DocxToPdfOptions | None): Complete PDF conversion options.
            exe_libreoffice (Path | None): LibreOffice executable path.
            dir_user_profile (Path | None): Isolated LibreOffice profile directory.
            file_out_pdf (Path | None): Output PDF path.
            file_in_docx (Path | None): Optional input DOCX override.
            file_out_docx_refreshed (Path | None): Optional refreshed DOCX output.
            file_listener_log (Path | None): Optional listener log path.
            should_update_fields (bool | None): Optional field-marker override.
            should_freeze_fields (bool | None): Optional field-freeze override.

        Returns:
            Self: This renderer, for method chaining.

        Raises:
            ValueError: Required PDF conversion paths are missing.
        """

        if options is not None:
            self._pdf_settings = _PdfConversionSettings(
                exe_libreoffice=options.exe_libreoffice,
                dir_user_profile=options.dir_user_profile,
                file_out_pdf=options.file_out_pdf,
                file_in_docx=options.file_in_docx,
                file_out_docx_refreshed=options.file_out_docx_refreshed,
                file_listener_log=options.file_listener_log,
                should_update_fields=options.should_update_fields,
                should_freeze_fields=options.should_freeze_fields,
            )
            self._file_docx = options.file_in_docx
            return self
        if exe_libreoffice is None or dir_user_profile is None or file_out_pdf is None:
            raise ValueError(
                "exe_libreoffice, dir_user_profile, and file_out_pdf are required "
                "when DocxToPdfOptions is not provided."
            )
        self._pdf_settings = _PdfConversionSettings(
            exe_libreoffice=exe_libreoffice,
            dir_user_profile=dir_user_profile,
            file_out_pdf=file_out_pdf,
            file_in_docx=file_in_docx,
            file_out_docx_refreshed=file_out_docx_refreshed,
            file_listener_log=file_listener_log,
            should_update_fields=should_update_fields,
            should_freeze_fields=should_freeze_fields,
        )
        if file_in_docx is not None:
            self._file_docx = file_in_docx
        return self

    def build_style(self) -> DocxStyle:
        """Build the current complete DOCX style.

        Returns:
            DocxStyle: Complete style object.
        """

        return self.style

    def build_options(
        self,
        *,
        file_template: Path,
        file_out_docx: Path,
        context: Mapping[str, Any],
        markdown_body: str,
        dir_base: Path,
        should_update_fields: bool | None = None,
        should_freeze_fields: bool | None = None,
        field_refresh: DocxFieldRefreshOptions | None = None,
        header_footer_images: DocxHeaderFooterImageOptions | None = None,
    ) -> DocxWriteOptions:
        """Build `DocxWriteOptions` from fluent settings and write inputs.

        Args:
            file_template (Path): Input DOCX template path.
            file_out_docx (Path): Output DOCX path to write.
            context (Mapping[str, Any]): Template context passed to `docxtpl`.
            markdown_body (str): Markdown body to insert into the DOCX.
            dir_base (Path): Base directory used to resolve relative image paths.
            should_update_fields (bool | None): Optional field-marker override.
            should_freeze_fields (bool | None): Optional field-freeze override.
            field_refresh (DocxFieldRefreshOptions | None): Optional per-call field
                refresh override.
            header_footer_images (DocxHeaderFooterImageOptions | None): Optional
                per-call header/footer image override.

        Returns:
            DocxWriteOptions: Complete options for the core writer.
        """

        options = DocxWriteOptions(
            file_template=file_template,
            file_out_docx=file_out_docx,
            context=context,
            markdown_body=markdown_body,
            dir_base=dir_base,
            style=self.style,
            body_anchor=self._body_anchor,
            should_update_fields=(
                should_update_fields
                if should_update_fields is not None
                else self._field_markers.should_update_fields
            ),
            should_freeze_fields=(
                should_freeze_fields
                if should_freeze_fields is not None
                else self._field_markers.should_freeze_fields
            ),
            field_refresh=field_refresh or self._field_refresh,
            header_footer_images=header_footer_images or self._header_footer_images,
        )
        self._docx_options = options
        return options

    def write_docx(
        self,
        *,
        file_template: Path | None = None,
        file_out_docx: Path | None = None,
        context: Mapping[str, Any] | None = None,
        markdown_body: str | None = None,
        dir_base: Path | None = None,
        should_update_fields: bool | None = None,
        should_freeze_fields: bool | None = None,
        field_refresh: DocxFieldRefreshOptions | None = None,
        header_footer_images: DocxHeaderFooterImageOptions | None = None,
    ) -> DocxWriteResult:
        """Write a DOCX or post-process the current DOCX.

        Args:
            file_template (Path | None): Optional input DOCX template path.
            file_out_docx (Path | None): Optional output DOCX path.
            context (Mapping[str, Any] | None): Template context.
            markdown_body (str | None): Markdown body to insert.
            dir_base (Path | None): Base directory for relative image paths.
            should_update_fields (bool | None): Optional field-marker override.
            should_freeze_fields (bool | None): Optional field-freeze override.
            field_refresh (DocxFieldRefreshOptions | None): Per-call refresh override.
            header_footer_images (DocxHeaderFooterImageOptions | None): Per-call
                header/footer image override.

        Returns:
            DocxWriteResult: Result containing the written or processed DOCX path.

        Raises:
            ValueError: Required template inputs or current DOCX path are missing.
        """

        has_template_inputs = any(
            value is not None
            for value in (
                file_template,
                file_out_docx,
                context,
                markdown_body,
                dir_base,
            )
        )
        if has_template_inputs:
            if (
                file_template is None
                or file_out_docx is None
                or context is None
                or markdown_body is None
                or dir_base is None
            ):
                raise ValueError(
                    "file_template, file_out_docx, context, markdown_body, and "
                    "dir_base are required for template DOCX writing."
                )
            result = write_docx(
                self.build_options(
                    file_template=file_template,
                    file_out_docx=file_out_docx,
                    context=context,
                    markdown_body=markdown_body,
                    dir_base=dir_base,
                    should_update_fields=should_update_fields,
                    should_freeze_fields=should_freeze_fields,
                    field_refresh=field_refresh,
                    header_footer_images=header_footer_images,
                )
            )
            self._file_docx = result.file_docx
            return result

        if self._file_docx is None:
            raise ValueError(
                "file_docx is required when write_docx is used without template inputs."
            )
        result = write_existing_docx(
            file_in_docx=self._file_docx,
            file_out_docx=file_out_docx,
            should_update_fields=(
                should_update_fields
                if should_update_fields is not None
                else self._field_markers.should_update_fields
            ),
            should_freeze_fields=(
                should_freeze_fields
                if should_freeze_fields is not None
                else self._field_markers.should_freeze_fields
            ),
            field_refresh=field_refresh or self._field_refresh,
            header_footer_images=header_footer_images or self._header_footer_images,
        )
        self._file_docx = result.file_docx
        return result

    def write_pdf(
        self,
        *,
        file_template: Path | None = None,
        file_out_docx: Path | None = None,
        context: Mapping[str, Any] | None = None,
        markdown_body: str | None = None,
        dir_base: Path | None = None,
        file_out_pdf: Path | None = None,
        exe_libreoffice: Path | None = None,
        dir_user_profile: Path | None = None,
        file_out_docx_refreshed: Path | None = None,
        file_listener_log: Path | None = None,
        should_update_fields: bool | None = None,
        should_freeze_fields: bool | None = None,
    ) -> DocxToPdfResult:
        """Write DOCX if needed and convert it to PDF.

        Args:
            file_template (Path | None): Optional template path when no DOCX exists.
            file_out_docx (Path | None): Optional DOCX output path.
            context (Mapping[str, Any] | None): Optional template context.
            markdown_body (str | None): Optional markdown body.
            dir_base (Path | None): Optional base directory for image paths.
            file_out_pdf (Path | None): Optional PDF output path override.
            exe_libreoffice (Path | None): Optional LibreOffice executable override.
            dir_user_profile (Path | None): Optional LibreOffice profile override.
            file_out_docx_refreshed (Path | None): Optional refreshed DOCX output.
            file_listener_log (Path | None): Optional listener log path.
            should_update_fields (bool | None): Optional field-marker override.
            should_freeze_fields (bool | None): Optional field-freeze override.

        Returns:
            DocxToPdfResult: PDF conversion result.
        """

        if self._file_docx is None:
            self.write_docx(
                file_template=file_template,
                file_out_docx=file_out_docx,
                context=context,
                markdown_body=markdown_body,
                dir_base=dir_base,
                should_update_fields=should_update_fields,
                should_freeze_fields=should_freeze_fields,
            )
        if self._file_docx is None:
            raise ValueError("file_docx is required before writing PDF.")
        options = self._create_pdf_options(
            file_in_docx=self._file_docx,
            file_out_pdf=file_out_pdf,
            exe_libreoffice=exe_libreoffice,
            dir_user_profile=dir_user_profile,
            file_out_docx_refreshed=file_out_docx_refreshed,
            file_listener_log=file_listener_log,
            should_update_fields=should_update_fields,
            should_freeze_fields=should_freeze_fields,
        )
        result = convert_docx_to_pdf(options)
        if result.file_docx_refreshed is not None:
            self._file_docx = result.file_docx_refreshed
        return result

    def _create_pdf_options(
        self,
        *,
        file_in_docx: Path,
        file_out_pdf: Path | None = None,
        exe_libreoffice: Path | None = None,
        dir_user_profile: Path | None = None,
        file_out_docx_refreshed: Path | None = None,
        file_listener_log: Path | None = None,
        should_update_fields: bool | None = None,
        should_freeze_fields: bool | None = None,
    ) -> DocxToPdfOptions:
        settings = self._pdf_settings
        exe_effective = exe_libreoffice or (
            settings.exe_libreoffice if settings is not None else None
        )
        profile_effective = dir_user_profile or (
            settings.dir_user_profile if settings is not None else None
        )
        pdf_effective = file_out_pdf or (
            settings.file_out_pdf if settings is not None else None
        )
        if exe_effective is None or profile_effective is None or pdf_effective is None:
            raise ValueError(
                "exe_libreoffice, dir_user_profile, and file_out_pdf are required "
                "for DOCX-to-PDF conversion."
            )
        return DocxToPdfOptions(
            exe_libreoffice=exe_effective,
            file_in_docx=file_in_docx,
            file_out_pdf=pdf_effective,
            dir_user_profile=profile_effective,
            file_out_docx_refreshed=(
                file_out_docx_refreshed
                if file_out_docx_refreshed is not None
                else settings.file_out_docx_refreshed
                if settings is not None
                else None
            ),
            file_listener_log=(
                file_listener_log
                if file_listener_log is not None
                else settings.file_listener_log
                if settings is not None
                else None
            ),
            should_update_fields=(
                should_update_fields
                if should_update_fields is not None
                else settings.should_update_fields
                if settings is not None and settings.should_update_fields is not None
                else self._field_markers.should_update_fields
            ),
            should_freeze_fields=(
                should_freeze_fields
                if should_freeze_fields is not None
                else settings.should_freeze_fields
                if settings is not None and settings.should_freeze_fields is not None
                else self._field_markers.should_freeze_fields
            ),
        )
