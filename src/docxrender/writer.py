from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any, Self

from docxrender.contracts import (
    DocxFieldRefreshOptions,
    DocxFontStyle,
    DocxParagraphStyle,
    DocxSizeStyle,
    DocxStyle,
    DocxTableStyle,
    DocxWriteOptions,
    DocxWriteResult,
)


def create_default_docx_style() -> DocxStyle:
    """Create the default DOCX style used by the fluent writer.

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


class DocxWriter:
    """Fluent facade for configuring and writing DOCX files.

    `DocxWriter` is an ergonomic wrapper around `DocxWriteOptions` and the
    module-level `write_docx` function. It does not own a separate rendering
    pipeline.
    """

    def __init__(self, style: DocxStyle | None = None) -> None:
        """Initialize a fluent DOCX writer.

        Args:
            style (DocxStyle | None): Optional starting style. When omitted,
                shared report-style defaults are used.
        """

        self._style = style or create_default_docx_style()
        self._field_refresh: DocxFieldRefreshOptions | None = None

    def with_style(self, style: DocxStyle) -> Self:
        """Replace the writer style.

        Args:
            style (DocxStyle): Complete style object to use for future writes.

        Returns:
            Self: This writer, for method chaining.
        """

        self._style = style
        return self

    @property
    def style(self) -> DocxStyle:
        """Current complete DOCX style.

        Returns:
            DocxStyle: Complete style after applying fluent overrides.
        """

        return self._style

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
            Self: This writer, for method chaining.
        """

        fonts = self._style.fonts
        self._style = DocxStyle(
            fonts=DocxFontStyle(
                font_name_latin=font_name_latin or fonts.font_name_latin,
                font_name_body_east_asia=(
                    font_name_body_east_asia or fonts.font_name_body_east_asia
                ),
                font_name_heading_east_asia=(
                    font_name_heading_east_asia or fonts.font_name_heading_east_asia
                ),
            ),
            sizes=self._style.sizes,
            table=self._style.table,
            paragraph=self._style.paragraph,
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
            Self: This writer, for method chaining.
        """

        self._style = DocxStyle(
            fonts=self._style.fonts,
            sizes=self._style.sizes.with_overrides(
                pt_title_page_title=pt_title_page_title,
                pt_title_page_meta=pt_title_page_meta,
                pt_title_page_compiler=pt_title_page_compiler,
                pt_body=pt_body,
                pt_caption=pt_caption,
                pt_table=pt_table,
                pt_heading_by_level=pt_heading_by_level,
            ),
            table=self._style.table,
            paragraph=self._style.paragraph,
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
            Self: This writer, for method chaining.
        """

        table = self._style.table
        self._style = DocxStyle(
            fonts=self._style.fonts,
            sizes=self._style.sizes,
            table=DocxTableStyle(
                border_color=border_color or table.border_color,
                stripe_fill_color=stripe_fill_color or table.stripe_fill_color,
                border_size_main=border_size_main or table.border_size_main,
                border_size_header=border_size_header or table.border_size_header,
                line_spacing=(
                    line_spacing if line_spacing is not None else table.line_spacing
                ),
            ),
            paragraph=self._style.paragraph,
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
            Self: This writer, for method chaining.
        """

        paragraph = self._style.paragraph
        self._style = DocxStyle(
            fonts=self._style.fonts,
            sizes=self._style.sizes,
            table=self._style.table,
            paragraph=DocxParagraphStyle(
                line_spacing_body=(
                    line_spacing_body
                    if line_spacing_body is not None
                    else paragraph.line_spacing_body
                ),
                line_spacing_note=(
                    line_spacing_note
                    if line_spacing_note is not None
                    else paragraph.line_spacing_note
                ),
                first_line_indent_cm=(
                    first_line_indent_cm
                    if first_line_indent_cm is not None
                    else paragraph.first_line_indent_cm
                ),
                note_prefixes=note_prefixes or paragraph.note_prefixes,
            ),
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
            Self: This writer, for method chaining.

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
        anchor_token: str = "__REPORT_BODY_ANCHOR__",
        should_update_fields: bool = True,
        should_freeze_fields: bool = False,
        field_refresh: DocxFieldRefreshOptions | None = None,
    ) -> DocxWriteOptions:
        """Build `DocxWriteOptions` from fluent settings and write inputs.

        Args:
            file_template (Path): Input DOCX template path.
            file_out_docx (Path): Output DOCX path to write.
            context (Mapping[str, Any]): Template context passed to `docxtpl`.
            markdown_body (str): Markdown body to insert into the DOCX.
            dir_base (Path): Base directory used to resolve relative image paths.
            anchor_token (str): Paragraph text marking markdown insertion point.
            should_update_fields (bool): Whether fields should be marked for update.
            should_freeze_fields (bool): Whether fields should be frozen after writing.
            field_refresh (DocxFieldRefreshOptions | None): Optional per-call field
                refresh override.

        Returns:
            DocxWriteOptions: Complete options for the core writer.
        """

        return DocxWriteOptions(
            file_template=file_template,
            file_out_docx=file_out_docx,
            context=context,
            markdown_body=markdown_body,
            dir_base=dir_base,
            style=self.style,
            anchor_token=anchor_token,
            should_update_fields=should_update_fields,
            should_freeze_fields=should_freeze_fields,
            field_refresh=field_refresh or self._field_refresh,
        )

    def write_docx(
        self,
        *,
        file_template: Path,
        file_out_docx: Path,
        context: Mapping[str, Any],
        markdown_body: str,
        dir_base: Path,
        anchor_token: str = "__REPORT_BODY_ANCHOR__",
        should_update_fields: bool = True,
        should_freeze_fields: bool = False,
        field_refresh: DocxFieldRefreshOptions | None = None,
    ) -> DocxWriteResult:
        """Write a DOCX file using the fluent writer settings.

        Args:
            file_template (Path): Input DOCX template path.
            file_out_docx (Path): Output DOCX path to write.
            context (Mapping[str, Any]): Template context passed to `docxtpl`.
            markdown_body (str): Markdown body to insert into the DOCX.
            dir_base (Path): Base directory used to resolve relative image paths.
            anchor_token (str): Paragraph text marking markdown insertion point.
            should_update_fields (bool): Whether fields should be marked for update.
            should_freeze_fields (bool): Whether fields should be frozen after writing.
            field_refresh (DocxFieldRefreshOptions | None): Optional per-call field
                refresh override.

        Returns:
            DocxWriteResult: Result containing the written DOCX path.
        """

        from docxrender.api import write_docx

        return write_docx(
            self.build_options(
                file_template=file_template,
                file_out_docx=file_out_docx,
                context=context,
                markdown_body=markdown_body,
                dir_base=dir_base,
                anchor_token=anchor_token,
                should_update_fields=should_update_fields,
                should_freeze_fields=should_freeze_fields,
                field_refresh=field_refresh,
            )
        )
