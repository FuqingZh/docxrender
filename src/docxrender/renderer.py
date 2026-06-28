from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Final, Literal

from docxrender.api import convert_docx_to_pdf, write_docx, write_existing_docx
from docxrender.contracts import (
    DocxBodyAnchorOptions,
    DocxBodyRenderPolicy,
    DocxFieldMarkerOptions,
    DocxFieldRefreshOptions,
    DocxFontStyle,
    DocxHeaderFooterImageOptions,
    DocxMarkdownOptions,
    DocxParagraphStyle,
    DocxSizeStyle,
    DocxStyle,
    DocxTableStyle,
    DocxTemplateContextPolicy,
    DocxTemplateImageSpec,
    DocxTemplateRenderOptions,
    DocxTemplateRenderResult,
    DocxToPdfOptions,
    DocxToPdfResult,
    DocxWriteOptions,
    DocxWriteResult,
)
from docxrender.template import write_docx_template


def create_empty_template_context() -> dict[str, Any]:
    return {}


def create_empty_template_inline_images() -> dict[str, DocxTemplateImageSpec]:
    return {}


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


@dataclass(frozen=True, slots=True)
class _TemplateState:
    file_template: Path | None = None
    context: Mapping[str, Any] = field(default_factory=create_empty_template_context)
    inline_images: Mapping[str, DocxTemplateImageSpec] = field(
        default_factory=create_empty_template_inline_images
    )
    context_policy: DocxTemplateContextPolicy = DocxTemplateContextPolicy()


@dataclass(frozen=True, slots=True)
class _RendererState:
    style: DocxStyle
    template: _TemplateState = _TemplateState()
    body_anchor: DocxBodyAnchorOptions = DocxBodyAnchorOptions()
    markdown: DocxMarkdownOptions = DocxMarkdownOptions()
    body_render_policy: DocxBodyRenderPolicy = DocxBodyRenderPolicy()
    field_markers: DocxFieldMarkerOptions = DocxFieldMarkerOptions()
    field_refresh: DocxFieldRefreshOptions | None = None
    header_footer_images: DocxHeaderFooterImageOptions | None = None
    pdf_settings: _PdfConversionSettings | None = None
    file_docx: Path | None = None


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


_UNSET: Final = object()


class DocxRenderer:
    """Primary fluent pipeline facade for DOCX rendering.

    `DocxRenderer` is an immutable configuration object. Its `with_*` methods do
    not perform I/O and return new renderer instances. Terminal methods such as
    `write_docx`, `write_docx_template`, and `write_pdf` materialize runtime
    template objects only when needed.
    """

    def __init__(
        self,
        renderer: DocxRenderer | None = None,
        *,
        file_docx: Path | None = None,
        style: DocxStyle | None = None,
    ) -> None:
        if renderer is None:
            self._state = _RendererState(
                style=style or create_default_docx_style(),
                file_docx=file_docx,
            )
            return

        state = renderer._state
        self._state = replace(
            state,
            style=style or state.style,
            file_docx=file_docx if file_docx is not None else state.file_docx,
        )

    @classmethod
    def _from_state(cls, state: _RendererState) -> DocxRenderer:
        renderer = cls()
        renderer._state = state
        return renderer

    @property
    def style(self) -> DocxStyle:
        return self._state.style

    @property
    def file_docx(self) -> Path | None:
        return self._state.file_docx

    @property
    def docx_options(self) -> DocxWriteOptions | None:
        """Return `None`; immutable renderers do not cache built write options."""

        return None

    @property
    def template_context(self) -> Mapping[str, Any]:
        return self._state.template.context

    @property
    def template_context_policy(self) -> DocxTemplateContextPolicy:
        return self._state.template.context_policy

    @property
    def body_anchor(self) -> DocxBodyAnchorOptions:
        return self._state.body_anchor

    @property
    def field_markers(self) -> DocxFieldMarkerOptions:
        return self._state.field_markers

    @property
    def markdown(self) -> DocxMarkdownOptions:
        return self._state.markdown

    @property
    def body_render_policy(self) -> DocxBodyRenderPolicy:
        return self._state.body_render_policy

    @property
    def pdf_options(self) -> DocxToPdfOptions | None:
        settings = self._state.pdf_settings
        if settings is None:
            return None
        file_in_docx = settings.file_in_docx or self._state.file_docx
        if file_in_docx is None:
            return None
        return self._create_pdf_options(file_in_docx=file_in_docx)

    def with_docx(self, file_docx: Path) -> DocxRenderer:
        return self._replace_state(file_docx=file_docx)

    def with_style(self, style: DocxStyle) -> DocxRenderer:
        return self._replace_state(style=style)

    def with_template(
        self,
        *,
        file_template: Path | None | object = _UNSET,
        context: Mapping[str, Any] | None = None,
        inline_images: Mapping[str, DocxTemplateImageSpec] | None = None,
        rule_conflict: Literal["caller_wins", "defaults_win"] | object = _UNSET,
        required_keys: tuple[str, ...] | object = _UNSET,
    ) -> DocxRenderer:
        """Return a renderer with template-state overrides applied.

        `context` and `inline_images` are incremental state fields. Passing `None`
        keeps the existing value; passing an empty mapping clears the stored value.
        """

        template = self._state.template
        context_policy = template.context_policy
        template_updated = replace(
            template,
            file_template=(
                template.file_template
                if file_template is _UNSET
                else file_template
            ),
            context=(
                template.context
                if context is None
                else dict(context)
            ),
            inline_images=(
                template.inline_images
                if inline_images is None
                else dict(inline_images)
            ),
            context_policy=replace(
                context_policy,
                rule_conflict=(
                    context_policy.rule_conflict
                    if rule_conflict is _UNSET
                    else rule_conflict
                ),
                required_keys=(
                    context_policy.required_keys
                    if required_keys is _UNSET
                    else required_keys
                ),
            ),
        )
        return self._replace_state(template=template_updated)

    def with_fonts(
        self,
        *,
        font_name_latin: str | None = None,
        font_name_body_east_asia: str | None = None,
        font_name_heading_east_asia: str | None = None,
    ) -> DocxRenderer:
        return self._replace_state(
            style=self._state.style.with_overrides(
                fonts=self._state.style.fonts.with_overrides(
                    font_name_latin=font_name_latin,
                    font_name_body_east_asia=font_name_body_east_asia,
                    font_name_heading_east_asia=font_name_heading_east_asia,
                )
            )
        )

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
    ) -> DocxRenderer:
        return self._replace_state(
            style=self._state.style.with_overrides(
                sizes=self._state.style.sizes.with_overrides(
                    pt_title_page_title=pt_title_page_title,
                    pt_title_page_meta=pt_title_page_meta,
                    pt_title_page_compiler=pt_title_page_compiler,
                    pt_body=pt_body,
                    pt_caption=pt_caption,
                    pt_table=pt_table,
                    pt_heading_by_level=pt_heading_by_level,
                )
            )
        )

    def with_table(
        self,
        *,
        border_color: str | None = None,
        stripe_fill_color: str | None = None,
        border_size_main: str | None = None,
        border_size_header: str | None = None,
        line_spacing: float | None = None,
    ) -> DocxRenderer:
        return self._replace_state(
            style=self._state.style.with_overrides(
                table=self._state.style.table.with_overrides(
                    border_color=border_color,
                    stripe_fill_color=stripe_fill_color,
                    border_size_main=border_size_main,
                    border_size_header=border_size_header,
                    line_spacing=line_spacing,
                )
            )
        )

    def with_paragraph(
        self,
        *,
        line_spacing_body: float | None = None,
        line_spacing_note: float | None = None,
        first_line_indent_cm: float | None = None,
        note_prefixes: tuple[str, ...] | None = None,
    ) -> DocxRenderer:
        return self._replace_state(
            style=self._state.style.with_overrides(
                paragraph=self._state.style.paragraph.with_overrides(
                    line_spacing_body=line_spacing_body,
                    line_spacing_note=line_spacing_note,
                    first_line_indent_cm=first_line_indent_cm,
                    note_prefixes=note_prefixes,
                )
            )
        )

    def with_body_anchor(
        self,
        options: DocxBodyAnchorOptions | None = None,
        *,
        anchor_token: str = "__REPORT_BODY_ANCHOR__",
        rule_match: Literal["equals", "contains"] = "equals",
        rule_missing: Literal["append", "raise"] = "append",
        should_preserve_section_properties: bool = True,
    ) -> DocxRenderer:
        body_anchor = options or DocxBodyAnchorOptions(
            anchor_token=anchor_token,
            rule_match=rule_match,
            rule_missing=rule_missing,
            should_preserve_section_properties=should_preserve_section_properties,
        )
        return self._replace_state(body_anchor=body_anchor)

    def with_markdown(
        self,
        options: DocxMarkdownOptions | None = None,
        *,
        should_parse_inline_bold: bool = True,
        should_parse_inline_code: bool = True,
        should_parse_links_as_text: bool = True,
        should_parse_image_width_attr: bool = True,
        default_image_width_pct: float = 90.0,
    ) -> DocxRenderer:
        markdown = options or DocxMarkdownOptions(
            should_parse_inline_bold=should_parse_inline_bold,
            should_parse_inline_code=should_parse_inline_code,
            should_parse_links_as_text=should_parse_links_as_text,
            should_parse_image_width_attr=should_parse_image_width_attr,
            default_image_width_pct=default_image_width_pct,
        )
        return self._replace_state(markdown=markdown)

    def with_body_render_policy(
        self,
        options: DocxBodyRenderPolicy | None = None,
        *,
        should_number_headings: bool = False,
        rule_ordered_list: Literal["word_style", "plain_text"] = "word_style",
        rule_unordered_list: Literal["word_style", "plain_text"] = "word_style",
        should_stripe_table_rows: bool = False,
    ) -> DocxRenderer:
        policy = options or DocxBodyRenderPolicy(
            should_number_headings=should_number_headings,
            rule_ordered_list=rule_ordered_list,
            rule_unordered_list=rule_unordered_list,
            should_stripe_table_rows=should_stripe_table_rows,
        )
        return self._replace_state(body_render_policy=policy)

    def with_field_update_markers(
        self,
        options: DocxFieldMarkerOptions | None = None,
        *,
        should_update_fields: bool = True,
        should_freeze_fields: bool = False,
    ) -> DocxRenderer:
        field_markers = options or DocxFieldMarkerOptions(
            should_update_fields=should_update_fields,
            should_freeze_fields=should_freeze_fields,
        )
        return self._replace_state(field_markers=field_markers)

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
    ) -> DocxRenderer:
        image_options = options or DocxHeaderFooterImageOptions(
            file_header_image=file_header_image,
            file_footer_image=file_footer_image,
            width_cm=width_cm,
            should_replace_existing=should_replace_existing,
            should_insert_when_missing=should_insert_when_missing,
            idx_section_start=idx_section_start,
        )
        return self._replace_state(header_footer_images=image_options)

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
    ) -> DocxRenderer:
        if options is not None:
            return self._replace_state(field_refresh=options)
        if exe_libreoffice is None or dir_user_profile is None:
            raise ValueError(
                "exe_libreoffice and dir_user_profile are required when "
                "DocxFieldRefreshOptions is not provided."
            )
        return self._replace_state(
            field_refresh=DocxFieldRefreshOptions(
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
        )

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
    ) -> DocxRenderer:
        if options is not None:
            return self._replace_state(
                pdf_settings=_PdfConversionSettings(
                    exe_libreoffice=options.exe_libreoffice,
                    dir_user_profile=options.dir_user_profile,
                    file_out_pdf=options.file_out_pdf,
                    file_in_docx=options.file_in_docx,
                    file_out_docx_refreshed=options.file_out_docx_refreshed,
                    file_listener_log=options.file_listener_log,
                    should_update_fields=options.should_update_fields,
                    should_freeze_fields=options.should_freeze_fields,
                ),
                file_docx=options.file_in_docx,
            )
        if exe_libreoffice is None or dir_user_profile is None or file_out_pdf is None:
            raise ValueError(
                "exe_libreoffice, dir_user_profile, and file_out_pdf are required "
                "when DocxToPdfOptions is not provided."
            )
        return self._replace_state(
            pdf_settings=_PdfConversionSettings(
                exe_libreoffice=exe_libreoffice,
                dir_user_profile=dir_user_profile,
                file_out_pdf=file_out_pdf,
                file_in_docx=file_in_docx,
                file_out_docx_refreshed=file_out_docx_refreshed,
                file_listener_log=file_listener_log,
                should_update_fields=should_update_fields,
                should_freeze_fields=should_freeze_fields,
            ),
            file_docx=(
                file_in_docx if file_in_docx is not None else self._state.file_docx
            ),
        )

    def build_style(self) -> DocxStyle:
        return self.style

    def build_options(
        self,
        *,
        file_out_docx: Path,
        markdown_body: str,
        dir_base: Path,
        file_template: Path | None = None,
        context: Mapping[str, Any] | None = None,
        should_update_fields: bool | None = None,
        should_freeze_fields: bool | None = None,
        field_refresh: DocxFieldRefreshOptions | None = None,
        header_footer_images: DocxHeaderFooterImageOptions | None = None,
    ) -> DocxWriteOptions:
        template_state = self._resolve_template_state(
            file_template=file_template,
            context=context,
        )
        if template_state.file_template is None:
            raise ValueError(
                "file_template is required unless it has been configured through "
                "with_template(...)."
            )
        return DocxWriteOptions(
            file_template=template_state.file_template,
            file_out_docx=file_out_docx,
            context=template_state.context,
            markdown_body=markdown_body,
            dir_base=dir_base,
            style=self._state.style,
            body_anchor=self._state.body_anchor,
            markdown=self._state.markdown,
            body_render_policy=self._state.body_render_policy,
            template_context_policy=template_state.context_policy,
            template_inline_images=template_state.inline_images,
            should_update_fields=(
                should_update_fields
                if should_update_fields is not None
                else self._state.field_markers.should_update_fields
            ),
            should_freeze_fields=(
                should_freeze_fields
                if should_freeze_fields is not None
                else self._state.field_markers.should_freeze_fields
            ),
            field_refresh=field_refresh or self._state.field_refresh,
            header_footer_images=(
                header_footer_images or self._state.header_footer_images
            ),
        )

    def write_docx_template(
        self,
        *,
        file_out_docx: Path,
        file_template: Path | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> DocxTemplateRenderResult:
        template_state = self._resolve_template_state(
            file_template=file_template,
            context=context,
        )
        if template_state.file_template is None:
            raise ValueError(
                "file_template is required unless it has been configured through "
                "with_template(...)."
            )
        return write_docx_template(
            DocxTemplateRenderOptions(
                file_template=template_state.file_template,
                file_out_docx=file_out_docx,
                context=template_state.context,
                inline_images=template_state.inline_images,
                context_policy=template_state.context_policy,
            )
        )

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
        has_template_state = (
            file_template is not None
            or context is not None
            or self._state.template.file_template is not None
        )
        has_template_write = has_template_state and (
            markdown_body is not None
            or dir_base is not None
            or file_out_docx is not None
        )
        if has_template_write:
            if file_out_docx is None or markdown_body is None or dir_base is None:
                raise ValueError(
                    "file_out_docx, markdown_body, and dir_base are required for "
                    "template DOCX writing. file_template and context may be "
                    "supplied per call or through with_template(...)."
                )
            return write_docx(
                self.build_options(
                    file_out_docx=file_out_docx,
                    markdown_body=markdown_body,
                    dir_base=dir_base,
                    file_template=file_template,
                    context=context,
                    should_update_fields=should_update_fields,
                    should_freeze_fields=should_freeze_fields,
                    field_refresh=field_refresh,
                    header_footer_images=header_footer_images,
                )
            )

        if self._state.file_docx is None:
            raise ValueError(
                "file_docx is required when write_docx is used without template "
                "inputs."
            )
        return write_existing_docx(
            file_in_docx=self._state.file_docx,
            file_out_docx=file_out_docx,
            should_update_fields=(
                should_update_fields
                if should_update_fields is not None
                else self._state.field_markers.should_update_fields
            ),
            should_freeze_fields=(
                should_freeze_fields
                if should_freeze_fields is not None
                else self._state.field_markers.should_freeze_fields
            ),
            field_refresh=field_refresh or self._state.field_refresh,
            header_footer_images=(
                header_footer_images or self._state.header_footer_images
            ),
        )

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
        file_docx_effective = self._state.file_docx
        if file_docx_effective is None:
            result_docx = self.write_docx(
                file_template=file_template,
                file_out_docx=file_out_docx,
                context=context,
                markdown_body=markdown_body,
                dir_base=dir_base,
                should_update_fields=should_update_fields,
                should_freeze_fields=should_freeze_fields,
            )
            file_docx_effective = result_docx.file_docx
        return convert_docx_to_pdf(
            self._create_pdf_options(
                file_in_docx=file_docx_effective,
                file_out_pdf=file_out_pdf,
                exe_libreoffice=exe_libreoffice,
                dir_user_profile=dir_user_profile,
                file_out_docx_refreshed=file_out_docx_refreshed,
                file_listener_log=file_listener_log,
                should_update_fields=should_update_fields,
                should_freeze_fields=should_freeze_fields,
            )
        )

    def _replace_state(self, **changes: object) -> DocxRenderer:
        return self._from_state(replace(self._state, **changes))

    def _resolve_template_state(
        self,
        *,
        file_template: Path | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> _TemplateState:
        template_state = self._state.template
        return replace(
            template_state,
            file_template=(
                template_state.file_template
                if file_template is None
                else file_template
            ),
            context=(
                template_state.context
                if context is None
                else dict(context)
            ),
        )

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
        settings = self._state.pdf_settings
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
                else self._state.field_markers.should_update_fields
            ),
            should_freeze_fields=(
                should_freeze_fields
                if should_freeze_fields is not None
                else settings.should_freeze_fields
                if settings is not None and settings.should_freeze_fields is not None
                else self._state.field_markers.should_freeze_fields
            ),
        )
