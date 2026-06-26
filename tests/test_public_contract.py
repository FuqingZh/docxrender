from __future__ import annotations

import base64
import sys
import tempfile
import types
import zipfile
from pathlib import Path
from typing import Any, cast
from unittest import mock

import pytest
from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docxrender import (
    DocxFieldRefreshOptions,
    DocxFontStyle,
    DocxHeaderFooterImageOptions,
    DocxParagraphStyle,
    DocxRenderer,
    DocxSizeStyle,
    DocxStyle,
    DocxTableStyle,
    DocxToPdfOptions,
    DocxToPdfResult,
    DocxWriteOptions,
    DocxWriteResult,
    convert_docx_to_pdf,
    write_docx,
)
from docxrender.docx.fields import (
    write_docx_field_update_markers,
    write_frozen_docx_fields,
)
from docxrender.pdf_uno import (
    create_libreoffice_listener_command,
    create_load_failure_fields,
    import_uno_module,
    start_libreoffice_listener,
    wait_for_listener,
)


class FakeProcess:
    def __init__(self) -> None:
        self.terminated = False

    def poll(self) -> int:
        return 0

    def terminate(self) -> None:
        self.terminated = True

    def wait(self, timeout: float | None = None) -> int:
        return 0

    def kill(self) -> None:
        self.terminated = True


class FakeDocument:
    def __init__(self) -> None:
        self.closed = False
        self.stored = False
        self.store_url_calls: list[tuple[str, tuple[Any, ...]]] = []

    def store(self) -> None:
        self.stored = True

    def storeToURL(self, url: str, properties: tuple[Any, ...]) -> None:
        self.store_url_calls.append((url, properties))

    def close(self, deliver_ownership: bool) -> None:
        self.closed = True


def create_fake_file_url(path: str) -> str:
    return f"file://{path}"


def create_fake_property(name: str, value: object) -> tuple[str, object]:
    return name, value


def create_docx_style() -> DocxStyle:
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
            pt_heading_by_level={1: 16.0, 2: 14.0, 3: 12.0},
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


class TestPublicContract:
    def test_public_imports_are_explicit(self) -> None:
        import docxrender

        assert docxrender.__all__ == [
            "DocxRenderer",
            "DocxFieldRefreshOptions",
            "DocxFontStyle",
            "DocxHeaderFooterImageOptions",
            "DocxParagraphStyle",
            "DocxSizeStyle",
            "DocxStyle",
            "DocxTableStyle",
            "DocxToPdfOptions",
            "DocxToPdfResult",
            "DocxWriteOptions",
            "DocxWriteResult",
            "convert_docx_to_pdf",
            "write_docx",
        ]

    def test_docx_write_options_construct_from_structured_inputs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            options = DocxWriteOptions(
                file_template=path_tmp / "template.docx",
                file_out_docx=path_tmp / "report.docx",
                context={"report_title": "Example"},
                markdown_body="# Heading\n\nBody.",
                dir_base=path_tmp,
                style=create_docx_style(),
            )

            assert options.anchor_token == "__REPORT_BODY_ANCHOR__"
            assert options.should_update_fields is True
            assert options.should_freeze_fields is False
            assert options.field_refresh is None
            assert options.header_footer_images is None
            assert options.style.paragraph.first_line_indent_cm == 0.74

    def test_docx_to_pdf_options_construct_from_conversion_inputs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            options = DocxToPdfOptions(
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                file_in_docx=path_tmp / "report.docx",
                file_out_pdf=path_tmp / "report.pdf",
                dir_user_profile=path_tmp / "lo-profile",
            )

            assert options.file_out_docx_refreshed is None
            assert options.file_listener_log is None
            assert options.should_update_fields is True
            assert options.should_freeze_fields is False

    def test_public_results_are_structured_paths(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            result_docx = DocxWriteResult(file_docx=path_tmp / "report.docx")
            result_pdf = DocxToPdfResult(file_pdf=path_tmp / "report.pdf")

            assert result_docx.file_docx.name == "report.docx"
            assert result_pdf.file_pdf.name == "report.pdf"
            assert result_pdf.file_docx_refreshed is None

    def test_docx_size_style_with_overrides_changes_selected_values(self) -> None:
        sizes = create_docx_style().sizes

        updated = sizes.with_overrides(pt_body=11.0)

        assert updated.pt_body == 11.0
        assert updated.pt_caption == sizes.pt_caption
        assert updated.pt_heading_by_level == sizes.pt_heading_by_level

    def test_docx_size_style_with_overrides_copies_heading_sizes(self) -> None:
        sizes = create_docx_style().sizes
        heading_sizes = {1: 18.0, 2: 15.0}

        updated = sizes.with_overrides(pt_heading_by_level=heading_sizes)
        heading_sizes[1] = 99.0

        assert updated.pt_heading_by_level == {1: 18.0, 2: 15.0}
        assert sizes.pt_heading_by_level[1] == 16.0

    def test_docx_renderer_style_returns_default_style(self) -> None:
        style = DocxRenderer().style

        assert style.fonts.font_name_latin == "Times New Roman"
        assert style.fonts.font_name_body_east_asia == "宋体"
        assert style.sizes.pt_body == 12.0
        assert style.sizes.pt_heading_by_level[1] == 16.0
        assert style.sizes.pt_heading_by_level[6] == 12.0
        assert style.table.border_color == "000000"
        assert style.paragraph.first_line_indent_cm == 0.74

    def test_docx_renderer_fluent_overrides_are_partial(self) -> None:
        style = (
            DocxRenderer()
            .with_fonts(font_name_body_east_asia="黑体")
            .with_sizes(pt_body=11.0)
            .with_table(stripe_fill_color="FFFFFF")
            .with_paragraph(note_prefixes=("Note:",))
            .style
        )

        assert style.fonts.font_name_latin == "Times New Roman"
        assert style.fonts.font_name_body_east_asia == "黑体"
        assert style.sizes.pt_body == 11.0
        assert style.sizes.pt_caption == 10.5
        assert style.table.border_color == "000000"
        assert style.table.stripe_fill_color == "FFFFFF"
        assert style.paragraph.note_prefixes == ("Note:",)
        assert style.paragraph.line_spacing_body == 1.5

    def test_docx_renderer_build_style_matches_style_property(self) -> None:
        renderer = DocxRenderer().with_sizes(pt_body=11.0)

        assert renderer.build_style() is renderer.style

    def test_docx_renderer_build_options_uses_fluent_state(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            field_refresh = DocxFieldRefreshOptions(
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                dir_user_profile=path_tmp / "lo-profile",
            )
            header_footer = DocxHeaderFooterImageOptions(
                file_header_image=path_tmp / "header.png",
                idx_section_start=1,
            )

            options = (
                DocxRenderer()
                .with_sizes(pt_body=11.0)
                .with_field_refresh(field_refresh)
                .with_header_footer_images(header_footer)
                .build_options(
                    file_template=path_tmp / "template.docx",
                    file_out_docx=path_tmp / "report.docx",
                    context={"report_title": "Builder"},
                    markdown_body="Body.",
                    dir_base=path_tmp,
                )
            )

            assert options.style.sizes.pt_body == 11.0
            assert options.field_refresh is field_refresh
            assert options.header_footer_images is header_footer
            header_footer_options = options.header_footer_images
            assert header_footer_options is not None
            assert header_footer_options.idx_section_start == 1
            assert options.anchor_token == "__REPORT_BODY_ANCHOR__"
            renderer = DocxRenderer().with_field_refresh(field_refresh)
            built = renderer.build_options(
                file_template=path_tmp / "template.docx",
                file_out_docx=path_tmp / "report.docx",
                context={},
                markdown_body="Body.",
                dir_base=path_tmp,
            )
            assert renderer.docx_options is built

    def test_docx_renderer_header_footer_images_can_be_built_from_keywords(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            options = (
                DocxRenderer()
                .with_header_footer_images(
                    file_header_image=path_tmp / "header.png",
                    idx_section_start=1,
                )
                .build_options(
                    file_template=path_tmp / "template.docx",
                    file_out_docx=path_tmp / "report.docx",
                    context={},
                    markdown_body="Body.",
                    dir_base=path_tmp,
                )
            )

            assert options.header_footer_images is not None
            assert options.header_footer_images.idx_section_start == 1

    def test_docx_renderer_field_refresh_can_be_built_from_keywords(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)

            options = (
                DocxRenderer()
                .with_field_refresh(
                    exe_libreoffice=Path("/usr/bin/libreoffice"),
                    dir_user_profile=path_tmp / "lo-profile",
                    should_require_toc=True,
                )
                .build_options(
                    file_template=path_tmp / "template.docx",
                    file_out_docx=path_tmp / "report.docx",
                    context={},
                    markdown_body="Body.",
                    dir_base=path_tmp,
                )
            )

            assert options.field_refresh is not None
            field_refresh = options.field_refresh
            assert field_refresh is not None
            assert field_refresh.exe_libreoffice == Path("/usr/bin/libreoffice")
            assert field_refresh.should_require_toc is True

    def test_docx_renderer_requires_field_refresh_runtime_paths(self) -> None:
        with pytest.raises(ValueError, match="exe_libreoffice"):
            DocxRenderer().with_field_refresh()

    def test_docx_writer_is_not_public(self) -> None:
        import docxrender

        assert not hasattr(docxrender, "DocxWriter")

    def test_docx_renderer_write_docx_uses_core_writer(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_out_docx = path_tmp / "report.docx"
            _write_template(file_template)

            result = (
                DocxRenderer()
                .with_sizes(pt_body=11.0)
                .write_docx(
                    file_template=file_template,
                    file_out_docx=file_out_docx,
                    context={"report_title": "Builder Report"},
                    markdown_body="Body.",
                    dir_base=path_tmp,
                )
            )

            document = Document(str(result.file_docx))
            paragraph_by_text = {
                paragraph.text: paragraph for paragraph in document.paragraphs
            }
            assert result.file_docx == file_out_docx
            assert (
                _run_font_size_pt(_first_text_run(paragraph_by_text["Body."])) == 11.0
            )

    def test_docx_renderer_clone_copies_state_without_linking(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            base = (
                DocxRenderer(file_docx=path_tmp / "report.docx")
                .with_sizes(pt_body=11.0)
                .with_pdf_conversion(
                    exe_libreoffice=Path("/usr/bin/libreoffice"),
                    dir_user_profile=path_tmp / "lo-profile",
                    file_out_pdf=path_tmp / "report.pdf",
                )
            )
            clone = DocxRenderer(base, file_docx=path_tmp / "edited.docx").with_sizes(
                pt_body=10.0
            )

            assert base.file_docx == path_tmp / "report.docx"
            assert clone.file_docx == path_tmp / "edited.docx"
            assert base.style.sizes.pt_body == 11.0
            assert clone.style.sizes.pt_body == 10.0
            assert clone.pdf_options is not None
            assert clone.pdf_options.file_in_docx == path_tmp / "edited.docx"

    def test_docx_renderer_can_postprocess_existing_docx(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_docx = path_tmp / "report.docx"
            _write_template(file_docx)

            result = DocxRenderer(file_docx=file_docx).write_docx()

            assert result.file_docx == file_docx
            with zipfile.ZipFile(file_docx) as zip_file:
                settings = zip_file.read("word/settings.xml").decode("utf-8")
            assert "<w:updateFields" in settings

    def test_docx_renderer_write_pdf_uses_current_docx(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_docx = path_tmp / "report.docx"
            file_pdf = path_tmp / "report.pdf"
            _write_template(file_docx)
            captured_options: list[DocxToPdfOptions] = []

            def fake_convert(options: DocxToPdfOptions) -> DocxToPdfResult:
                captured_options.append(options)
                return DocxToPdfResult(file_pdf=options.file_out_pdf)

            with mock.patch(
                "docxrender.renderer.convert_docx_to_pdf",
                side_effect=fake_convert,
            ):
                result = (
                    DocxRenderer(file_docx=file_docx)
                    .with_pdf_conversion(
                        exe_libreoffice=Path("/usr/bin/libreoffice"),
                        dir_user_profile=path_tmp / "lo-profile",
                        file_out_pdf=file_pdf,
                    )
                    .write_pdf()
                )

            assert result.file_pdf == file_pdf
            assert captured_options[0].file_in_docx == file_docx

    def test_docx_renderer_does_not_expose_convert_docx_to_pdf(self) -> None:
        assert not hasattr(DocxRenderer(), "convert_docx_to_pdf")

    def test_docx_renderer_write_pdf_writes_docx_before_pdf(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_docx = path_tmp / "report.docx"
            file_pdf = path_tmp / "report.pdf"
            _write_template(file_template)
            captured_options: list[DocxToPdfOptions] = []

            def fake_convert(options: DocxToPdfOptions) -> DocxToPdfResult:
                captured_options.append(options)
                return DocxToPdfResult(file_pdf=options.file_out_pdf)

            with mock.patch(
                "docxrender.renderer.convert_docx_to_pdf",
                side_effect=fake_convert,
            ):
                renderer = DocxRenderer().with_pdf_conversion(
                    exe_libreoffice=Path("/usr/bin/libreoffice"),
                    dir_user_profile=path_tmp / "lo-profile",
                    file_out_pdf=file_pdf,
                )
                result = renderer.write_pdf(
                    file_template=file_template,
                    file_out_docx=file_docx,
                    context={"report_title": "PDF"},
                    markdown_body="Body.",
                    dir_base=path_tmp,
                )

            assert result.file_pdf == file_pdf
            assert renderer.file_docx == file_docx
            assert captured_options[0].file_in_docx == file_docx

    def test_write_docx_creates_minimal_document(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_image = path_tmp / "image.png"
            file_out_docx = path_tmp / "report.docx"
            _write_template(file_template)
            _write_png(file_image)

            options_docx = DocxWriteOptions(
                file_template=file_template,
                file_out_docx=file_out_docx,
                context={"report_title": "Contract Report"},
                markdown_body=(
                    "# Heading\n\n"
                    "Body first line.  \n"
                    "Body second line.\n\n"
                    "注：Note text.\n\n"
                    "1. First item\n"
                    "2. Second item\n\n"
                    "| A | B |\n"
                    "| --- | --- |\n"
                    "| 1 | 2 |\n\n"
                    "![Example image](image.png){width=20%}\n"
                ),
                dir_base=path_tmp,
                style=create_docx_style(),
            )

            result = write_docx(options_docx)
            document = Document(str(result.file_docx))
            texts = [paragraph.text for paragraph in document.paragraphs]

            assert result.file_docx == file_out_docx
            assert "Contract Report" in texts
            assert "Heading" in texts
            assert "Body first line.\nBody second line." in texts
            assert "注：Note text." in texts
            assert "First item" in texts
            assert "Second item" in texts
            assert "Example image" in texts
            assert document.tables[0].cell(0, 0).text == "A"
            assert document.tables[0].cell(1, 1).text == "2"
            assert len(document.inline_shapes) == 1
            paragraph_by_text = {
                paragraph.text: paragraph for paragraph in document.paragraphs
            }
            assert (
                _run_font_size_pt(_first_text_run(paragraph_by_text["Heading"]))
                == 16.0
            )
            assert (
                _run_font_size_pt(
                    _first_text_run(
                        paragraph_by_text["Body first line.\nBody second line."]
                    )
                )
                == 12.0
            )
            assert (
                _run_font_size_pt(_first_text_run(paragraph_by_text["注：Note text."]))
                == 10.5
            )
            assert (
                _run_font_size_pt(_first_text_run(paragraph_by_text["Example image"]))
                == 10.5
            )
            assert 'w:val="single"' in cast(Any, document.tables[0].cell(0, 0))._tc.xml

    def test_write_docx_without_field_refresh_does_not_import_uno(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_out_docx = path_tmp / "report.docx"
            _write_template(file_template)

            def fail_on_uno_import(
                name: str,
                package: str | None = None,
            ) -> Any:
                if name == "uno":
                    raise AssertionError("Base write_docx must not import uno.")
                return import_module_original(name, package=package)

            import importlib

            import_module_original = importlib.import_module
            with mock.patch("importlib.import_module", side_effect=fail_on_uno_import):
                result = write_docx(
                    DocxWriteOptions(
                        file_template=file_template,
                        file_out_docx=file_out_docx,
                        context={"report_title": "No UNO"},
                        markdown_body="Body.",
                        dir_base=path_tmp,
                        style=create_docx_style(),
                    )
                )

            assert result.file_docx == file_out_docx
            assert file_out_docx.exists()

    def test_write_docx_field_refresh_overwrites_output_docx(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_out_docx = path_tmp / "report.docx"
            _write_template(file_template)

            def fake_refresh_docx_with_uno(**kwargs: object) -> None:
                file_out = cast(Path, kwargs["file_out_docx"])
                file_out.write_bytes(b"refreshed")

            from docxrender import pdf_uno

            with mock.patch.object(
                pdf_uno,
                "refresh_docx_with_uno",
                side_effect=fake_refresh_docx_with_uno,
            ):
                result = write_docx(
                    DocxWriteOptions(
                        file_template=file_template,
                        file_out_docx=file_out_docx,
                        context={"report_title": "Refresh"},
                        markdown_body="Body.",
                        dir_base=path_tmp,
                        style=create_docx_style(),
                        field_refresh=DocxFieldRefreshOptions(
                            exe_libreoffice=Path("/usr/bin/libreoffice"),
                            dir_user_profile=path_tmp / "lo-profile",
                            poll_interval_seconds=0.0,
                            stable_checks=1,
                        ),
                    )
                )

            assert result.file_docx == file_out_docx
            assert file_out_docx.read_bytes() == b"refreshed"

    def test_write_docx_field_refresh_can_write_separate_output_docx(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_out_docx = path_tmp / "report.docx"
            file_refreshed = path_tmp / "report-refreshed.docx"
            _write_template(file_template)

            def fake_refresh_docx_with_uno(**kwargs: object) -> None:
                file_in = cast(Path, kwargs["file_in_docx"])
                file_out = cast(Path, kwargs["file_out_docx"])
                assert file_in == file_out_docx
                file_out.write_bytes(b"refreshed separate")

            from docxrender import pdf_uno

            with mock.patch.object(
                pdf_uno,
                "refresh_docx_with_uno",
                side_effect=fake_refresh_docx_with_uno,
            ):
                result = write_docx(
                    DocxWriteOptions(
                        file_template=file_template,
                        file_out_docx=file_out_docx,
                        context={"report_title": "Refresh"},
                        markdown_body="Body.",
                        dir_base=path_tmp,
                        style=create_docx_style(),
                        field_refresh=DocxFieldRefreshOptions(
                            exe_libreoffice=Path("/usr/bin/libreoffice"),
                            dir_user_profile=path_tmp / "lo-profile",
                            file_out_docx_refreshed=file_refreshed,
                            poll_interval_seconds=0.0,
                            stable_checks=1,
                        ),
                    )
                )

            assert result.file_docx == file_out_docx
            assert file_out_docx.exists()
            assert file_refreshed.read_bytes() == b"refreshed separate"

    def test_write_docx_field_refresh_can_require_toc_and_freeze(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_out_docx = path_tmp / "report.docx"
            _write_template(file_template)

            def fake_refresh_docx_with_uno(**kwargs: object) -> None:
                _write_minimal_field_docx(cast(Path, kwargs["file_out_docx"]))

            from docxrender import pdf_uno

            with mock.patch.object(
                pdf_uno,
                "refresh_docx_with_uno",
                side_effect=fake_refresh_docx_with_uno,
            ):
                write_docx(
                    DocxWriteOptions(
                        file_template=file_template,
                        file_out_docx=file_out_docx,
                        context={"report_title": "Refresh"},
                        markdown_body="Body.",
                        dir_base=path_tmp,
                        style=create_docx_style(),
                        field_refresh=DocxFieldRefreshOptions(
                            exe_libreoffice=Path("/usr/bin/libreoffice"),
                            dir_user_profile=path_tmp / "lo-profile",
                            should_require_toc=True,
                            should_freeze_fields=True,
                            poll_interval_seconds=0.0,
                            stable_checks=1,
                        ),
                    )
                )

            text_document = _read_docx_part(file_out_docx, "word/document.xml")
            assert "Rendered TOC" in text_document
            assert "fldChar" not in text_document

    def test_write_docx_field_refresh_reports_missing_required_toc(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_out_docx = path_tmp / "report.docx"
            _write_template(file_template)

            def fake_refresh_docx_with_uno(**kwargs: object) -> None:
                _write_minimal_field_docx_without_result(
                    cast(Path, kwargs["file_out_docx"])
                )

            from docxrender import pdf_uno

            with mock.patch.object(
                pdf_uno,
                "refresh_docx_with_uno",
                side_effect=fake_refresh_docx_with_uno,
            ):
                with pytest.raises(RuntimeError, match="TOC result"):
                    write_docx(
                        DocxWriteOptions(
                            file_template=file_template,
                            file_out_docx=file_out_docx,
                            context={"report_title": "Refresh"},
                            markdown_body="Body.",
                            dir_base=path_tmp,
                            style=create_docx_style(),
                            field_refresh=DocxFieldRefreshOptions(
                                exe_libreoffice=Path("/usr/bin/libreoffice"),
                                dir_user_profile=path_tmp / "lo-profile",
                                should_require_toc=True,
                                poll_interval_seconds=0.0,
                                stable_checks=1,
                            ),
                        )
                    )

    def test_create_libreoffice_listener_command_uses_isolated_profile(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_pdf_") as dir_tmp:
            command = create_libreoffice_listener_command(
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                dir_user_profile=Path(dir_tmp) / "profile",
                port=23001,
            )

            assert command[0] == "/usr/bin/libreoffice"
            assert "--headless" in command
            assert "--accept=socket,host=127.0.0.1,port=23001;urp;" in command
            assert command[-1].startswith("-env:UserInstallation=file://")

    def test_convert_docx_to_pdf_stages_input_before_load(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_in_docx = path_tmp / "report.docx"
            file_in_docx.write_bytes(b"docx payload")
            file_source_lock = path_tmp / ".~lock.report.docx#"
            file_source_lock.write_text("lock", encoding="utf-8")
            file_out_pdf = path_tmp / "report.pdf"
            file_out_docx_refreshed = path_tmp / "report-refreshed.docx"
            file_listener_log = path_tmp / "listener.log"
            options_pdf = DocxToPdfOptions(
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                file_in_docx=file_in_docx,
                file_out_pdf=file_out_pdf,
                dir_user_profile=path_tmp / "lo-profile",
                file_out_docx_refreshed=file_out_docx_refreshed,
                file_listener_log=file_listener_log,
                should_update_fields=False,
            )
            fake_process = FakeProcess()
            fake_doc = FakeDocument()
            fake_uno = types.SimpleNamespace(
                systemPathToFileUrl=create_fake_file_url
            )
            captured: dict[str, object] = {}

            def fake_load_document_or_raise(**kwargs: object) -> FakeDocument:
                captured.update(kwargs)
                file_staged = cast(Path, kwargs["file_in_docx_staged"])
                assert file_staged != file_in_docx
                assert file_staged.exists()
                assert file_staged.read_bytes() == file_in_docx.read_bytes()
                file_staged.write_bytes(b"refreshed staged payload")
                return fake_doc

            from docxrender import pdf_uno

            with (
                mock.patch.object(pdf_uno, "select_free_port", return_value=23001),
                mock.patch.object(
                    pdf_uno,
                    "create_libreoffice_listener_command",
                    return_value=["/usr/bin/libreoffice"],
                ),
                mock.patch.object(pdf_uno, "validate_libreoffice_executable"),
                mock.patch.object(
                    pdf_uno.subprocess,
                    "Popen",
                    return_value=fake_process,
                ),
                mock.patch.object(pdf_uno, "wait_for_listener"),
                mock.patch.object(pdf_uno, "import_uno_module", return_value=fake_uno),
                mock.patch.object(pdf_uno, "connect_desktop", return_value=object()),
                mock.patch.object(
                    pdf_uno,
                    "load_uno_document_or_raise",
                    side_effect=fake_load_document_or_raise,
                ),
                mock.patch.object(pdf_uno, "refresh_uno_document_fields"),
                mock.patch.object(
                    pdf_uno,
                    "create_property",
                    side_effect=create_fake_property,
                ),
            ):
                result = convert_docx_to_pdf(options_pdf)

            assert result.file_pdf == file_out_pdf
            assert result.file_docx_refreshed == file_out_docx_refreshed
            assert captured["file_in_docx_source"] == file_in_docx
            assert captured["file_source_lock"] == file_source_lock
            assert captured["file_listener_log"] == file_listener_log.resolve()
            assert fake_doc.stored
            assert fake_doc.closed
            assert len(fake_doc.store_url_calls) == 1
            assert file_out_docx_refreshed.read_bytes() == b"refreshed staged payload"
            assert fake_process.terminated

    def test_convert_docx_to_pdf_preserves_load_failure(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_in_docx = path_tmp / "report.docx"
            file_in_docx.write_bytes(b"docx payload")
            options_pdf = DocxToPdfOptions(
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                file_in_docx=file_in_docx,
                file_out_pdf=path_tmp / "report.pdf",
                dir_user_profile=path_tmp / "lo-profile",
                should_update_fields=False,
            )
            fake_process = FakeProcess()
            fake_uno = types.SimpleNamespace(
                systemPathToFileUrl=create_fake_file_url
            )

            from docxrender import pdf_uno

            with (
                mock.patch.object(pdf_uno, "select_free_port", return_value=23001),
                mock.patch.object(
                    pdf_uno,
                    "create_libreoffice_listener_command",
                    return_value=["/usr/bin/libreoffice"],
                ),
                mock.patch.object(pdf_uno, "validate_libreoffice_executable"),
                mock.patch.object(
                    pdf_uno.subprocess,
                    "Popen",
                    return_value=fake_process,
                ),
                mock.patch.object(pdf_uno, "wait_for_listener"),
                mock.patch.object(pdf_uno, "import_uno_module", return_value=fake_uno),
                mock.patch.object(pdf_uno, "connect_desktop", return_value=object()),
                mock.patch.object(
                    pdf_uno,
                    "load_uno_document_or_raise",
                    side_effect=RuntimeError("load failed"),
                ),
            ):
                with pytest.raises(RuntimeError, match="load failed"):
                    convert_docx_to_pdf(options_pdf)

            assert fake_process.terminated

    def test_pdf_load_failure_fields_include_staged_and_lock_info(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_pdf_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_source = path_tmp / "source.docx"
            file_source.write_bytes(b"source")
            file_staged = path_tmp / "staged.docx"
            file_staged.write_bytes(b"staged")
            file_listener_log = path_tmp / "listener.log"
            file_listener_log.write_text("listener tail", encoding="utf-8")
            file_source_lock = path_tmp / ".~lock.source.docx#"
            file_source_lock.write_text("lock", encoding="utf-8")

            fields = create_load_failure_fields(
                file_in_docx_source=file_source,
                file_in_docx_staged=file_staged,
                file_url="file:///tmp/staged.docx",
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                dir_user_profile=path_tmp / "profile",
                process_listener=FakeProcess(),
                file_listener_log=file_listener_log,
                file_source_lock=file_source_lock,
                probe_ok=True,
                load_default_ok=False,
                load_hidden_only_ok=False,
            )
            text_error = "\n".join(fields)

            assert "error_code=libreoffice_uno_load_failed" in text_error
            assert "reason_code=staged_docx_import_failed" in text_error
            assert f"file_in_docx={file_source.resolve()}" in text_error
            assert f"file_in_docx_staged={file_staged.resolve()}" in text_error
            assert f"source_lock_file={file_source_lock.resolve()}" in text_error
            assert "probe_swriter_factory=ok" in text_error
            assert "load_staged_default_props=failed" in text_error
            assert "load_staged_hidden_only=failed" in text_error
            assert "listener_log_tail=listener tail" in text_error
            assert (
                "validate_libreoffice=libreoffice --headless --version" in text_error
            )
            assert (
                "install_debian_ubuntu=sudo apt install libreoffice python3-uno"
                in text_error
            )

    def test_pdf_uno_runtime_is_loaded_only_when_requested(self) -> None:
        with mock.patch("docxrender.pdf_uno.importlib.import_module") as import_module:
            import_module.side_effect = ImportError("missing uno")
            with pytest.raises(RuntimeError, match="libreoffice_uno_import_failed"):
                import_uno_module()
            import_module.assert_called_once_with("uno")

    def test_pdf_uno_import_failure_includes_install_guidance(self) -> None:
        with mock.patch("docxrender.pdf_uno.importlib.import_module") as import_module:
            import_module.side_effect = ImportError("missing uno")

            with pytest.raises(RuntimeError) as ctx:
                import_uno_module()

        text_error = str(ctx.value)
        assert "error_code=libreoffice_uno_import_failed" in text_error
        assert "validate_libreoffice=libreoffice --headless --version" in text_error
        assert 'validate_uno=python -c "import uno"' in text_error
        assert (
            "install_debian_ubuntu=sudo apt install libreoffice python3-uno"
            in text_error
        )
        assert "docs_libreoffice_parameters=" in text_error

    def test_convert_docx_to_pdf_reports_missing_libreoffice_executable(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_pdf_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_in_docx = path_tmp / "report.docx"
            file_in_docx.write_bytes(b"docx payload")
            options_pdf = DocxToPdfOptions(
                exe_libreoffice=path_tmp / "missing-libreoffice",
                file_in_docx=file_in_docx,
                file_out_pdf=path_tmp / "report.pdf",
                dir_user_profile=path_tmp / "lo-profile",
                should_update_fields=False,
            )

            from docxrender import pdf_uno

            with (
                mock.patch.object(pdf_uno, "import_uno_module", return_value=object()),
                pytest.raises(FileNotFoundError) as ctx,
            ):
                convert_docx_to_pdf(options_pdf)

        text_error = str(ctx.value)
        assert "error_code=libreoffice_executable_missing" in text_error
        assert "missing-libreoffice" in text_error
        assert (
            "install_debian_ubuntu=sudo apt install libreoffice python3-uno"
            in text_error
        )

    def test_libreoffice_listener_start_failure_includes_guidance(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_pdf_") as dir_tmp:
            path_tmp = Path(dir_tmp)

            from docxrender import pdf_uno

            with (
                mock.patch.object(pdf_uno, "validate_libreoffice_executable"),
                mock.patch.object(
                    pdf_uno.subprocess,
                    "Popen",
                    side_effect=PermissionError("not executable"),
                ),
                pytest.raises(
                    RuntimeError,
                    match="libreoffice_listener_start_failed",
                ) as ctx,
            ):
                start_libreoffice_listener(
                    exe_libreoffice=Path("/usr/bin/libreoffice"),
                    dir_user_profile=path_tmp / "lo-profile",
                    port=23001,
                    stdout=None,
                    stderr=None,
                    file_listener_log=path_tmp / "listener.log",
                )

        text_error = str(ctx.value)
        assert "launch_error=PermissionError" in text_error
        assert "listener.log" in text_error
        assert "docs_libreoffice_api=" in text_error

    def test_libreoffice_listener_timeout_includes_guidance(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_pdf_") as dir_tmp:
            file_listener_log = Path(dir_tmp) / "listener.log"
            file_listener_log.write_text("cannot start", encoding="utf-8")

            from docxrender import pdf_uno

            with (
                mock.patch.object(pdf_uno, "LISTENER_START_TIMEOUT_SECONDS", 0.0),
                pytest.raises(
                    TimeoutError,
                    match="libreoffice_uno_listener_timeout",
                ) as ctx,
            ):
                wait_for_listener(23001, file_listener_log=file_listener_log)

        text_error = str(ctx.value)
        assert "listener_port=23001" in text_error
        assert "listener_log_tail=cannot start" in text_error
        assert "validate_libreoffice=libreoffice --headless --version" in text_error

    def test_docx_field_update_markers_are_written(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_fields_") as dir_tmp:
            file_docx = Path(dir_tmp) / "fields.docx"
            _write_minimal_field_docx(file_docx)

            write_docx_field_update_markers(file_docx)

            text_settings = _read_docx_part(file_docx, "word/settings.xml")
            text_document = _read_docx_part(file_docx, "word/document.xml")
            assert '<w:updateFields w:val="true"/>' in text_settings
            assert 'w:dirty="true"' in text_document

    def test_docx_field_freeze_preserves_result_text(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_fields_") as dir_tmp:
            file_docx = Path(dir_tmp) / "fields.docx"
            _write_minimal_field_docx(file_docx)
            write_docx_field_update_markers(file_docx)

            write_frozen_docx_fields(file_docx)

            text_settings = _read_docx_part(file_docx, "word/settings.xml")
            text_document = _read_docx_part(file_docx, "word/document.xml")
            assert "w:updateFields" not in text_settings
            assert "fldChar" not in text_document
            assert "instrText" not in text_document
            assert "w:dirty" not in text_document
            assert "Rendered TOC" in text_document

    def test_docxrender_does_not_import_product_repositories(self) -> None:
        product_module_prefixes = (
            "proteomics",
            "trait_association",
            "joint_proteome_phospho",
        )

        imported_product_modules = [
            name
            for name in sys.modules
            if name in product_module_prefixes
            or name.startswith(
                tuple(f"{prefix}." for prefix in product_module_prefixes)
            )
        ]

        assert imported_product_modules == []


def _write_template(file_template: Path) -> None:
    document = Document()
    document.add_paragraph("{{ report_title }}")
    document.add_paragraph("{{ body_anchor }}")
    document.save(str(file_template))


def _write_png(file_image: Path) -> None:
    file_image.write_bytes(
        base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
            "AAAADUlEQVR42mP8z8BQDwAFgwJ/lv6OSwAAAABJRU5ErkJggg=="
        )
    )


def _write_minimal_field_docx(file_docx: Path) -> None:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:p>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        "<w:r><w:instrText>TOC</w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>Rendered TOC</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "</w:p></w:body></w:document>"
    )
    settings_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "</w:settings>"
    )
    with zipfile.ZipFile(file_docx, "w") as zip_out:
        zip_out.writestr("word/document.xml", document_xml)
        zip_out.writestr("word/settings.xml", settings_xml)


def _write_minimal_field_docx_without_result(file_docx: Path) -> None:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:p>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        "<w:r><w:instrText>TOC</w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "</w:p></w:body></w:document>"
    )
    settings_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "</w:settings>"
    )
    with zipfile.ZipFile(file_docx, "w") as zip_out:
        zip_out.writestr("word/document.xml", document_xml)
        zip_out.writestr("word/settings.xml", settings_xml)


def _read_docx_part(file_docx: Path, name: str) -> str:
    with zipfile.ZipFile(file_docx) as zip_in:
        return zip_in.read(name).decode("utf-8")


def _first_text_run(paragraph: Paragraph) -> Run:
    return next(run for run in paragraph.runs if run.text)


def _run_font_size_pt(run: Run) -> float:
    size = run.font.size
    if size is None:
        raise AssertionError("Expected text run to have an explicit font size.")
    return size.pt
