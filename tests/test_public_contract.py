from __future__ import annotations

import base64
import sys
import tempfile
import types
import unittest
import zipfile
from pathlib import Path
from typing import Any, cast
from unittest import mock

from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docxrender import (
    DocxFieldRefreshOptions,
    DocxFontStyle,
    DocxParagraphStyle,
    DocxSizeStyle,
    DocxStyle,
    DocxTableStyle,
    DocxToPdfOptions,
    DocxToPdfResult,
    DocxWriteOptions,
    DocxWriter,
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


class PublicContractTest(unittest.TestCase):
    def test_public_imports_are_explicit(self) -> None:
        import docxrender

        self.assertEqual(
            docxrender.__all__,
            [
                "DocxWriter",
                "DocxFieldRefreshOptions",
                "DocxFontStyle",
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
            ],
        )

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

            self.assertEqual(options.anchor_token, "__REPORT_BODY_ANCHOR__")
            self.assertIs(options.should_update_fields, True)
            self.assertIs(options.should_freeze_fields, False)
            self.assertIsNone(options.field_refresh)
            self.assertEqual(options.style.paragraph.first_line_indent_cm, 0.74)

    def test_docx_to_pdf_options_construct_from_conversion_inputs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            options = DocxToPdfOptions(
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                file_in_docx=path_tmp / "report.docx",
                file_out_pdf=path_tmp / "report.pdf",
                dir_user_profile=path_tmp / "lo-profile",
            )

            self.assertIsNone(options.file_out_docx_refreshed)
            self.assertIsNone(options.file_listener_log)

    def test_public_results_are_structured_paths(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            result_docx = DocxWriteResult(file_docx=path_tmp / "report.docx")
            result_pdf = DocxToPdfResult(file_pdf=path_tmp / "report.pdf")

            self.assertEqual(result_docx.file_docx.name, "report.docx")
            self.assertEqual(result_pdf.file_pdf.name, "report.pdf")
            self.assertIsNone(result_pdf.file_docx_refreshed)

    def test_docx_size_style_with_overrides_changes_selected_values(self) -> None:
        sizes = create_docx_style().sizes

        updated = sizes.with_overrides(pt_body=11.0)

        self.assertEqual(updated.pt_body, 11.0)
        self.assertEqual(updated.pt_caption, sizes.pt_caption)
        self.assertEqual(updated.pt_heading_by_level, sizes.pt_heading_by_level)

    def test_docx_size_style_with_overrides_copies_heading_sizes(self) -> None:
        sizes = create_docx_style().sizes
        heading_sizes = {1: 18.0, 2: 15.0}

        updated = sizes.with_overrides(pt_heading_by_level=heading_sizes)
        heading_sizes[1] = 99.0

        self.assertEqual(updated.pt_heading_by_level, {1: 18.0, 2: 15.0})
        self.assertEqual(sizes.pt_heading_by_level[1], 16.0)

    def test_docx_writer_style_returns_default_style(self) -> None:
        style = DocxWriter().style

        self.assertEqual(style.fonts.font_name_latin, "Times New Roman")
        self.assertEqual(style.fonts.font_name_body_east_asia, "宋体")
        self.assertEqual(style.sizes.pt_body, 12.0)
        self.assertEqual(style.sizes.pt_heading_by_level[1], 16.0)
        self.assertEqual(style.sizes.pt_heading_by_level[6], 12.0)
        self.assertEqual(style.table.border_color, "000000")
        self.assertEqual(style.paragraph.first_line_indent_cm, 0.74)

    def test_docx_writer_fluent_overrides_are_partial(self) -> None:
        style = (
            DocxWriter()
            .with_fonts(font_name_body_east_asia="黑体")
            .with_sizes(pt_body=11.0)
            .with_table(stripe_fill_color="FFFFFF")
            .with_paragraph(note_prefixes=("Note:",))
            .style
        )

        self.assertEqual(style.fonts.font_name_latin, "Times New Roman")
        self.assertEqual(style.fonts.font_name_body_east_asia, "黑体")
        self.assertEqual(style.sizes.pt_body, 11.0)
        self.assertEqual(style.sizes.pt_caption, 10.5)
        self.assertEqual(style.table.border_color, "000000")
        self.assertEqual(style.table.stripe_fill_color, "FFFFFF")
        self.assertEqual(style.paragraph.note_prefixes, ("Note:",))
        self.assertEqual(style.paragraph.line_spacing_body, 1.5)

    def test_docx_writer_build_style_matches_style_property(self) -> None:
        writer = DocxWriter().with_sizes(pt_body=11.0)

        self.assertIs(writer.build_style(), writer.style)

    def test_docx_writer_build_options_uses_fluent_state(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            field_refresh = DocxFieldRefreshOptions(
                exe_libreoffice=Path("/usr/bin/libreoffice"),
                dir_user_profile=path_tmp / "lo-profile",
            )

            options = (
                DocxWriter()
                .with_sizes(pt_body=11.0)
                .with_field_refresh(field_refresh)
                .build_options(
                    file_template=path_tmp / "template.docx",
                    file_out_docx=path_tmp / "report.docx",
                    context={"report_title": "Builder"},
                    markdown_body="Body.",
                    dir_base=path_tmp,
                )
            )

            self.assertEqual(options.style.sizes.pt_body, 11.0)
            self.assertIs(options.field_refresh, field_refresh)
            self.assertEqual(options.anchor_token, "__REPORT_BODY_ANCHOR__")

    def test_docx_writer_field_refresh_can_be_built_from_keywords(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)

            options = (
                DocxWriter()
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

            self.assertIsNotNone(options.field_refresh)
            self.assertEqual(
                cast(DocxFieldRefreshOptions, options.field_refresh).exe_libreoffice,
                Path("/usr/bin/libreoffice"),
            )
            self.assertIs(
                cast(DocxFieldRefreshOptions, options.field_refresh).should_require_toc,
                True,
            )

    def test_docx_writer_requires_field_refresh_runtime_paths(self) -> None:
        with self.assertRaisesRegex(ValueError, "exe_libreoffice"):
            DocxWriter().with_field_refresh()

    def test_docx_writer_write_docx_uses_core_writer(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_contract_") as dir_tmp:
            path_tmp = Path(dir_tmp)
            file_template = path_tmp / "template.docx"
            file_out_docx = path_tmp / "report.docx"
            _write_template(file_template)

            result = (
                DocxWriter()
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
            self.assertEqual(result.file_docx, file_out_docx)
            self.assertEqual(
                _run_font_size_pt(_first_text_run(paragraph_by_text["Body."])),
                11.0,
            )

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

            self.assertEqual(result.file_docx, file_out_docx)
            self.assertIn("Contract Report", texts)
            self.assertIn("Heading", texts)
            self.assertIn("Body first line.\nBody second line.", texts)
            self.assertIn("注：Note text.", texts)
            self.assertIn("First item", texts)
            self.assertIn("Second item", texts)
            self.assertIn("Example image", texts)
            self.assertEqual(document.tables[0].cell(0, 0).text, "A")
            self.assertEqual(document.tables[0].cell(1, 1).text, "2")
            self.assertEqual(len(document.inline_shapes), 1)
            paragraph_by_text = {
                paragraph.text: paragraph for paragraph in document.paragraphs
            }
            self.assertEqual(
                _run_font_size_pt(_first_text_run(paragraph_by_text["Heading"])),
                16.0,
            )
            self.assertEqual(
                _run_font_size_pt(
                    _first_text_run(
                        paragraph_by_text["Body first line.\nBody second line."]
                    )
                ),
                12.0,
            )
            self.assertEqual(
                _run_font_size_pt(_first_text_run(paragraph_by_text["注：Note text."])),
                10.5,
            )
            self.assertEqual(
                _run_font_size_pt(_first_text_run(paragraph_by_text["Example image"])),
                10.5,
            )
            self.assertIn(
                'w:val="single"',
                cast(Any, document.tables[0].cell(0, 0))._tc.xml,
            )

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

            self.assertEqual(result.file_docx, file_out_docx)
            self.assertTrue(file_out_docx.exists())

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

            self.assertEqual(result.file_docx, file_out_docx)
            self.assertEqual(file_out_docx.read_bytes(), b"refreshed")

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
                self.assertEqual(file_in, file_out_docx)
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

            self.assertEqual(result.file_docx, file_out_docx)
            self.assertTrue(file_out_docx.exists())
            self.assertEqual(file_refreshed.read_bytes(), b"refreshed separate")

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
            self.assertIn("Rendered TOC", text_document)
            self.assertNotIn("fldChar", text_document)

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
                with self.assertRaisesRegex(RuntimeError, "TOC result"):
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

            self.assertEqual(command[0], "/usr/bin/libreoffice")
            self.assertIn("--headless", command)
            self.assertIn(
                "--accept=socket,host=127.0.0.1,port=23001;urp;",
                command,
            )
            self.assertTrue(command[-1].startswith("-env:UserInstallation=file://"))

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
                self.assertNotEqual(file_staged, file_in_docx)
                self.assertTrue(file_staged.exists())
                self.assertEqual(file_staged.read_bytes(), file_in_docx.read_bytes())
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

            self.assertEqual(result.file_pdf, file_out_pdf)
            self.assertEqual(result.file_docx_refreshed, file_out_docx_refreshed)
            self.assertEqual(captured["file_in_docx_source"], file_in_docx)
            self.assertEqual(captured["file_source_lock"], file_source_lock)
            self.assertEqual(captured["file_listener_log"], file_listener_log.resolve())
            self.assertTrue(fake_doc.stored)
            self.assertTrue(fake_doc.closed)
            self.assertEqual(len(fake_doc.store_url_calls), 1)
            self.assertEqual(
                file_out_docx_refreshed.read_bytes(),
                b"refreshed staged payload",
            )
            self.assertTrue(fake_process.terminated)

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
                with self.assertRaisesRegex(RuntimeError, "load failed"):
                    convert_docx_to_pdf(options_pdf)

            self.assertTrue(fake_process.terminated)

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

            self.assertIn("error_code=libreoffice_uno_load_failed", text_error)
            self.assertIn("reason_code=staged_docx_import_failed", text_error)
            self.assertIn(f"file_in_docx={file_source.resolve()}", text_error)
            self.assertIn(f"file_in_docx_staged={file_staged.resolve()}", text_error)
            self.assertIn(f"source_lock_file={file_source_lock.resolve()}", text_error)
            self.assertIn("probe_swriter_factory=ok", text_error)
            self.assertIn("load_staged_default_props=failed", text_error)
            self.assertIn("load_staged_hidden_only=failed", text_error)
            self.assertIn("listener_log_tail=listener tail", text_error)
            self.assertIn(
                "validate_libreoffice=libreoffice --headless --version",
                text_error,
            )
            self.assertIn(
                "install_debian_ubuntu=sudo apt install libreoffice python3-uno",
                text_error,
            )

    def test_pdf_uno_runtime_is_loaded_only_when_requested(self) -> None:
        with mock.patch("docxrender.pdf_uno.importlib.import_module") as import_module:
            import_module.side_effect = ImportError("missing uno")
            with self.assertRaisesRegex(RuntimeError, "libreoffice_uno_import_failed"):
                import_uno_module()
            import_module.assert_called_once_with("uno")

    def test_pdf_uno_import_failure_includes_install_guidance(self) -> None:
        with mock.patch("docxrender.pdf_uno.importlib.import_module") as import_module:
            import_module.side_effect = ImportError("missing uno")

            with self.assertRaises(RuntimeError) as ctx:
                import_uno_module()

        text_error = str(ctx.exception)
        self.assertIn("error_code=libreoffice_uno_import_failed", text_error)
        self.assertIn(
            "validate_libreoffice=libreoffice --headless --version",
            text_error,
        )
        self.assertIn('validate_uno=python -c "import uno"', text_error)
        self.assertIn(
            "install_debian_ubuntu=sudo apt install libreoffice python3-uno",
            text_error,
        )
        self.assertIn("docs_libreoffice_parameters=", text_error)

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
            )

            from docxrender import pdf_uno

            with (
                mock.patch.object(pdf_uno, "import_uno_module", return_value=object()),
                self.assertRaises(FileNotFoundError) as ctx,
            ):
                convert_docx_to_pdf(options_pdf)

        text_error = str(ctx.exception)
        self.assertIn("error_code=libreoffice_executable_missing", text_error)
        self.assertIn("missing-libreoffice", text_error)
        self.assertIn(
            "install_debian_ubuntu=sudo apt install libreoffice python3-uno",
            text_error,
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
                self.assertRaisesRegex(
                    RuntimeError,
                    "libreoffice_listener_start_failed",
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

        text_error = str(ctx.exception)
        self.assertIn("launch_error=PermissionError", text_error)
        self.assertIn("listener.log", text_error)
        self.assertIn("docs_libreoffice_api=", text_error)

    def test_libreoffice_listener_timeout_includes_guidance(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_pdf_") as dir_tmp:
            file_listener_log = Path(dir_tmp) / "listener.log"
            file_listener_log.write_text("cannot start", encoding="utf-8")

            from docxrender import pdf_uno

            with (
                mock.patch.object(pdf_uno, "LISTENER_START_TIMEOUT_SECONDS", 0.0),
                self.assertRaisesRegex(
                    TimeoutError,
                    "libreoffice_uno_listener_timeout",
                ) as ctx,
            ):
                wait_for_listener(23001, file_listener_log=file_listener_log)

        text_error = str(ctx.exception)
        self.assertIn("listener_port=23001", text_error)
        self.assertIn("listener_log_tail=cannot start", text_error)
        self.assertIn(
            "validate_libreoffice=libreoffice --headless --version",
            text_error,
        )

    def test_docx_field_update_markers_are_written(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_fields_") as dir_tmp:
            file_docx = Path(dir_tmp) / "fields.docx"
            _write_minimal_field_docx(file_docx)

            write_docx_field_update_markers(file_docx)

            text_settings = _read_docx_part(file_docx, "word/settings.xml")
            text_document = _read_docx_part(file_docx, "word/document.xml")
            self.assertIn('<w:updateFields w:val="true"/>', text_settings)
            self.assertIn('w:dirty="true"', text_document)

    def test_docx_field_freeze_preserves_result_text(self) -> None:
        with tempfile.TemporaryDirectory(prefix="docxrender_fields_") as dir_tmp:
            file_docx = Path(dir_tmp) / "fields.docx"
            _write_minimal_field_docx(file_docx)
            write_docx_field_update_markers(file_docx)

            write_frozen_docx_fields(file_docx)

            text_settings = _read_docx_part(file_docx, "word/settings.xml")
            text_document = _read_docx_part(file_docx, "word/document.xml")
            self.assertNotIn("w:updateFields", text_settings)
            self.assertNotIn("fldChar", text_document)
            self.assertNotIn("instrText", text_document)
            self.assertNotIn("w:dirty", text_document)
            self.assertIn("Rendered TOC", text_document)

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

        self.assertEqual(imported_product_modules, [])


if __name__ == "__main__":
    unittest.main()


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
