# docxrender

`docxrender` is a small Python package for Word-first DOCX rendering.

Its core boundary is intentionally narrow:

```text
file_template + context + markdown_body + DocxStyle -> DOCX -> PDF
```

The package owns technical rendering mechanics: DOCX template rendering,
markdown body insertion, Word style application, DOCX field handling, and
eventual LibreOffice-based PDF conversion. Product repositories own report
content, workflow resource layout, section rendering, manifest schemas, figure
selection, captions, and delivery directory semantics.

## Capabilities

Current package surface:

- Public style/options/result dataclasses are available.
- `write_docx(...)` can create a minimal DOCX from a DOCX template, context,
  markdown body, image assets, and `DocxStyle`.
- Markdown support currently covers a CommonMark-ish shared subset: headings,
  paragraphs, hard line breaks, unordered lists, ordered lists, pipe tables,
  images, inline bold, inline code text cleanup, link-text reduction, page
  breaks, and spacers.
- Basic Word styling is applied from caller-provided `DocxStyle`.
- DOCX field update/freeze behavior is implemented through DOCX XML rewriting.
- `write_docx(...)` can optionally refresh TOC/page fields through LibreOffice
  UNO when `DocxFieldRefreshOptions` is provided.
- `convert_docx_to_pdf(...)` converts through LibreOffice UNO when the external
  LibreOffice/UNO runtime is available.

## Install

```bash
pdm install
```

Runtime dependencies are declared in `pyproject.toml`:

- `docxtpl`
- `python-docx`

PDF conversion and DOCX field refresh are optional runtime features. They do
require an external LibreOffice/UNO runtime.

```bash
libreoffice --headless --version
python -c "import uno"
```

On Debian or Ubuntu, that runtime is typically installed outside Python:

```bash
sudo apt install libreoffice python3-uno
```

Base DOCX writing with `field_refresh=None` does not import UNO and works
without LibreOffice.

## Public API

The stable public API is exported from the package root. Product repositories
should prefer `DocxRenderer` for normal use. The dataclasses and module-level
functions remain public for advanced callers that want explicit contracts,
configuration adapters, or focused tests. Implementation modules such as
`docxrender.markdown` and `docxrender.docx` are technical layers and are not
compatibility-stable public contracts.

`DocxRenderer` follows value semantics. Every `with_*` call returns a new
renderer object and leaves the original unchanged. Runtime `docxtpl` objects,
including inline images, are materialized only by terminal methods such as
`write_docx_template(...)`, `write_docx(...)`, and `write_pdf()`.

```python
from docxrender import (
    DocxRenderer,
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
    DocxWriteOptions,
    write_docx_template,
    write_docx,
)
```

`DocxFieldMarkerOptions` controls DOCX field update markers and field freezing
without LibreOffice or UNO:

```python
DocxRenderer(file_docx=Path("report.docx")).with_field_update_markers(
    should_update_fields=True,
    should_freeze_fields=False,
).write_docx()
```

`DocxFieldRefreshOptions` is optional. Use it only when the caller has provided
a LibreOffice/UNO runtime and wants a DOCX whose TOC, page fields, or other
Word fields have been refreshed by LibreOffice:

```python
DocxWriteOptions(
    ...,
    field_refresh=DocxFieldRefreshOptions(
        exe_libreoffice=Path("/usr/bin/libreoffice"),
        dir_user_profile=Path("tmp/lo-profile"),
        should_require_toc=True,
        should_freeze_fields=True,
    ),
)
```

`write_docx_template(...)` is the generic `docxtpl` technical boundary. It
renders a DOCX template with caller context, optional inline images, and
optional default injections. It does not insert markdown bodies, apply DOCX
body styling, or run field/PDF post-processing:

```python
from pathlib import Path

from docxrender import DocxTemplateRenderOptions, write_docx_template

result = write_docx_template(
    DocxTemplateRenderOptions(
        file_template=Path("template.docx"),
        file_out_docx=Path("template-rendered.docx"),
        context={"report_title": "Example Report"},
        context_defaults={"body_anchor": "__REPORT_BODY_ANCHOR__"},
        context_policy=DocxTemplateContextPolicy(
            rule_conflict="caller_wins",
            required_keys=("report_title",),
        ),
    )
)
print(result.file_docx)
```

`DocxTemplateContextPolicy` controls:

- merge behavior between caller context and default injections
- conflict priority: `caller_wins` or `defaults_win`
- required keys that must exist after merge and before render

Minimal `DocxRenderer` DOCX write example:

```python
from pathlib import Path

from docxrender import DocxRenderer

result = (
    DocxRenderer()
    .with_template(
        file_template=Path("template.docx"),
        context={"report_title": "Example Report"},
        rule_conflict="caller_wins",
        required_keys=("report_title",),
    )
    .with_fonts(
        font_name_latin="Times New Roman",
        font_name_body_east_asia="宋体",
        font_name_heading_east_asia="宋体",
    )
    .with_sizes(
        pt_title_page_title=36.0,
        pt_title_page_meta=18.0,
        pt_title_page_compiler=15.0,
        pt_body=12.0,
        pt_caption=10.5,
        pt_table=12.0,
        pt_heading_by_level={1: 16.0, 2: 14.0, 3: 12.0},
    )
    .with_table(
        border_color="000000",
        stripe_fill_color="D9D9D9",
        border_size_main="12",
        border_size_header="6",
        line_spacing=1.5,
    )
    .with_paragraph(
        line_spacing_body=1.5,
        line_spacing_note=1.2,
        first_line_indent_cm=0.74,
    )
    .with_header_footer_images(
        file_header_image=Path("header.png"),
        file_footer_image=Path("footer.png"),
        idx_section_start=1,
    )
    .with_markdown(
        should_parse_inline_bold=True,
        should_parse_inline_code=True,
        should_parse_links_as_text=True,
        should_parse_image_width_attr=True,
        default_image_width_pct=90.0,
    )
    .with_body_render_policy(
        should_number_headings=False,
        rule_ordered_list="word_style",
        rule_unordered_list="word_style",
        should_stripe_table_rows=False,
    )
    .with_body_anchor(rule_match="equals", rule_missing="raise")
    .write_docx(
        file_out_docx=Path("report.docx"),
        markdown_body="# Summary **Bold**\n\nBody text with [link](https://example.com).",
        dir_base=Path("."),
    )
)
print(result.file_docx)
```

`markdown_body` is the already-rendered Markdown body to insert into the DOCX
template. `dir_base` is the base directory used to resolve relative image paths
inside that Markdown body.

`DocxMarkdownOptions` only covers the shared CommonMark-ish subset. Product
repositories should keep custom markdown dialects outside `docxrender`, either
through caller-side preprocessing or explicit higher-level options in their own
repo.

`DocxBodyRenderPolicy` controls structural rendering choices such as heading
numbering, Word-style versus plain-text lists, and striped table body rows. It
does not classify product-specific paragraphs.

`DocxBodyAnchorOptions` controls where the Markdown body is inserted. The search
is limited to top-level body paragraphs in the DOCX main document. `equals`
matches `paragraph.text.strip() == anchor_token`; `contains` matches templates
where the token is embedded in a larger paragraph. Missing anchors can either
append content or raise a template error.

`DocxRenderer` can also start from an existing DOCX and run only later
technical steps:

```python
from pathlib import Path

from docxrender import DocxRenderer

DocxRenderer(file_docx=Path("report.docx")).with_field_refresh(
    exe_libreoffice=Path("/usr/bin/libreoffice"),
    dir_user_profile=Path("tmp/lo-profile"),
    should_require_toc=True,
).write_docx()
```

Generic `docxtpl` inline-image binding is also supported through the same
template entrypoint:

```python
from pathlib import Path

from docxrender import DocxRenderer, DocxTemplateImageSpec

renderer = DocxRenderer().with_template(
    file_template=Path("template.docx"),
    context={"report_title": "Example Report"},
    inline_images={
        "cover_image": DocxTemplateImageSpec(
            file_image=Path("cover.png"),
            width_mm=120,
        ),
    },
)
```

The same renderer can convert the current DOCX to PDF:

```python
from pathlib import Path

from docxrender import DocxRenderer

result = (
    DocxRenderer(file_docx=Path("report.docx"))
    .with_pdf_conversion(
        exe_libreoffice=Path("/usr/bin/libreoffice"),
        dir_user_profile=Path("tmp/lo-profile"),
        file_out_pdf=Path("report.pdf"),
    )
    .write_pdf()
)
print(result.file_pdf)
```

Advanced explicit dataclass DOCX write example:

```python
from pathlib import Path

from docxrender import (
    DocxFontStyle,
    DocxParagraphStyle,
    DocxSizeStyle,
    DocxStyle,
    DocxTableStyle,
    DocxWriteOptions,
    write_docx,
)

style = DocxStyle(
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

result = write_docx(
    DocxWriteOptions(
        file_template=Path("template.docx"),
        file_out_docx=Path("report.docx"),
        context={"report_title": "Example Report"},
        markdown_body="# Summary\n\nBody text.",
        dir_base=Path("."),
        style=style,
    )
)
print(result.file_docx)
```

The template should contain a paragraph whose text is the body anchor token:

```text
{{ body_anchor }}
```

`docxrender` sets `body_anchor` in the template context when the caller does not
provide it.

## Style Configuration

`docxrender` does not read TOML, JSON, YAML, or any other config file in its public
API. Callers convert their own configuration into `DocxStyle`.

The initial style model is based on:

```text
/home/fqzhang/project/workflows/resources/common/report/style.toml
```

That file is a reference for fields and defaults, not a runtime dependency of
the package.

## Non-Goals

`docxrender` does not own:

- report manifest schemas
- workflow resource layout
- Jinja section discovery
- product-specific context builders
- figure registries or captions
- `Result/...` delivery path semantics
- `结果目录` text generation
- style config file readers

## Tests

Run the current test suite:

```bash
pdm run python -m pytest -v
```

`ty` is available as an advisory type checker beside pyright:

```bash
pdm run ty check .
```

Pyright remains the primary type gate.

The suite currently covers public API construction, minimal DOCX writing,
markdown body insertion, basic style application, and the boundary that
`docxrender` does not import product repositories.
