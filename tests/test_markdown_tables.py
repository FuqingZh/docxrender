from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn

from docxrender import DocxRenderer
from docxrender.markdown import MarkdownTable, parse_markdown_blocks


@pytest.mark.parametrize(
    ("cell", "expected"),
    [
        ("宿主<br>外源序列", "宿主\n外源序列"),
        ("first<br/>second<BR />third", "first\nsecond\nthird"),
        ("<br>middle<br><br>", "\nmiddle\n\n"),
        (r"A\|B", "A|B"),
        (r"\|boundary\|", "|boundary|"),
        (r"C:\\data\\path", r"C:\data\path"),
        (r"slash\\\|pipe", r"slash\|pipe"),
        ("&lt;br&gt; &amp; &#124; &amp;lt;br&amp;gt;", "<br> & | &lt;br&gt;"),
        ("**first<br>second**", "first\nsecond"),
        (r"literal\n", r"literal\n"),
    ],
)
def test_cell_text_survives_public_docx_render(
    tmp_path: Path, cell: str, expected: str
) -> None:
    template = tmp_path / "template.docx"
    output = tmp_path / "result.docx"
    document = Document()
    document.add_paragraph("{{ body_anchor }}")
    document.save(str(template))
    DocxRenderer().with_template(file_template=template, context={}).write_docx(
        file_out_docx=output,
        markdown_body=f"| Name | Content |\n| --- | --- |\n| ref | {cell} |",
        dir_base=tmp_path,
        should_update_fields=False,
    )
    result = Document(str(output))
    assert len(result.tables) == 1
    table = result.tables[0]
    assert len(table.rows) == 2 and len(table.columns) == 2
    assert table.cell(1, 1).text == expected
    with ZipFile(output) as archive:
        xml = ElementTree.fromstring(archive.read("word/document.xml"))
    assert len(list(xml.iter(qn("w:br")))) == expected.count("\n")
    assert not [p.text for p in result.paragraphs if p.text]
    if cell.startswith("**"):
        assert all(
            run.bold for run in table.cell(1, 1).paragraphs[0].runs if run.text.strip()
        )


@pytest.mark.parametrize(
    ("row", "expected"),
    [
        ("| left | |", ["left", ""]),
        ("|| right |", ["", "right"]),
        ("| left ||", ["left", ""]),
        (r"| left | edge\|", ["left", "edge|"]),
        (r"| left\\|right |", ["left\\", "right"]),
        (r"| left\\\|right | end |", [r"left\|right", "end"]),
        (r"| left | &#124; |", ["left", "|"]),
    ],
)
def test_table_boundaries_and_escape_parity(row: str, expected: list[str]) -> None:
    blocks = parse_markdown_blocks("| A | B |\n| --- | --- |\n" + row)
    assert len(blocks) == 1
    table = blocks[0]
    assert isinstance(table, MarkdownTable)
    assert [
        "".join(segment.text for segment in cell) for cell in table.rows[1]
    ] == expected
