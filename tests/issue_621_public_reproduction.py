# -*- coding: utf-8 -*-

from pathlib import Path
from tempfile import TemporaryDirectory

from docxtpl import DocxTemplate

from _issue_621_ole import (
    assert_file_identity,
    assert_ole_integrity,
    assert_xlsx_fixture,
)


MAIN_FIXTURE = Path("templates/issue_621_main.docx")
SUBDOC_FIXTURE = Path("templates/issue_621_subdoc.docx")
EMBEDDING_ZIPNAME = "word/embeddings/Microsoft_Excel_Worksheet.xlsx"
XLSX_FIXTURES = [
    (
        Path("templates/issue_621_excel_0.xlsx"),
        4887,
        "3b7d1befbe9fe6d4dfa8bee06b3dc424774b37079a884e89427c78577516805f",
        "No.1",
    ),
    (
        Path("templates/issue_621_excel_1.xlsx"),
        4867,
        "d353a37445081eb1dd471a5f9cdeec55a5222de128483d8998283776c6186f15",
        "No.2",
    ),
    (
        Path("templates/issue_621_excel_2.xlsx"),
        4868,
        "291cf1832083f0c0476679fb0030c1f818c079ef4326e5e064b9ce76c541124e",
        "No.3",
    ),
    (
        Path("templates/issue_621_excel_3.xlsx"),
        4867,
        "ff759e011e2cf1216aea110819e46cde04f3726a6c7b6d0b67b53cf25a7535a3",
        "No.4",
    ),
    (
        Path("templates/issue_621_excel_4.xlsx"),
        4867,
        "7a26b676c693c8f58deeec4879751257a02d6955957a0fcad5bbdef9ada55396",
        "No.5",
    ),
    (
        Path("templates/issue_621_excel_5.xlsx"),
        4867,
        "a4f3b5c2de16078dbb77757437fbbeca1ceb2e7d99d99b2dbaba609cf9b83197",
        "No.6",
    ),
    (
        Path("templates/issue_621_excel_6.xlsx"),
        4868,
        "51d03cae946ff4763a29049c08acd6a2af5cb1fc167fbce25d5ca53c96227ed6",
        "No.7",
    ),
]


assert_file_identity(
    MAIN_FIXTURE,
    12772,
    "8445e28450aa911b92d83b251ff634c2f49d8216a6cc829832965ef984f14ca0",
)
assert_file_identity(
    SUBDOC_FIXTURE,
    23798,
    "c57714a23c1d17f991f95dff21c1e6931195fd7f8848258e82a34586bb5a302d",
)
for fixture in XLSX_FIXTURES:
    assert_xlsx_fixture(*fixture)

with TemporaryDirectory() as temporary_directory:
    output_dir = Path(temporary_directory)
    result = output_dir / "issue_621_public_reproduction.docx"
    template = DocxTemplate(MAIN_FIXTURE)
    sub_docs = []
    for index, (xlsx_path, _, _, _) in enumerate(XLSX_FIXTURES):
        subdoc_path = output_dir / ("issue_621_subdoc_%d.docx" % index)
        subdoc_template = DocxTemplate(SUBDOC_FIXTURE)
        subdoc_template.replace_zipname(EMBEDDING_ZIPNAME, str(xlsx_path))
        subdoc_template.save(subdoc_path)
        sub_docs.append(template.new_subdoc(subdoc_path))

    template.render({"sub_docs": sub_docs})
    template.save(result)
    summary = assert_ole_integrity(
        result,
        expected_count=7,
        expected_anchor_count=7,
        required_embedding_member="xl/workbook.xml",
        expected_embedding_sha256=[item[2] for item in XLSX_FIXTURES],
        expected_workbook_sheets=[item[3] for item in XLSX_FIXTURES],
    )
    assert summary["ole_count"] == len(XLSX_FIXTURES)
