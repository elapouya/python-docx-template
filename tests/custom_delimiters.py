"""Test that custom Jinja2 delimiters work with patch_xml.

This verifies that patch_xml properly strips XML tags from inside
user-configured Jinja2 blocks when using non-default delimiters
(like single braces {} instead of double braces {{}}).
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from docxtpl import DocxTemplate
from jinja2 import Environment
import zipfile, re

TEMPLATE = os.path.join(os.path.dirname(__file__),
                        "templates", "custom_delimiters_tpl.docx")
OUTPUT = os.path.join(os.path.dirname(__file__),
                      "output", "custom_delimiters.docx")


def _get_text_from_docx(path):
    """Extract plain text from a docx file for assertion."""
    with zipfile.ZipFile(path, "r") as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return re.sub(r"<[^>]+>", "", xml)


def test_custom_delimiters():
    """Custom { } delimiters should render correctly even when
    variables are split across multiple XML runs."""
    tpl = DocxTemplate(TEMPLATE)
    jinja_env = Environment(
        variable_start_string="{",
        variable_end_string="}",
    )
    tpl.render({"name": "Alice", "score": "95"}, jinja_env)
    tpl.save(OUTPUT)

    text = _get_text_from_docx(OUTPUT)
    print("Rendered text:", repr(text))

    # Both variables should be substituted
    assert "{name}" not in text, "Variable {name} was not rendered!"
    assert "{score}" not in text, "Variable {score} was not rendered!"
    assert "Alice" in text, "Name should appear in output"
    assert "95" in text, "Score should appear in output"

    # No leftover braces
    assert "{" not in text, "Leftover { in output"
    assert "}" not in text, "Leftover } in output"

    print("✅ custom_delimiters: PASS")


def test_default_delimiters_still_work():
    """Default {{ }} delimiters should still work (backward compat)."""
    import io as _io

    # Create a template with default delimiters
    default_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:rPr></w:rPr><w:t>Hello {{</w:t></w:r>
      <w:r><w:rPr></w:rPr><w:t>name</w:t></w:r>
      <w:r><w:rPr></w:rPr><w:t>}}!</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""
    buf = _io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""")
        zf.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""")
        zf.writestr("word/document.xml", default_xml)

    out_default = os.path.join(os.path.dirname(__file__),
                               "output", "custom_delimiters_default.docx")
    tpl = DocxTemplate(buf)
    tpl.render({"name": "Bob"})
    tpl.save(out_default)

    text = _get_text_from_docx(out_default)
    print("Rendered text (default):", repr(text))
    assert "Bob" in text, "Name should appear in output"
    assert "{{" not in text, "Leftover {{ in output"
    print("✅ default_delimiters: PASS")


if __name__ == "__main__":
    test_custom_delimiters()
    test_default_delimiters_still_work()
