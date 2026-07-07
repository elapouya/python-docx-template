# -*- coding: utf-8 -*-
import re
import zipfile

from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/hyperlink_tpl.docx")
context = {"foo": "https://example.com"}
tpl.render(context)
output_path = "output/hyperlink.docx"
tpl.save(output_path)

with zipfile.ZipFile(output_path) as docx:
    rels = docx.read("word/_rels/document.xml.rels").decode()
    document = docx.read("word/document.xml").decode()

assert 'Target="https://example.com"' in rels, rels
assert "{{ foo }}" not in rels, rels
assert re.search(r"<w:t[^>]*>https://example.com</w:t>", document), document

undeclared = DocxTemplate("templates/hyperlink_tpl.docx").get_undeclared_template_variables(
    context=context
)
assert undeclared == set(), undeclared

print("hyperlink test passed")
