# Regression test for #559: pandoc-produced tables collapse empty cell
# properties to a self-closing <w:tcPr/>, which used to make cellbg and
# colspan insert <w:shd>/<w:gridSpan> as a sibling after it instead of
# inside, so Word silently ignored them. No template file needed, the
# checks work on patch_xml output directly.

import re

from docxtpl import DocxTemplate

tpl = DocxTemplate("dummy.docx")


def patched(body):
    return tpl.patch_xml("<w:tbl><w:tr><w:tc>%s</w:tc></w:tr></w:tbl>" % body)


def assert_inside_tcpr(xml, element):
    m = re.search(r"<w:tcPr[^/>]*>(.*?)</w:tcPr>", xml, flags=re.DOTALL)
    assert m, "no expanded <w:tcPr> pair in: %s" % xml
    assert element in m.group(1), "%s not inside <w:tcPr>: %s" % (element, xml)


# self-closing <w:tcPr/> (the #559 case)
xml = patched("<w:tcPr/><w:p><w:r><w:t>{% cellbg color %}x</w:t></w:r></w:p>")
assert_inside_tcpr(xml, "<w:shd ")

xml = patched("<w:tcPr/><w:p><w:r><w:t>{% colspan span %}x</w:t></w:r></w:p>")
assert_inside_tcpr(xml, "<w:gridSpan ")

# self-closing with attributes
xml = patched('<w:tcPr w:dummy="1"/><w:p><w:r><w:t>{% cellbg color %}x</w:t></w:r></w:p>')
assert_inside_tcpr(xml, "<w:shd ")

# already expanded <w:tcPr> keeps working as before
xml = patched(
    '<w:tcPr><w:tcW w:w="100"/></w:tcPr><w:p><w:r><w:t>{% cellbg color %}x</w:t></w:r></w:p>'
)
assert_inside_tcpr(xml, "<w:shd ")
assert_inside_tcpr(xml, "<w:tcW ")

print("cellbg_selfclosing ok")
