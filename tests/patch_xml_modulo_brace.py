# Regression test for the {%p %}/{%tr %}/{%tc %}/{%r %} relocation regex.
#
# patch_xml() moves a {%p ... %} (or tr/tc/r) directive out of its surrounding
# <w:p>/<w:tr>/<w:tc>/<w:r> tags. The capture group used to be `[^}%]*`, which
# stopped at the first "%" or "}" *inside* the expression. As a result a
# directive whose Jinja expression contained a "%" (e.g. a modulo operation or a
# literal percent) or a "}" (e.g. a dict literal) was silently left in place,
# producing invalid XML. The expression is now captured up to the real closing
# "%}"/"}}", so such expressions round-trip correctly.

from docxtpl import DocxTemplate

# patch_xml() is a pure string transformation and never opens the file, so we
# can exercise it without a real .docx template.
tpl = DocxTemplate("dummy.docx")

CASES = [
    # "%" inside the expression (modulo / literal percent)
    (
        "<w:p><w:r><w:t>{%p set x = a % b %}</w:t></w:r></w:p>",
        "{% set x = a % b %}",
    ),
    # "}" inside the expression (dict literal)
    (
        "<w:p><w:r><w:t>{%p set d = {'a': 1} %}</w:t></w:r></w:p>",
        "{% set d = {'a': 1} %}",
    ),
    # same fix for a table-row directive carrying a "%"
    (
        "<w:tr><w:tc><w:p><w:r><w:t>{%tr for x in items if x % 2 %}"
        "</w:t></w:r></w:p></w:tc></w:tr>",
        "{% for x in items if x % 2 %}",
    ),
]

for src_xml, expected in CASES:
    result = tpl.patch_xml(src_xml)
    assert result == expected, "patch_xml(%r) -> %r, expected %r" % (
        src_xml,
        result,
        expected,
    )

print("patch_xml_modulo_brace: OK")
