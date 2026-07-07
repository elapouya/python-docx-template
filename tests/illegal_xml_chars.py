# -*- coding: utf-8 -*-
"""
Created : 2026-07-07

@author: Jack Byrne

Checks that XML-1.0-illegal characters injected through context values (such as
the C0 control "\\x01") are stripped from rendered output, so docxtpl never emits
a part XML that fails strict parsing. Illegal chars are placed in the body, an
escaped run, the header, the footer and a core property.
"""

from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/illegal_xml_chars.docx")

context = {
    "myvar": "a\x01b",
    "escaped": "<c>\x01",
    "hdr": "header\x01",
    "ftr": "footer\x01",
    "prop": "property\x01",
}

tpl.render(context, autoescape=True)
tpl.save("output/illegal_xml_chars.docx")
