"""
@author: Max Podolskii
"""

import os
from unicodedata import name

from docxtpl import DocxTemplate


XML_RESERVED = """<"&'>"""

tpl = DocxTemplate("templates/escape_tpl_auto.docx")

context = {
    "nested_dict": {name(str(c)): c for c in XML_RESERVED},
    "autoescape": 'Escaped "str & ing"!',
    "autoescape_unicode": "This is an escaped <unicode> example \u4f60 & \u6211",
    "iteritems": lambda x: x.items(),
}

tpl.render(context, autoescape=True)

OUTPUT = "output"
if not os.path.exists(OUTPUT):
    os.makedirs(OUTPUT)
tpl.save(OUTPUT + "/escape_auto.docx")
