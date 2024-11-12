# -*- coding: utf-8 -*-
"""
Created : 2024-09-23

@author: Bart Broere
"""

from docxtpl import DocxTemplate

DEST_FILE = "output/footnotes.docx"

tpl = DocxTemplate("templates/footnotes_tpl.docx")

context = {
    "a_jinja_variable": "A Jinja variable!"
}

tpl.render(context)
tpl.save(DEST_FILE)
