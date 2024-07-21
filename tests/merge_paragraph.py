# -*- coding: utf-8 -*-
"""
Created : 2015-03-12

@author: Eric Lapouyade
"""

from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/merge_paragraph_tpl.docx")

context = {
    "living_in_town": True,
}

tpl.render(context)
tpl.save("output/merge_paragraph.docx")
