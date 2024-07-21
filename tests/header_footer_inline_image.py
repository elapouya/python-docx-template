# -*- coding: utf-8 -*-
"""
Created : 2021-04-06

@author: Eric Lapouyade
"""

from docxtpl import DocxTemplate, InlineImage

# for height and width you have to use millimeters (Mm), inches or points(Pt) class :
from docx.shared import Mm

tpl = DocxTemplate("templates/header_footer_inline_image_tpl.docx")

context = {
    "inline_image": InlineImage(tpl, "templates/django.png", height=Mm(10)),
    "images": [
        InlineImage(tpl, "templates/python.png", height=Mm(10)),
        InlineImage(tpl, "templates/python.png", height=Mm(10)),
        InlineImage(tpl, "templates/python.png", height=Mm(10)),
    ],
}
tpl.render(context)
tpl.save("output/header_footer_inline_image.docx")
