# -*- coding: utf-8 -*-
"""
Created : 2016-07-19

@author: AhnSeongHyun

Edited : 2016-07-19 by Eric Lapouyade
"""

from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/header_footer_tpl_utf8.docx")

sd = tpl.new_subdoc()
p = sd.add_paragraph(
    "This is a sub-document to check it does not break header and footer with utf-8 "
    "characters inside the template .docx"
)

context = {
    "title": "헤더와 푸터",
    "company_name": "세계적 회사",
    "date": "2016-03-17",
    "mysubdoc": sd,
}

tpl.render(context)
tpl.save("output/header_footer_utf8.docx")
