# -*- coding: utf-8 -*-

from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/horizontal_merge_tpl.docx")
tpl.render({})
tpl.save("output/horizontal_merge.docx")
