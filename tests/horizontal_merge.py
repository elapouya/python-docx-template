# -*- coding: utf-8 -*-

from docxtpl import DocxTemplate

tpl = DocxTemplate('test_files/horizontal_merge_tpl.docx')
tpl.render({})
tpl.save('test_files/horizontal_merge.docx')
