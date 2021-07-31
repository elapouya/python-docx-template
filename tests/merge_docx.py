# -*- coding: utf-8 -*-
'''
Created : 2021-07-30

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate
from docx.shared import Inches

tpl = DocxTemplate('templates/merge_docx_master_tpl.docx')
sd = tpl.new_subdoc('templates/merge_docx_subdoc.docx')

context = {
    'mysubdoc': sd,
}

tpl.render(context)
tpl.save('output/merge_docx.docx')
