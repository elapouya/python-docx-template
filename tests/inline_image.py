# -*- coding: utf-8 -*-
'''
Created : 2017-01-14

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

tpl=DocxTemplate('test_files/inline_image_tpl.docx')

context = {
    'myimage' : InlineImage(tpl,'/home/elapouya/tmp/dmozrated.png'),
}

tpl.render(context)
tpl.save('test_files/inline_image.docx')
