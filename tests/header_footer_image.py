# -*- coding: utf-8 -*-
'''
Created : 2017-09-03

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

DEST_FILE = 'test_files/header_footer_image.docx'

tpl=DocxTemplate('test_files/header_footer_image_tpl.docx')

context = {
    'mycompany' : 'The World Wide company',
}
tpl.replace_media('test_files/dummy_pic_for_header.png','test_files/python.png')
tpl.render(context)
tpl.save(DEST_FILE)