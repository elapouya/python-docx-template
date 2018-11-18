# -*- coding: utf-8 -*-
'''
Created : 2017-09-03

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

DEST_FILE = 'output/header_footer_image.docx'

tpl=DocxTemplate('templates/header_footer_image_tpl.docx')

context = {
    'mycompany' : 'The World Wide company',
}
tpl.replace_media('templates/dummy_pic_for_header.png','templates/python.png')
tpl.render(context)
tpl.save(DEST_FILE)