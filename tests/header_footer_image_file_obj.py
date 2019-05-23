# -*- coding: utf-8 -*-
'''
Created : 2019-05-22

@author: Eric Dufresne
'''

from docxtpl import DocxTemplate
import io

DEST_FILE = 'output/header_footer_image_file_obj.docx'

tpl=DocxTemplate('templates/header_footer_image_tpl.docx')

context = {
    'mycompany' : 'The World Wide company',
}

dummy_pic = io.BytesIO(open('templates/dummy_pic_for_header.png', 'rb').read())
new_image = io.BytesIO(open('templates/python.png', 'rb').read())
tpl.replace_media(dummy_pic, new_image)
tpl.render(context)
tpl.save(DEST_FILE)