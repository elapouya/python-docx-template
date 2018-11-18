# -*- coding: utf-8 -*-
'''
Created : 2017-09-03

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

DEST_FILE = 'output/replace_picture.docx'

tpl=DocxTemplate('templates/replace_picture_tpl.docx')

context = {}

tpl.replace_pic('python_logo.png','templates/python.png')
tpl.render(context)
tpl.save(DEST_FILE)