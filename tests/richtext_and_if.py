# -*- coding: utf-8 -*-
'''
Created : 2015-03-26

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate, RichText

tpl=DocxTemplate('test_files/richtext_and_if_tpl.docx')


context = {
    'foobar': RichText('Foobar!', color='ff0000')
}

tpl.render(context)
tpl.save('test_files/richtext_and_if.docx')
