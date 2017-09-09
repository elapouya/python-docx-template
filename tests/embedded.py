# -*- coding: utf-8 -*-
'''
Created : 2017-09-09

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

# rendering the "dynamic embedded docx":
embedded_docx_tpl=DocxTemplate('test_files/embedded_embedded_docx_tpl.docx')
context = {
    'name' : 'John Doe',
}
embedded_docx_tpl.render(context)
embedded_docx_tpl.save('test_files/embedded_embedded_docx.docx')


# rendring the main document :
tpl=DocxTemplate('test_files/embedded_main_tpl.docx')

context = {
    'name' : 'John Doe',
}

tpl.replace_embedded('test_files/embedded_dummy.docx','test_files/embedded_static_docx.docx')
tpl.replace_embedded('test_files/embedded_dummy2.docx','test_files/embedded_embedded_docx.docx')
tpl.render(context)
tpl.save('test_files/embedded.docx')