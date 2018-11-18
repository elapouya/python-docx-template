# -*- coding: utf-8 -*-
'''
Created : 2017-09-09

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

# rendering the "dynamic embedded docx":
embedded_docx_tpl=DocxTemplate('templates/embedded_embedded_docx_tpl.docx')
context = {
    'name' : 'John Doe',
}
embedded_docx_tpl.render(context)
embedded_docx_tpl.save('output/embedded_embedded_docx.docx')


# rendring the main document :
tpl=DocxTemplate('templates/embedded_main_tpl.docx')

context = {
    'name' : 'John Doe',
}

tpl.replace_embedded('templates/embedded_dummy.docx','templates/embedded_static_docx.docx')
tpl.replace_embedded('templates/embedded_dummy2.docx','templates/embedded_embedded_docx.docx')
tpl.render(context)
tpl.save('output/embedded.docx')