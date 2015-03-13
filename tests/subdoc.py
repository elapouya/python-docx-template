# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

tpl=DocxTemplate('test_files/subdoc_tpl.docx')

sd = tpl.new_subdoc()
p = sd.add_paragraph('This is a sub-document inserted into a bigger one')
p = sd.add_paragraph('It has been ')
p.add_run('dynamically').style = 'dynamic'
p.add_run(' generated with python by using ')
p.add_run('python-docx').italic = True
p.add_run(' library')

context = { 
    'mysubdoc' : sd,
}

tpl.render(context)
tpl.save('test_files/subdoc.docx')
