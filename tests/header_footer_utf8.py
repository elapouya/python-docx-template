# -*- coding: utf-8 -*-
'''
Created : 2016-07-19

@author: AhnSeongHyun

Edited : 2016-07-19 by Eric Lapouyade
'''

from docxtpl import DocxTemplate

tpl=DocxTemplate('test_files/header_footer_tpl_utf8.docx')

sd = tpl.new_subdoc()
p = sd.add_paragraph(u'This is a sub-document to check it does not break header and footer with utf-8 characters inside the template .docx')

context = {
    'title' : u'헤더와 푸터',
    'company_name' : u'세계적 회사',
    'date' : u'2016-03-17',
    'mysubdoc' : sd,
}

tpl.render(context)
tpl.save('test_files/header_footer_utf8.docx')
