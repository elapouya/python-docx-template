# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

tpl=DocxTemplate('test_files/header_footer_tpl.docx')

context = {
    'title' : 'Header and footer test',
    'company_name' : 'The World Wide company',
    'date' : '2016-03-17'
}

tpl.render(context)
tpl.save('test_files/header_footer.docx')