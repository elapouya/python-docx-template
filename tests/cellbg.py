# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

tpl=DocxTemplate('test_files/cellbg_tpl.docx')

context = { 
    'alerts' : [
        {'date' : '2015-03-10', 'desc' : 'Very critical alert', 'type' : 'CRITICAL', 'bg': 'FF0000' },
        {'date' : '2015-03-11', 'desc' : 'Just a warning', 'type' : 'WARNING', 'bg': 'FFDD00' },
        {'date' : '2015-03-12', 'desc' : 'Information', 'type' : 'INFO', 'bg': '8888FF' },
        {'date' : '2015-03-13', 'desc' : 'Debug trace', 'type' : 'DEBUG', 'bg': 'FF00FF' },
    ], 
}

tpl.render(context)
tpl.save('test_files/cellbg.docx')
