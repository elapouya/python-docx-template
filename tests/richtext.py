# -*- coding: utf-8 -*-
'''
Created : 2015-03-26

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate, RichText

tpl=DocxTemplate('test_files/richtext_tpl.docx')

rt = RichText('an exemple of ')
rt.add('a rich text', style='myrichtextstyle')
rt.add(' with ')
rt.add('some italic', italic=True)
rt.add(' and ')
rt.add('some violet', color='#ff00ff')
rt.add(' and ')
rt.add('some striked', strike=True)
rt.add(' and ')
rt.add('some small', size=14)
rt.add(' or ')
rt.add('big', size=60)
rt.add(' text.')
rt.add(' Et voilà ! ')
rt.add('\n1st line')
rt.add('\n2nd line')
rt.add('\n3rd line')
rt.add('\n\n<cool>')
rt.add('\nFonts :\n',underline=True)
rt.add('Arial\n',font='Arial')
rt.add('Courier New\n',font='Courier New')
rt.add('Times New Roman\n',font='Times New Roman')

context = {
    'example' : rt,
}

tpl.render(context)
tpl.save('test_files/richtext.docx')
