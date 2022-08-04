# -*- coding: utf-8 -*-
'''
Created : 2022-08-03
@author: Dongfang Song
'''
from docxtpl import DocxTemplate, RichText
tpl=DocxTemplate('templates/richtext_eastAsia_tpl.docx')
rt = RichText('测试TEST', font = 'Microsoft YaHei')
ch = RichText('测试TEST', font = '微软雅黑')
sun = RichText('测试TEST', font = 'SimSun')
context = {
    'example' : rt,
    'Chinese' : ch,
    'simsun' : sun,
}

tpl.render(context)
tpl.save('output/richtext_eastAsia.docx')
