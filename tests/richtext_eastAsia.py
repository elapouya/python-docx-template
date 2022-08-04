from docxtpl import DocxTemplate, RichText
tpl=DocxTemplate('test_files/richtext_eastAsia_tpl.docx')
rt = RichText('测试TEST', font = 'Microsoft YaHei')
ch = RichText('测试TEST', font = '微软雅黑')
sun = RichText('测试TEST', font = 'SimSun')
context = {
    'example' : rt,
    'Chinese' : ch,
    'simsun' : sun,
}

tpl.render(context)
tpl.save('test_files/richtext_eastAsia.docx')