from docxtpl import DocxTemplate, RichText

tpl = DocxTemplate('test_files/word2016_tpl.docx')
tpl.render({
    'test_space' : '          ',
    'test_tabs': 5*'\t',
    'test_space_r' : RichText('          '),
    'test_tabs_r': RichText(5*'\t'),
})
tpl.save('test_files/word2016.docx')
