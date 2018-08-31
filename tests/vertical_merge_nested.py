from docxtpl import DocxTemplate

tpl = DocxTemplate('test_files/vertical_merge_nested.docx')
tpl.render()
tpl.save('test_files/word2016.docx')
