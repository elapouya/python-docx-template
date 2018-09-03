from docxtpl import DocxTemplate

tpl = DocxTemplate('test_files/vertical_merge_nested_tpl.docx')
tpl.render({})
tpl.save('test_files/vertical_merge_nested.docx')
