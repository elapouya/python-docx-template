from docxtpl import DocxTemplate

tpl = DocxTemplate('templates/comments.docx')

tpl.render({})
tpl.save('output/comments.docx')
