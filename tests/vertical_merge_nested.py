from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/vertical_merge_nested_tpl.docx")
tpl.render({})
tpl.save("output/vertical_merge_nested.docx")
