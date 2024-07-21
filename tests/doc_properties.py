from docxtpl import DocxTemplate

doctemplate = r"templates/doc_properties_tpl.docx"

tpl = DocxTemplate(doctemplate)

context = {"test": "HelloWorld"}

tpl.render(context)
tpl.save("output/doc_properties.docx")
