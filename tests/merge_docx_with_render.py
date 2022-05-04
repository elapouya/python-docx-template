from docxtpl import DocxTemplate

templateFileName = 'templates/merge_docx_with_render_master_tpl.docx'
subTemplateFileName = 'templates/merge_docx_with_render_subdoc.docx'
outFilename = 'output/merge_rendered_docx.docx'

doc = DocxTemplate(templateFileName)

mySubDoc = doc.new_subdoc(subTemplateFileName)

context = {
  'Parent_Doc_Field_1' : 'Foo',
  'Parent_Doc_Field_2' : mySubDoc
}

doc.render(context)

subContext = {
  'Sub_Doc_Field' : 'Bar'
}

doc.render(subContext)

doc.save(outFilename)
