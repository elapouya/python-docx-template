import os.path

import jinja2
from docxtpl import DocxTemplate, InlineImage

doctemplate = r'templates/doc_properties_tpl.docx'

tpl = DocxTemplate(doctemplate)

context = {
    'test': 'HelloWorld'
}

tpl.render(context)
tpl.save("output/doc_properties.docx")
