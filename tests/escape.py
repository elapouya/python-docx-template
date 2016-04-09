from docxtpl import *

tpl = DocxTemplate("test_files/escape_tpl.docx")

context = {'myvar': R('"less than" must be escaped : <, this can be done with RichText() or R()'),
           'myescvar':'It can be escaped with a "|e" jinja filter in the template too : < ',
           'nlnp' : R('Here is a multiple\nlines\nstring\aand some\aother\aparagraphs')}

tpl.render(context)
tpl.save("test_files/escape.docx")
