from docxtpl import *

tpl = DocxTemplate("test_files/escape_tpl.docx")

context = {'myvar': R('"less than" must be escaped : <, this can be done with RichText() or R()'),
           'myescvar':'It can be escaped with a "|e" jinja filter in the template too : < ',
           'nlnp' : R('Here is a multiple\nlines\nstring\aand some\aother\aparagraphs\aNOTE: the current character styling is removed'),
           'mylisting': Listing('the listing\nwith\nsome\nlines\nand special chars : <>&'),
           }

tpl.render(context)
tpl.save("test_files/escape.docx")
