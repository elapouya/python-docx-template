from docxtpl import *

tpl = DocxTemplate("test_files/escape_tpl_auto.docx")

context = {'myvar': R('"less than" must be escaped : <, this can be done with RichText() or R()'),
           'myescvar':'It can be escaped with a "|e" jinja filter in the template too : < ',
           'nlnp' : R('Here is a multiple\nlines\nstring\aand some\aother\aparagraphs\aNOTE: the current character styling is removed'),
           'mylisting': Listing('the listing\nwith\nsome\nlines\nand special chars : <>&'),
           'autoescape': """These string should be auto escaped for an XML Word document which may contain <, >, &, ", and '."""
           }

tpl.render(context)
tpl.save("test_files/escape_auto.docx")
