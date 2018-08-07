from docxtpl import DocxTemplate, RichText
from jinja2.exceptions import TemplateError

try:
    tpl = DocxTemplate('test_files/template_error_tpl.docx')
    tpl.render({
        'test_variable' : 'test variable value'
    })
except TemplateError as the_error:
    print unicode(the_error)
    if hasattr(the_error, 'docx_context'):
        print "Context:"
        for line in the_error.docx_context:
            print line
tpl.save('test_files/template_error.docx')
