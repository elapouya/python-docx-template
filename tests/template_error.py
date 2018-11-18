from docxtpl import DocxTemplate, RichText
from jinja2.exceptions import TemplateError
import six

six.print_('=' * 80)
six.print_("Generating template error for testing (so it is safe to ignore) :")
six.print_('.' * 80)
try:
    tpl = DocxTemplate('templates/template_error_tpl.docx')
    tpl.render({
        'test_variable' : 'test variable value'
    })
except TemplateError as the_error:
    six.print_(six.text_type(the_error))
    if hasattr(the_error, 'docx_context'):
        six.print_("Context:")
        for line in the_error.docx_context:
            six.print_(line)
tpl.save('output/template_error.docx')
six.print_('.' * 80)
six.print_(" End of TemplateError Test ")
six.print_('=' * 80)
