from docxtpl import DocxTemplate
from jinja2.exceptions import TemplateError

print("=" * 80)
print("Generating template error for testing (so it is safe to ignore) :")
print("." * 80)
try:
    tpl = DocxTemplate("templates/template_error_tpl.docx")
    tpl.render({"test_variable": "test variable value"})
except TemplateError as the_error:
    print(str(the_error))
    if hasattr(the_error, "docx_context"):
        print("Context:")
        for line in the_error.docx_context:
            print(line)
tpl.save("output/template_error.docx")
print("." * 80)
print(" End of TemplateError Test ")
print("=" * 80)
