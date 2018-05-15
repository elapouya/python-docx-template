from docxtpl import *

tpl = DocxTemplate('test_files/unescape_jinja2_tags_tpl.docx')
tpl.render({})
tpl.save('test_files/unescape_jinja2_tags.docx')
