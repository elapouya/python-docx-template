import io
import zipfile

from docx import Document
from lxml import etree

from docxtpl import DocxTemplate


main_template = Document()
main_template.add_paragraph("{{p subdoc }}")
main_template_stream = io.BytesIO()
main_template.save(main_template_stream)
main_template_stream.seek(0)

template = DocxTemplate(main_template_stream)
subdoc = template.new_subdoc("templates/pandoc_subdoc.docx")
template.render({"subdoc": subdoc})
output_stream = io.BytesIO()
template.save(output_stream)

output_stream.seek(0)
document = Document(output_stream)
assert len(document.inline_shapes) == 1

output_stream.seek(0)
with zipfile.ZipFile(output_stream) as archive:
    document_xml = archive.read("word/document.xml")

root = etree.fromstring(document_xml)
assert root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}graphic") is not None
assert root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/picture}pic") is not None
