from zipfile import ZipFile

from lxml import etree

from docxtpl import DocxTemplate

doctemplate = r"templates/doc_properties_tpl.docx"
output = "output/core_properties.docx"

tpl = DocxTemplate(doctemplate)
tpl.render({"test": "HelloWorld"})
tpl.docx.core_properties.keywords = "rendered by docxtpl"
tpl.save(output)

with ZipFile(output) as archive:
    core_properties = etree.fromstring(archive.read("docProps/core.xml"))

keywords = [
    element
    for element in core_properties
    if etree.QName(element).localname == "keywords"
]
assert len(keywords) == 1
assert etree.QName(keywords[0]).namespace == (
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
)
assert keywords[0].text == "rendered by docxtpl"
