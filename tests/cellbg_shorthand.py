from lxml import etree

from docxtpl import DocxTemplate


WORD_NAMESPACE = (
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
)
SOURCE_XML = f"""
<w:tc xmlns:w="{WORD_NAMESPACE}">
  <w:tcPr />
  <w:p><w:r><w:t>{{% cellbg background %}}content</w:t></w:r></w:p>
</w:tc>
"""

patched_xml = DocxTemplate(None).patch_xml(SOURCE_XML)
cell = etree.fromstring(patched_xml.encode())
namespaces = {"w": WORD_NAMESPACE}

assert len(cell.xpath("./w:tcPr/w:shd", namespaces=namespaces)) == 1
assert len(cell.xpath("./w:shd", namespaces=namespaces)) == 0
