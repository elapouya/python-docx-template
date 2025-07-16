"""
Created : 2019-05-22

@author: Eric Dufresne
"""

import io
from pathlib import Path

from docxtpl import DocxTemplate

DEST_FILE = "output/header_footer_image_file_obj.docx"
DEST_FILE2 = "output/header_footer_image_file_obj2.docx"

tpl = DocxTemplate("templates/header_footer_image_tpl.docx")

context = {
    "mycompany": "The World Wide company",
}

dummy_pic = io.BytesIO(Path("templates/dummy_pic_for_header.png").read_bytes())
new_image = io.BytesIO(Path("templates/python.png").read_bytes())
tpl.replace_media(dummy_pic, new_image)
tpl.render(context)
tpl.save(DEST_FILE)

tpl = DocxTemplate("templates/header_footer_image_tpl.docx")
dummy_pic.seek(0)
new_image.seek(0)
tpl.replace_media(dummy_pic, new_image)
tpl.render(context)

file_obj = io.BytesIO()
tpl.save(file_obj)
file_obj.seek(0)
with open(DEST_FILE2, "wb") as f:
    f.write(file_obj.read())
