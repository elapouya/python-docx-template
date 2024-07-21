# -*- coding: utf-8 -*-
"""
Created : 2019-05-22

@author: Eric Dufresne
"""

from docxtpl import DocxTemplate
import io

DEST_FILE = "output/header_footer_image_file_obj.docx"
DEST_FILE2 = "output/header_footer_image_file_obj2.docx"

tpl = DocxTemplate("templates/header_footer_image_tpl.docx")

context = {
    "mycompany": "The World Wide company",
}

dummy_pic = io.BytesIO(open("templates/dummy_pic_for_header.png", "rb").read())
new_image = io.BytesIO(open("templates/python.png", "rb").read())
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
