# -*- coding: utf-8 -*-
"""
Created : 2017-09-09

@author: Eric Lapouyade
"""

from docxtpl import DocxTemplate

# rendering the "dynamic embedded docx":
embedded_docx_tpl = DocxTemplate("templates/embedded_embedded_docx_tpl.docx")
context = {
    "name": "John Doe",
}
embedded_docx_tpl.render(context)
embedded_docx_tpl.save("output/embedded_embedded_docx.docx")


# rendering the main document :
tpl = DocxTemplate("templates/embedded_main_tpl.docx")

context = {
    "name": "John Doe",
}

tpl.replace_embedded(
    "templates/embedded_dummy.docx", "templates/embedded_static_docx.docx"
)
tpl.replace_embedded(
    "templates/embedded_dummy2.docx", "output/embedded_embedded_docx.docx"
)

# The zipname is the one you can find when you open docx with WinZip, 7zip (Windows)
# or unzip -l (Linux). The zipname starts with "word/embeddings/".
# Note that the file is renamed by MSWord, so you have to guess a little bit...
tpl.replace_zipname(
    "word/embeddings/Feuille_Microsoft_Office_Excel3.xlsx", "templates/real_Excel.xlsx"
)
tpl.replace_zipname(
    "word/embeddings/Pr_sentation_Microsoft_Office_PowerPoint4.pptx",
    "templates/real_PowerPoint.pptx",
)

tpl.render(context)
tpl.save("output/embedded.docx")
