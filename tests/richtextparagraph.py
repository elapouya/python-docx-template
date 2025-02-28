"""
Created : 2025-02-28

@author: Hannah Imrie
"""

from docxtpl import DocxTemplate, RichText, RichTextParagraph

tpl = DocxTemplate("templates/richtext_paragraph_tpl.docx")

rtp = RichTextParagraph()
rt = RichText()

rtp.add("The rich text paragraph function allows paragraph styles to be added to text",parastyle="myrichparastyle")

rtp.add("This allows for the use of") 
rtp.add("bullet\apoints.", parastyle="SquareBullet")

rt.add("This works with ")
rt.add("Rich ", bold=True)
rt.add("Text ", italic=True)
rt.add("Strings", underline="single")
rt.add(" too.")

rtp.add(rt, parastyle="SquareBullet")

context = {
    "example": rtp,
}

tpl.render(context)
tpl.save("output/richtext_paragraph.docx")