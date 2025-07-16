"""
Created : 2025-02-28

@author: Hannah Imrie
"""

from docxtpl import DocxTemplate, RichText, RichTextParagraph

tpl = DocxTemplate("templates/richtext_paragraph_tpl.docx")

rtp = RichTextParagraph()
rt = RichText()

rtp.add(
    "The rich text paragraph function allows paragraph styles to be added to text",
    parastyle="myrichparastyle",
)
rtp.add("Any built in paragraph style can be used", parastyle="IntenseQuote")
rtp.add("or you can add your own, unlocking all style options", parastyle="createdStyle")
rtp.add(
    "To use, just create a style in your template word doc with the formatting you want and call it in the code.",
    parastyle="normal",
)

rtp.add("This allows for the use of")
rtp.add("custom bullet\apoints", parastyle="SquareBullet")
rtp.add("Numbered Bullet Points", parastyle="BasicNumbered")
rtp.add("and Alpha Bullet Points.", parastyle="alphaBracketNumbering")
rtp.add("You can", parastyle="normal")
rtp.add("set the", parastyle="centerAlign")
rtp.add("text alignment", parastyle="rightAlign")
rtp.add(
    "as well as the spacing between lines of text. Like this for example, "
    "this text has very tight spacing between the lines.\aIt also has no space between "
    "paragraphs of the same style.",
    parastyle="TightLineSpacing",
)
rtp.add(
    "Unlike this one, which has extra large spacing between lines for when you want to "
    "space things out a bit or just write a little less.",
    parastyle="WideLineSpacing",
)
rtp.add("You can also set the background colour of a line.", parastyle="LineShadingGreen")

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
