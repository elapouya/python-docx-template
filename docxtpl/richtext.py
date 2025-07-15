"""
Created : 2021-07-30

@author: Eric Lapouyade
"""

try:
    from html import escape
except ImportError:
    # cgi.escape is deprecated in python 3.7
    from cgi import escape


class RichText:
    """class to generate Rich Text when using templates variables

    This is much faster than using Subdoc class,
    but this only for texts INSIDE an existing paragraph.
    """

    def __init__(self, text=None, **text_prop):
        self.xml = ""
        if text:
            self.add(text, **text_prop)

    def add(
        self,
        text,
        style=None,
        color=None,
        highlight=None,
        size=None,
        subscript=None,
        superscript=None,
        bold=False,
        italic=False,
        underline=False,
        strike=False,
        font=None,
        url_id=None,
        rtl=False,
        lang=None,
    ):
        # If a RichText is added
        if isinstance(text, RichText):
            self.xml += text.xml
            return

        # # If nothing to add : just return
        # if text is None or text == "":
        #     return

        # If not a string : cast to string (ex: int, dict etc...)
        if not isinstance(text, (str, bytes)):
            text = str(text)
        if not isinstance(text, str):
            text = text.decode("utf-8", errors="ignore")
        text = escape(text)

        prop = ""

        if style:
            prop += f'<w:rStyle w:val="{style}"/>'
        if color:
            if color[0] == "#":
                color = color[1:]
            prop += f'<w:color w:val="{color}"/>'
        if highlight:
            if highlight[0] == "#":
                highlight = highlight[1:]
            prop += f'<w:shd w:fill="{highlight}"/>'
        if size:
            prop += f'<w:sz w:val="{size}"/>'
            prop += f'<w:szCs w:val="{size}"/>'
        if subscript:
            prop += '<w:vertAlign w:val="subscript"/>'
        if superscript:
            prop += '<w:vertAlign w:val="superscript"/>'
        if bold:
            prop += "<w:b/>"
            if rtl:
                prop += "<w:bCs/>"
        if italic:
            prop += "<w:i/>"
            if rtl:
                prop += "<w:iCs/>"
        if underline:
            if underline not in [
                "single",
                "double",
                "thick",
                "dotted",
                "dash",
                "dotDash",
                "dotDotDash",
                "wave",
            ]:
                underline = "single"
            prop += f'<w:u w:val="{underline}"/>'
        if strike:
            prop += "<w:strike/>"
        if font:
            regional_font = ""
            if ":" in font:
                region, font = font.split(":", 1)
                regional_font = f' w:{region}="{font}"'
            prop += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"{regional_font}/>'
        if rtl:
            prop += '<w:rtl w:val="true"/>'
        if lang:
            prop += f'<w:lang w:val="{lang}"/>'
        xml = "<w:r>"
        if prop:
            xml += f"<w:rPr>{prop}</w:rPr>"
        xml += f'<w:t xml:space="preserve">{text}</w:t></w:r>'
        if url_id:
            xml = f'<w:hyperlink r:id="{url_id}" w:tgtFrame="_blank">{xml}</w:hyperlink>'
        self.xml += xml

    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml

    def __html__(self):
        return self.xml


class RichTextParagraph:
    """class to generate Rich Text Paragraphs when using templates variables

    This is much faster than using Subdoc class,
    but this only for texts OUTSIDE an existing paragraph.
    """

    def __init__(self, text=None, **text_prop):
        self.xml = ""
        if text:
            self.add(text, **text_prop)

    def add(
        self,
        text,
        parastyle=None,
    ):
        # If a RichText is added
        if not isinstance(text, RichText):
            text = RichText(text)

        prop = ""
        if parastyle:
            prop += f'<w:pStyle w:val="{parastyle}"/>'

        xml = "<w:p>"
        if prop:
            xml += f"<w:pPr>{prop}</w:pPr>"
        xml += text.xml
        xml += "</w:p>"
        self.xml += xml

    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml

    def __html__(self):
        return self.xml


R = RichText
RP = RichTextParagraph
