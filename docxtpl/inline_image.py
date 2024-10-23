# -*- coding: utf-8 -*-
"""
Created : 2021-07-30

@author: Eric Lapouyade
"""
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn


class InlineImage(object):
    """Class to generate an inline image

    This is much faster than using Subdoc class.
    """

    tpl = None
    image_descriptor = None
    width = None
    height = None
    anchor = None
    spacing_left = None
    spacing_right = None
    spacing_top = None
    spacing_bottom = None

    def __init__(self, tpl, image_descriptor, width=None, height=None, anchor=None, spacing_left=None, spacing_right=None, spacing_top=None, spacing_bottom=None):
        self.tpl, self.image_descriptor = tpl, image_descriptor
        self.width, self.height = width, height
        self.anchor = anchor
        self.spacing_left, self.spacing_right, self.spacing_top, self.spacing_bottom = spacing_left, spacing_right, spacing_top, spacing_bottom


    def _add_hyperlink(self, run, url, part):
        # Create a relationship for the hyperlink
        r_id = part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )

        # Find the <wp:docPr> and <pic:cNvPr> element
        docPr = run.xpath(".//wp:docPr")[0]
        cNvPr = run.xpath(".//pic:cNvPr")[0]

        # Create the <a:hlinkClick> element
        hlinkClick1 = OxmlElement("a:hlinkClick")
        hlinkClick1.set(qn("r:id"), r_id)
        hlinkClick2 = OxmlElement("a:hlinkClick")
        hlinkClick2.set(qn("r:id"), r_id)

        # Insert the <a:hlinkClick> element right after the <wp:docPr> element
        docPr.append(hlinkClick1)
        cNvPr.append(hlinkClick2)

        return run

    def _insert_image(self):
        pic = self.tpl.current_rendering_part.new_pic_inline(
            self.image_descriptor,
            self.width,
            self.height,
        )
        if self.spacing_left is not None or self.spacing_right is not None or self.spacing_top is not None or self.spacing_bottom is not None:
            if self.spacing_top is not None:
                pic.set('distT', str(self.spacing_top))
            if self.spacing_bottom is not None:
                pic.set('distB', str(self.spacing_bottom))
            if self.spacing_left is not None:
                pic.set('distL', str(self.spacing_left))
            if self.spacing_right is not None:
                pic.set('distR', str(self.spacing_right))
        if self.anchor:
            if pic.xpath(".//a:blip"):
                hyperlink = self._add_hyperlink(
                    pic, self.anchor, self.tpl.current_rendering_part
                )
                pic = hyperlink
        return (
            "</w:t></w:r><w:r><w:drawing>%s</w:drawing></w:r><w:r>"
            '<w:t xml:space="preserve">' % pic.xml
        )

    def __unicode__(self):
        return self._insert_image()

    def __str__(self):
        return self._insert_image()

    def __html__(self):
        return self._insert_image()
