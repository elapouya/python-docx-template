# -*- coding: utf-8 -*-
"""
Created : 2021-07-30

@author: Eric Lapouyade
"""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.image.image import Image


class InlineImage(object):
    """Class to generate an inline image

    This is much faster than using Subdoc class.
    """

    tpl = None
    image_descriptor = None
    width = None
    height = None
    anchor = None
    title = None
    descr = None

    def __init__(
        self,
        tpl,
        image_descriptor,
        width=None,
        height=None,
        anchor=None,
        title=None,
        descr=None,
    ):
        self.tpl, self.image_descriptor = tpl, image_descriptor
        self.width, self.height = width, height
        self.anchor = anchor
        self.title = str(title) if title is not None else None
        self.descr = str(descr) if descr is not None else None

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

    def _set_doc_properties(self, run):
        docPr = run.xpath(".//wp:docPr")[0]
        cNvPr = run.xpath(".//pic:cNvPr")[0]

        for elt in (docPr, cNvPr):
            if self.title is not None:
                elt.set("title", self.title)
            if self.descr is not None:
                elt.set("descr", self.descr)

    def _insert_image(self):
        package = self.tpl.current_rendering_part.package
        if any(
            not hasattr(image_part.image, "scaled_dimensions")
            for image_part in package.image_parts
        ):
            image_part = package.get_or_add_image_part(self.image_descriptor)
            if not hasattr(image_part.image, "scaled_dimensions"):
                # docxcompose caches a lightweight wrapper for copied images.
                # Replace only the matching wrapper, leaving unrelated image
                # parts and exceptions untouched.
                image_part._image = Image.from_blob(image_part.blob)

        pic = self.tpl.current_rendering_part.new_pic_inline(
            self.image_descriptor,
            self.width,
            self.height,
        )
        if self.title is not None or self.descr is not None:
            self._set_doc_properties(pic)
        if self.anchor:
            if pic.xpath(".//a:blip"):
                pic = self._add_hyperlink(
                    pic, self.anchor, self.tpl.current_rendering_part
                )
        pic = pic.xml

        return (
            "</w:t></w:r><w:r><w:drawing>%s</w:drawing></w:r><w:r>"
            '<w:t xml:space="preserve">' % pic
        )

    def __unicode__(self):
        return self._insert_image()

    def __str__(self):
        return self._insert_image()

    def __html__(self):
        return self._insert_image()
