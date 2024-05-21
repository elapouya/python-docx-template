# -*- coding: utf-8 -*-
"""
Created : 2021-07-30

@author: Eric Lapouyade
"""

from docx import Document
from docxcompose.composer import Composer
from lxml import etree
import re

class Subdoc(object):
    """ Class for subdocument to insert into master document """
    def __init__(self, tpl, docpath=None):
        self.tpl = tpl
        self.docx = tpl.get_docx()
        self.subdocx = Document(docpath)
        if docpath:
            compose = Composer(self.docx)
            compose.append(self.subdocx)
        else:
            self.subdocx._part = self.docx._part

    def __getattr__(self, name):
        return getattr(self.subdocx, name)

    def _get_xml(self):
        if self.subdocx.element.body.sectPr is not None:
            self.subdocx.element.body.remove(self.subdocx.element.body.sectPr)
        xml = re.sub(r'</?w:body[^>]*>', '', etree.tostring(
            self.subdocx.element.body, encoding='unicode', pretty_print=False))
        return xml

    def __unicode__(self):
        return self._get_xml()

    def __str__(self):
        return self._get_xml()

    def __html__(self):
        return self._get_xml()
