# -*- coding: utf-8 -*-
"""
Created : 2015-03-12

@author: Eric Lapouyade
"""

__version__ = "0.20.2"

# flake8: noqa
from .inline_image import InlineImage
from .listing import Listing
from .richtext import RichText, R, RichTextParagraph, RP
from .template import DocxTemplate

try:
    from .subdoc import Subdoc
except ImportError:
    pass
