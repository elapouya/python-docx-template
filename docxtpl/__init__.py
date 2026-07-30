# -*- coding: utf-8 -*-
"""
Created : 2015-03-12

@author: Eric Lapouyade
"""

__version__ = "0.20.2"
__all__ = (
    "RP",
    "DocxTemplate",
    "InlineImage",
    "Listing",
    "R",
    "RichText",
    "RichTextParagraph",
)

from .inline_image import InlineImage
from .listing import Listing
from .richtext import RP, R, RichText, RichTextParagraph
from .template import DocxTemplate

try:
    from .subdoc import Subdoc

    __all__ += ("Subdoc",)
except ImportError:
    pass
