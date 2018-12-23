"""
@author: Max Podolskii
"""

import os
from unicodedata import name

from six import iteritems, text_type

from docxtpl import DocxTemplate


XML_RESERVED = """<"&'>"""

tpl = DocxTemplate('templates/escape_tpl_auto.docx')

context = {'nested_dict': {name(text_type(c)): c for c in XML_RESERVED},
           'autoescape': 'Escaped "str & ing"!',
           'autoescape_unicode': u'This is an escaped <unicode> example \u4f60 & \u6211',
           'iteritems': iteritems,
           }

tpl.render(context, autoescape=True)

OUTPUT = 'output'
if not os.path.exists(OUTPUT):
    os.makedirs(OUTPUT)
tpl.save(OUTPUT + '/escape_auto.docx')
