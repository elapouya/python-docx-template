# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

__version__ = '0.1.1'

from lxml import etree
from docx import Document
from jinja2 import Template
import copy
import re

class DocxTemplate(object):
    """ Class for managing docx files as they were jinja2 templates """
    def __init__(self, docx):
        self.docx = Document(docx)
    
    def __getattr__(self, name):        
        return getattr(self.docx, name)
    
    def build_xml(self,context):
        src_xml = etree.tostring(self.docx._element.body, pretty_print=True)
        
        with open('/tmp/docx1.xml','w') as fh:
            fh.write(src_xml)

        # strip all xml tags inside {% %} and {{ }}
        # that Microsoft word can insert into xml code for this part of the document
        def striptags(m):
            return re.sub('</w:t>.*?(<w:t>|<w:t [^>]*>)','',m.group(0),flags=re.DOTALL)
        src_xml = re.sub(r'{%(?:(?!%}).)*|{{(?:(?!}}).)*',striptags,src_xml,flags=re.DOTALL)
        
        # manage table cell background color
        def cellbg(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(r'<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>','',cell_xml,flags=re.DOTALL)
            return re.sub(r'(<w:shd[^/]*w:fill=")[^"]*("[^/]*/>)',r'\1{{ %s}}\2' % m.group(2), cell_xml)
        src_xml = re.sub(r'(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+([^%]*)\s*%}(.*?</w:tc>)',cellbg,src_xml,flags=re.DOTALL)
        
        # replace xml code corresponding to the paragraph containing {{{ xxx }}} by {{ xxx }}
        src_xml = re.sub(r'<w:p[ >](?:(?!<w:p[ >]).)*{{{([^}]*)}}}.*?</w:p>',r'{{\1}}',src_xml,flags=re.DOTALL)
        
        # replace xml code corresponding to the row containing {% tr-xxx template tag by {% xxx template tag itself
        src_xml = re.sub(r'<w:tr[ >](?:(?!<w:tr[ >]).)*{%\s*tr-([^%]*%}).*?</w:tr>',r'{% \1',src_xml,flags=re.DOTALL)
        
        # replace xml code corresponding to the paragraph containing {% p-xxx template tag by {% xxx template tag itself
        src_xml = re.sub(r'<w:p[ >](?:(?!<w:p[ >]).)*{%\s*p-([^%]*%}).*?</w:p>',r'{% \1',src_xml,flags=re.DOTALL)
        
        with open('/tmp/docx2.xml','w') as fh:
            fh.write(src_xml)
        
        template = Template(src_xml)
        dst_xml = template.render(context)
        
        return dst_xml
        
    def map_xml(self,xml):
        root = self.docx._element
        body = root.body
        root.replace(body,etree.fromstring(xml))

    def render(self,context):
        xml = self.build_xml(context)
        self.map_xml(xml)

class Subdoc(object):
    """ Class for subdocumentation insertion into master document """
    pass

class Tpldoc(object):
    """ class to build documenation to be passed into template variables """
    pass