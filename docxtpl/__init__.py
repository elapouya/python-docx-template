# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

__version__ = '0.3.2'

from lxml import etree
from docx import Document
from jinja2 import Template
from cgi import escape
import re
import six

class DocxTemplate(object):
    """ Class for managing docx files as they were jinja2 templates """

    HEADER_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    FOOTER_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

    def __init__(self, docx):
        self.docx = Document(docx)

    def __getattr__(self, name):
        return getattr(self.docx, name)

    def get_docx(self):
        return self.docx

    def get_xml(self):
        # Be careful : pretty_print MUST be set to False, otherwise patch_xml() won't work properly
        return etree.tostring(self.docx._element.body, encoding='unicode', pretty_print=False)

    def write_xml(self,filename):
        with open(filename,'w') as fh:
            fh.write(self.get_xml())

    def patch_xml(self,src_xml):
        # strip all xml tags inside {% %} and {{ }} that MS word can insert into xml source
        src_xml = re.sub(r'(?<={)(<[^>]*>)+(?=[\{%])|(?<=[%\}])(<[^>]*>)+(?=\})','',src_xml,flags=re.DOTALL)
        def striptags(m):
            return re.sub('</w:t>.*?(<w:t>|<w:t [^>]*>)','',m.group(0),flags=re.DOTALL)
        src_xml = re.sub(r'{%(?:(?!%}).)*|{{(?:(?!}}).)*',striptags,src_xml,flags=re.DOTALL)

        # manage table cell colspan
        def colspan(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(r'<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>','',cell_xml,flags=re.DOTALL)
            cell_xml = re.sub(r'<w:gridSpan[^/]*/>','', cell_xml, count=1)
            return re.sub(r'(<w:tcPr[^>]*>)',r'\1<w:gridSpan w:val="{{%s}}"/>' % m.group(2), cell_xml)
        src_xml = re.sub(r'(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*colspan\s+([^%]*)\s*%}(.*?</w:tc>)',colspan,src_xml,flags=re.DOTALL)

        # manage table cell background color
        def cellbg(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(r'<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>','',cell_xml,flags=re.DOTALL)
            cell_xml = re.sub(r'<w:shd[^/]*/>','', cell_xml, count=1)
            return re.sub(r'(<w:tcPr[^>]*>)',r'\1<w:shd w:val="clear" w:color="auto" w:fill="{{%s}}"/>' % m.group(2), cell_xml)
        src_xml = re.sub(r'(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+([^%]*)\s*%}(.*?</w:tc>)',cellbg,src_xml,flags=re.DOTALL)

        for y in ['tr', 'p', 'r']:
            # replace into xml code the row/paragraph/run containing {%y xxx %} or {{y xxx}} template tag
            # by {% xxx %} or {{ xx }} without any surronding <w:y> tags :
            # This is mandatory to have jinja2 generating correct xml code
            pat = r'<w:%(y)s[ >](?:(?!<w:%(y)s[ >]).)*({%%|{{)%(y)s ([^}%%]*(?:%%}|}})).*?</w:%(y)s>' % {'y':y}
            src_xml = re.sub(pat, r'\1 \2',src_xml,flags=re.DOTALL)

        def clean_tags(m):
            return m.group(0).replace(r"&#8216;","'").replace('&lt;','<').replace('&gt;','>')
        src_xml = re.sub(r'(?<=\{[\{%])([^\}%]*)(?=[\}%]})',clean_tags,src_xml)

        return src_xml

    def render_xml(self,src_xml,context,jinja_env=None):
        if jinja_env:
            template = jinja_env.from_string(src_xml)
        else:
            template = Template(src_xml)
        dst_xml = template.render(context)
        dst_xml = dst_xml.replace('{_{','{{').replace('}_}','}}').replace('{_%','{%').replace('%_}','%}')
        return dst_xml

    def build_xml(self,context,jinja_env=None):
        xml = self.get_xml()
        xml = self.patch_xml(xml)
        xml = self.render_xml(xml, context, jinja_env)
        return xml

    def map_xml(self,xml):
        root = self.docx._element
        body = root.body
        root.replace(body,etree.fromstring(xml))

    def get_headers_footers_xml(self, uri):
        for relKey, val in self.docx._part._rels.items():
            if val.reltype == uri:
                yield relKey, val._target._blob

    def get_headers_footers_encoding(self,xml):
        m = re.match(r'<\?xml[^\?]+\bencoding="([^"]+)"',xml,re.I)
        if m:
            return m.group(1)
        return 'utf-8'

    def build_headers_footers_xml(self,context, uri,jinja_env=None):
        for relKey, xml in self.get_headers_footers_xml(uri):
            if six.PY3:
                xml = xml.decode('utf-8')
            encoding = self.get_headers_footers_encoding(xml)
            xml = self.patch_xml(xml)
            if not six.PY3:
                xml = xml.decode(encoding)
            xml = self.render_xml(xml, context, jinja_env)
            yield relKey, xml.encode(encoding)

    def map_headers_footers_xml(self, relKey, xml):
        self.docx._part._rels[relKey]._target._blob = xml

    def render(self,context,jinja_env=None):
        # Body
        xml = self.build_xml(context,jinja_env)
        self.map_xml(xml)

        # Headers
        for relKey, xml in self.build_headers_footers_xml(context, self.HEADER_URI, jinja_env):
            self.map_headers_footers_xml(relKey, xml)

        # Footers
        for relKey, xml in self.build_headers_footers_xml(context, self.FOOTER_URI, jinja_env):
            self.map_headers_footers_xml(relKey, xml)

    def new_subdoc(self):
        return Subdoc(self)

class Subdoc(object):
    """ Class for subdocument to insert into master document """
    def __init__(self, tpl):
        self.tpl = tpl
        self.docx = tpl.get_docx()
        self.subdocx = Document()
        self.subdocx._part = self.docx._part

    def __getattr__(self, name):
        return getattr(self.subdocx, name)

    def _get_xml(self):
        if self.subdocx._element.body.sectPr is not None:
            self.subdocx._element.body.remove(self.subdocx._element.body.sectPr)
        xml = re.sub(r'</?w:body[^>]*>','',etree.tostring(self.subdocx._element.body, encoding='unicode', pretty_print=False))
        return xml

    def __unicode__(self):
        return self._get_xml()

    def __str__(self):
        return self._get_xml()


class RichText(object):
    """ class to generate Rich Text when using templates variables

    This is much faster than using Subdoc class, but this only for texts INSIDE an existing paragraph.
    """
    def __init__(self, text=None, **text_prop):
        self.xml = ''
        if text:
            self.add(text, **text_prop)

    def add(self, text, style=None,
                        color=None,
                        highlight=None,
                        size=None,
                        bold=False,
                        italic=False,
                        underline=False,
                        strike=False):


        if not isinstance(text, six.text_type):
            text = text.decode('utf-8',errors='ignore')
        text = escape(text).replace('\n','</w:t><w:br/><w:t>').replace('\a','</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">')

        prop = u''

        if style:
            prop += u'<w:rStyle w:val="%s"/>' % style
        if color:
            if color[0] == '#':
                color = color[1:]
            prop += u'<w:color w:val="%s"/>' % color
        if highlight:
            if highlight[0] == '#':
                highlight = highlight[1:]
            prop += u'<w:highlight w:val="%s"/>' % highlight
        if size:
            prop += u'<w:sz w:val="%s"/>' % size
            prop += u'<w:szCs w:val="%s"/>' % size
        if bold:
            prop += u'<w:b/>'
        if italic:
            prop += u'<w:i/>'
        if underline:
            if underline not in ['single','double']:
                underline = 'single'
            prop += u'<w:u w:val="%s"/>' % underline
        if strike:
            prop += u'<w:strike/>'

        self.xml += u'<w:r>'
        if prop:
            self.xml += u'<w:rPr>%s</w:rPr>' % prop
        self.xml += u'<w:t xml:space="preserve">%s</w:t></w:r>' % text

    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml

class InlineImage(object):
    """ class to generate an inline image

    This is much faster than using Subdoc class.
    """
    def __init__(self, tpl, file, width=None, height=None):
        image_xml = tpl.docx._part.new_pic_inline(file, width, height).xml
        self.xml = '</w:t></w:r><w:r><w:drawing>%s</w:drawing></w:r><w:r><w:t xml:space="preserve">' % image_xml

    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml

R = RichText

