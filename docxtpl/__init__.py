# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

__version__ = '0.1.11'

from lxml import etree
from docx import Document
from jinja2 import Template
import zipfile
from cgi import escape
from StringIO import StringIO
import re
import six
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class DocxTemplate(object):
    """ Class for managing docx files as they were jinja2 templates """
    def __init__(self, file):
        self.document = zipfile.ZipFile(file, 'r')

        self.document_part = self.document.read('word/document.xml').decode().encode("utf-8")
        self.header = self.document.read('word/header1.xml').decode().encode("utf-8") \
            if 'word/header1.xml' in self.document.namelist() else None
        self.footer = self.document.read('word/footer1.xml').decode().encode("utf-8") \
            if 'word/footer1.xml' in self.document.namelist() else None

    def __getattr__(self, name):
        return getattr(self.docx, name)

    def save(self, write_file):
        temp_buffer = StringIO()
        with zipfile.ZipFile(temp_buffer, 'a', compression=zipfile.ZIP_DEFLATED) as document:
            for item in self.document.infolist():
                if item.filename != 'word/document.xml' and item.filename != 'word/header1.xml' \
                and item.filename != 'word/footer1.xml':
                    document.writestr(item.filename, (self.document.read(item.filename)))

            document.writestr('word/document.xml', str(self.document_part))
            if self.header:
                document.writestr('word/header1.xml', str(self.header))
            if self.footer:
                document.writestr('word/footer1.xml', str(self.footer))

        write_file.write(temp_buffer.getvalue())
        self.document.close()

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

        # manage table cell background color
        def cellbg(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(r'<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>','',cell_xml,flags=re.DOTALL)
            cell_xml = re.sub(r'<w:shd[^/]*/>','', cell_xml, count=1)
            return re.sub(r'(<w:tcPr[^>]*>)',r'\1<w:shd w:val="clear" w:color="auto" w:fill="{{%s}}"/>' % m.group(2), cell_xml)
        src_xml = re.sub(r'(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+([^%]*)\s*%}(.*?</w:tc>)',cellbg,src_xml,flags=re.DOTALL)

        for y in ['tr', 'p', 'r']:
            # replace into xml code the row/paragraph/run containing {%y xxx %} or {{y xxx}} template tag
            # by {% xxx %} or {{ xx }} without any surronding xml tags :
            # This is mandatory to have jinja2 generating correct xml code
            pat = r'<w:%(y)s[ >](?:(?!<w:%(y)s[ >]).)*({%%|{{)%(y)s ([^}%%]*(?:%%}|}})).*?</w:%(y)s>' % {'y':y}
            src_xml = re.sub(pat, r'\1 \2',src_xml,flags=re.DOTALL)

        def clean_tags(m):
            return m.group(0).replace(r"&#8216;","'").replace('&lt;','<').replace('&gt;','>')
        src_xml = re.sub(r'(?<=\{[\{%])([^\}%]*)(?=[\}%]})',clean_tags,src_xml)

        return src_xml

    def render_xml(self,src_xml,context,jinja_env=None):
        src_xml = src_xml.encode('ascii', 'ignore')
        if jinja_env:
            template = jinja_env.from_string(src_xml)
        else:
            template = Template(src_xml)
        dst_xml = template.render(context)
        dst_xml = dst_xml.replace('{_{','{{').replace('}_}','}}').replace('{_%','{%').replace('%_}','%}')
        return dst_xml

    def build_xml(self, context, xml, jinja_env=None):
        xml = self.patch_xml(xml)
        xml = self.render_xml(xml, context, jinja_env)
        return xml


    def map_xml(self,xml):
        root = self.docx._element
        body = root.body
        root.replace(body,etree.fromstring(xml))

    def render(self,context,jinja_env=None):
        self.document_part = self.build_xml(context, self.document_part, jinja_env)
        if self.header:
            self.header = self.build_xml(context, self.header, jinja_env)
        if self.footer:
            self.footer = self.build_xml(context, self.footer, jinja_env)


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
        text = escape(text).replace('\n','<w:br/>')

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
        self.xml += u'<w:t xml:space="preserve">%s</w:t></w:r>\n' % text

    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml