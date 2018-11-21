# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

__version__ = '0.5.11'

from lxml import etree
from docx import Document
from docx.opc.oxml import serialize_part_xml, parse_xml
import docx.oxml.ns
from docx.opc.constants import RELATIONSHIP_TYPE as REL_TYPE
from jinja2 import Template
from jinja2.exceptions import TemplateError
try:
    from html import escape, unescape
except ImportError:
    # cgi.escape is deprecated in python 3.7
    # import escape and unescape methods for Python 2.7
    from cgi import escape
    import HTMLParser
    unescape = HTMLParser.HTMLParser().unescape
import re
import six
import binascii
import os
import zipfile

NEWLINE_XML = '</w:t><w:br/><w:t xml:space="preserve">'
NEWPARAGRAPH_XML = '</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">'
TAB_XML = '</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">'

class DocxTemplate(object):
    """ Class for managing docx files as they were jinja2 templates """

    HEADER_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    FOOTER_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

    def __init__(self, docx):
        self.docx = Document(docx)
        self.crc_to_new_media = {}
        self.crc_to_new_embedded = {}
        self.pic_to_replace = {}
        self.pic_map = {}

    def __getattr__(self, name):
        return getattr(self.docx, name)

    def xml_to_string(self, xml, encoding='unicode'):
        # Be careful : pretty_print MUST be set to False, otherwise patch_xml()
        # won't work properly
        return etree.tostring(xml, encoding='unicode', pretty_print=False)

    def get_docx(self):
        return self.docx

    def get_xml(self):
        return self.xml_to_string(self.docx._element.body)

    def write_xml(self,filename):
        with open(filename,'w') as fh:
            fh.write(self.get_xml())

    def patch_xml(self,src_xml):
        # strip all xml tags inside {% %} and {{ }} that MS word can insert
        # into xml source also unescape html entities
        src_xml = re.sub(r'(?<={)(<[^>]*>)+(?=[\{%])|(?<=[%\}])(<[^>]*>)+(?=\})','',
                         src_xml,flags=re.DOTALL)
        def striptags(m):
            return re.sub('</w:t>.*?(<w:t>|<w:t [^>]*>)','',
                          m.group(0),flags=re.DOTALL)
        src_xml = re.sub(r'{%(?:(?!%}).)*|{{(?:(?!}}).)*',striptags,
                         src_xml,flags=re.DOTALL)

        # manage table cell colspan
        def colspan(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(r'<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>',
                              '', cell_xml,flags=re.DOTALL)
            cell_xml = re.sub(r'<w:gridSpan[^/]*/>','', cell_xml, count=1)
            return re.sub(r'(<w:tcPr[^>]*>)',r'\1<w:gridSpan w:val="{{%s}}"/>'
                          % m.group(2), cell_xml )
        src_xml = re.sub(r'(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*colspan\s+([^%]*)\s*%}(.*?</w:tc>)',
                         colspan,src_xml,flags=re.DOTALL)

        # manage table cell background color
        def cellbg(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(r'<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>',
                              '',cell_xml,flags=re.DOTALL)
            cell_xml = re.sub(r'<w:shd[^/]*/>','', cell_xml, count=1)
            return re.sub(r'(<w:tcPr[^>]*>)',
                          r'\1<w:shd w:val="clear" w:color="auto" w:fill="{{%s}}"/>'
                          % m.group(2), cell_xml)
        src_xml = re.sub(r'(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+([^%]*)\s*%}(.*?</w:tc>)',
                         cellbg,src_xml,flags=re.DOTALL)

        # avoid {{r and {%r tags to strip MS xml tags too far
        src_xml = re.sub(r'({{r\s.*?}}|{%r\s.*?%})',
                         r'</w:t></w:r><w:r><w:t>\1</w:t></w:r><w:r><w:t>',
                         src_xml,flags=re.DOTALL)

        for y in ['tr', 'tc', 'p', 'r']:
            # replace into xml code the row/paragraph/run containing
            # {%y xxx %} or {{y xxx}} template tag
            # by {% xxx %} or {{ xx }} without any surronding <w:y> tags :
            # This is mandatory to have jinja2 generating correct xml code
            pat = r'<w:%(y)s[ >](?:(?!<w:%(y)s[ >]).)*({%%|{{)%(y)s ([^}%%]*(?:%%}|}})).*?</w:%(y)s>' % {'y':y}
            src_xml = re.sub(pat, r'\1 \2',src_xml,flags=re.DOTALL)

        # add vMerge
        # use {% vm %} to make this table cell and its copies be vertically merged within a {% for %}
        def v_merge_tc(m):
            def v_merge(m1):
                return (
                    '<w:vMerge w:val="{% if loop.first %}restart{% else %}continue{% endif %}"/>' +
                    m1.group(1) +  # Everything between ``</w:tcPr>`` and ``<w:t>``.
                    "{% if loop.first %}" +
                    m1.group(2) +  # Everything before ``{% vm %}``.
                    m1.group(3) +  # Everything after ``{% vm %}``.
                    "{% endif %}" +
                    m1.group(4)  # ``</w:t>``.
                )
            return re.sub(
                r'(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*vm\s*%})(.*?)(</w:t>)',
                v_merge,
                m.group(),  # Everything between ``</w:tc>`` and ``</w:tc>`` with ``{% vm %}`` inside.
                flags=re.DOTALL,
            )
        src_xml = re.sub(r'<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*vm\s*%}.*?</w:tc[ >]',
                         v_merge_tc, src_xml, flags=re.DOTALL)

        # Use ``{% hm %}`` to make table cell become horizontally merged within
        # a ``{% for %}``.
        def h_merge_tc(m):
            xml_to_patch = m.group()  # Everything between ``</w:tc>`` and ``</w:tc>`` with ``{% hm %}`` inside.

            def with_gridspan(m1):
                return (
                    m1.group(1) +  # ``w:gridSpan w:val="``.
                    '{{ ' + m1.group(2) + ' * loop.length }}' +  # Content of ``w:val``, multiplied by loop length.
                    m1.group(3)  # Closing quotation mark.
                )

            def without_gridspan(m2):
                return (
                    '<w:gridSpan w:val="{{ loop.length }}"/>' +
                    m2.group(1) +  # Everything between ``</w:tcPr>`` and ``<w:t>``.
                    m2.group(2) +  # Everything before ``{% hm %}``.
                    m2.group(3) +  # Everything after ``{% hm %}``.
                    m2.group(4)  # ``</w:t>``.
                )

            if re.search(r'w:gridSpan', xml_to_patch):
                # Simple case, there's already ``gridSpan``, multiply its value.

                xml = re.sub(
                    r'(w:gridSpan w:val=")(\d+)(")',
                    with_gridspan,
                    xml_to_patch,
                    flags=re.DOTALL,
                )
                xml = re.sub(
                    r'{%\s*hm\s*%}',
                    '',
                    xml,  # Patched xml.
                    flags=re.DOTALL,
                )
            else:
                # There're no ``gridSpan``, add one.
                xml = re.sub(
                    r'(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*hm\s*%})(.*?)(</w:t>)',
                    without_gridspan,
                    xml_to_patch,
                    flags=re.DOTALL,
                )

            # Discard every other cell generated in loop.
            return "{% if loop.first %}" + xml + "{% endif %}"

        src_xml = re.sub(r'<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*hm\s*%}.*?</w:tc[ >]',
                         h_merge_tc, src_xml, flags=re.DOTALL)

        def clean_tags(m):
            return ( m.group(0)
                     .replace(r"&#8216;","'")
                     .replace('&lt;','<')
                     .replace('&gt;','>')
                     .replace(u'“',u'"')
                     .replace(u'”',u'"')
                     .replace(u"‘",u"'")
                     .replace(u"’",u"'") )
        src_xml = re.sub(r'(?<=\{[\{%])(.*?)(?=[\}%]})',clean_tags,src_xml)

        return src_xml

    def render_xml(self,src_xml,context,jinja_env=None):
        src_xml = src_xml.replace(r'<w:p>', '\n<w:p>')
        try:
            if jinja_env:
                template = jinja_env.from_string(src_xml)
            else:
                template = Template(src_xml)
            dst_xml = template.render(context)
        except TemplateError as exc:
            if hasattr(exc, 'lineno') and exc.lineno is not None:
                line_number = max(exc.lineno - 4, 0)
                exc.docx_context = map(lambda x: re.sub(r'<[^>]+>', '', x),
                                       src_xml.splitlines()[line_number:(line_number + 7)])
            raise exc
        dst_xml = dst_xml.replace('\n<w:p>', '<w:p>')
        dst_xml = ( dst_xml
                    .replace('{_{','{{')
                    .replace('}_}','}}')
                    .replace('{_%','{%')
                    .replace('%_}','%}') )
        return dst_xml

    def build_xml(self,context,jinja_env=None):
        xml = self.get_xml()
        xml = self.patch_xml(xml)
        xml = self.render_xml(xml, context, jinja_env)
        return xml

    def map_tree(self, tree):
        root = self.docx._element
        body = root.body
        root.replace(body, tree)

    def get_headers_footers_xml(self, uri):
        for relKey, val in self.docx._part._rels.items():
            if val.reltype == uri:
                yield relKey, self.xml_to_string(parse_xml(val._target._blob))

    def get_headers_footers_encoding(self,xml):
        m = re.match(r'<\?xml[^\?]+\bencoding="([^"]+)"',xml,re.I)
        if m:
            return m.group(1)
        return 'utf-8'

    def build_headers_footers_xml(self,context, uri,jinja_env=None):
        for relKey, xml in self.get_headers_footers_xml(uri):
            encoding = self.get_headers_footers_encoding(xml)
            xml = self.patch_xml(xml)
            xml = self.render_xml(xml, context, jinja_env)
            yield relKey, xml.encode(encoding)

    def map_headers_footers_xml(self, relKey, xml):
        self.docx._part._rels[relKey]._target._blob = xml

    @staticmethod
    def escape_values(context):
        """Escape strings for an XML Word document
        which may contain <, >, &, ', and ".
        """
        def escape_recursively(d, identities):
            """Escape string values of the passed :dict: `d` in-place
            including nested dictionaries of any depth.
            """
            for k, v in six.iteritems(d):
                if isinstance(v, dict):
                    identity = id(v)
                    if identity not in identities:
                        identities.add(identity)
                        escape_recursively(v, identities)
                else:
                    # Avoid dict, Listing, InlineImage, RichText, etc. classes
                    #  by comparing `v` to `str`. Do not use try-except.
                    if isinstance(v, str):
                        # Unescape at first to avoid secondary escaping
                        d[k] = escape(unescape(v))

        # Avoid RecursionError (if back edges, i.e. cycles, exist)
        # by using a set of unique identities of iterated dictionaries.
        initial_identities = {id(context)}

        escape_recursively(context, initial_identities)

    def render(self, context, jinja_env=None, autoescape=False):
        if autoescape:
            self.escape_values(context)

        # Body
        xml_src = self.build_xml(context, jinja_env)

        # fix tables if needed
        tree = self.fix_tables(xml_src)

        self.map_tree(tree)

        # Headers
        headers = self.build_headers_footers_xml(context, self.HEADER_URI,
                                                 jinja_env)
        for relKey, xml in headers:
            self.map_headers_footers_xml(relKey, xml)

        # Footers
        footers = self.build_headers_footers_xml(context, self.FOOTER_URI,
                                                 jinja_env)
        for relKey, xml in footers:
            self.map_headers_footers_xml(relKey, xml)

    # using of TC tag in for cycle can cause that count of columns does not
    # correspond to real count of columns in row. This function is able to fix it.
    def fix_tables(self, xml):
        parser = etree.XMLParser(recover=True)
        tree = etree.fromstring(xml, parser=parser)
        # get namespace
        ns = '{' + tree.nsmap['w'] + '}'
        # walk trough xml and find table
        for t in tree.iter(ns+'tbl'):
            tblGrid = t.find(ns+'tblGrid')
            columns = tblGrid.findall(ns+'gridCol')
            to_add = 0
            # walk trough all rows and try to find if there is higher cell count
            for r in t.iter(ns+'tr'):
                cells = r.findall(ns+'tc')
                if (len(columns) + to_add) < len(cells):
                    to_add = len(cells) - len(columns)
            # is neccessary to add columns?
            if to_add > 0:
                # at first, calculate width of table according to columns
                # (we want to preserve it)
                width = 0.0
                new_average = None
                for c in columns:
                    if not c.get(ns+'w') == None:
                        width += float(c.get(ns+'w'))
                # try to keep proportion of table
                if width > 0:
                    old_average = width / len(columns)
                    new_average = width / (len(columns) + to_add)
                    # scale the old columns
                    for c in columns:
                        c.set(ns+'w', str(int(float(c.get(ns+'w')) *
                                              new_average/old_average)))
                    # add new columns
                    for i in range(to_add):
                        etree.SubElement(tblGrid, ns+'gridCol',
                                         {ns+'w': str(int(new_average))})
        return tree

    def new_subdoc(self,docpath=None):
        return Subdoc(self,docpath)

    @staticmethod
    def get_file_crc(filename):
        with open(filename, 'rb') as fh:
            buf = fh.read()
            crc = (binascii.crc32(buf) & 0xFFFFFFFF)
        return crc

    def replace_media(self,src_file,dst_file):
        """Replace one media by another one into a docx

        This has been done mainly because it is not possible to add images in
        docx header/footer.
        With this function, put a dummy picture in your header/footer,
        then specify it with its replacement in this function

        Syntax: tpl.replace_media('dummy_media_to_replace.png','media_to_paste.jpg')

        Note: for images, the aspect ratio will be the same as the replaced image
        Note2 : it is important to have the source media file as it is required
                to calculate its CRC to find them in the docx
        """
        with open(dst_file, 'rb') as fh:
            crc = self.get_file_crc(src_file)
            self.crc_to_new_media[crc] = fh.read()

    def replace_pic(self,embedded_file,dst_file):
        """Replace embedded picture with original-name given by embedded_file.
           (give only the file basename, not the full path)
           The new picture is given by dst_file (either a filename or a file-like
           object)

        Notes:
            1) embedded_file and dst_file must have the same extension/format
               in case dst_file is a file-like object, no check is done on
               format compatibility
            2) the aspect ratio will be the same as the replaced image
            3) There is no need to keep the original file (this is not the case
               for replace_embedded and replace_media)
        """

        if hasattr(dst_file,'read'):
            # NOTE: file extension not checked
            self.pic_to_replace[embedded_file]=dst_file.read()
        else:
            emp_path,emb_ext=os.path.splitext(embedded_file)
            dst_path,dst_ext=os.path.splitext(dst_file)

            if emb_ext!=dst_ext:
                raise ValueError('replace_pic: extensions must match')

            with open(dst_file, 'rb') as fh:
                self.pic_to_replace[embedded_file]=fh.read()

    def replace_embedded(self,src_file,dst_file):
        """Replace one embdded object by another one into a docx

        This has been done mainly because it is not possible to add images
        in docx header/footer.
        With this function, put a dummy picture in your header/footer,
        then specify it with its replacement in this function

        Syntax: tpl.replace_embedded('dummy_doc.docx','doc_to_paste.docx')

        Note2 : it is important to have the source file as it is required to
                calculate its CRC to find them in the docx
        """
        with open(dst_file, 'rb') as fh:
            crc = self.get_file_crc(src_file)
            self.crc_to_new_embedded[crc] = fh.read()

    def post_processing(self,docx_filename):
        if self.crc_to_new_media or self.crc_to_new_embedded:
            backup_filename = '%s_docxtpl_before_replace_medias' % docx_filename
            os.rename(docx_filename,backup_filename)

            with zipfile.ZipFile(backup_filename) as zin:
                with zipfile.ZipFile(docx_filename, 'w') as zout:
                    for item in zin.infolist():
                        buf = zin.read(item.filename)
                        if ( item.filename.startswith('word/media/') and
                             item.CRC in self.crc_to_new_media ):
                            zout.writestr(item, self.crc_to_new_media[item.CRC])
                        elif ( item.filename.startswith('word/embeddings/')
                               and item.CRC in self.crc_to_new_embedded ):
                            zout.writestr(item, self.crc_to_new_embedded[item.CRC])
                        else:
                            zout.writestr(item, buf)

            os.remove(backup_filename)

    def pre_processing(self):

        if self.pic_to_replace:
            self.build_pic_map()

            # Do the actual replacement
            for embedded_file,stream in six.iteritems(self.pic_to_replace):
                if embedded_file not in self.pic_map:
                    raise ValueError('Picture "%s" not found in the docx template'
                                     % embedded_file)
                self.pic_map[embedded_file][1]._blob=stream

    def build_pic_map(self):
        """Searches in docx template all the xml pictures tag and store them
        in pic_map dict"""
        if self.pic_to_replace:
            # Main document
            part=self.docx.part
            self.pic_map.update(self._img_filename_to_part(part))

            # Header/Footer
            for relid, rel in six.iteritems(self.docx.part.rels):
                if rel.reltype in (REL_TYPE.HEADER,REL_TYPE.FOOTER):
                    self.pic_map.update(self._img_filename_to_part(rel.target_part))

    def get_pic_map(self):
        return self.pic_map

    def _img_filename_to_part(self,doc_part):

        et=etree.fromstring(doc_part.blob)

        part_map={}

        gds=et.xpath('//a:graphic/a:graphicData',namespaces=docx.oxml.ns.nsmap)
        for gd in gds:
            rel=None
            # Either IMAGE, CHART, SMART_ART, ...
            try:
                if gd.attrib['uri']==docx.oxml.ns.nsmap['pic']:
                    # Either PICTURE or LINKED_PICTURE image
                    blip=gd.xpath('pic:pic/pic:blipFill/a:blip',
                                  namespaces=docx.oxml.ns.nsmap)[0]
                    dest=blip.xpath('@r:embed',namespaces=docx.oxml.ns.nsmap)
                    if len(dest)>0:
                        rel=dest[0]
                    else:
                        continue
                else:
                    continue

                #title=inl.xpath('wp:docPr/@title',namespaces=docx.oxml.ns.nsmap)[0]
                name=gd.xpath('pic:pic/pic:nvPicPr/pic:cNvPr/@name',
                              namespaces=docx.oxml.ns.nsmap)[0]

                part_map[name]=(doc_part.rels[rel].target_ref,
                                doc_part.rels[rel].target_part)

            except:
                continue

        return part_map

    def build_url_id(self,url):
        return self.docx._part.relate_to(url, REL_TYPE.HYPERLINK,
                                         is_external=True)

    def save(self,filename,*args,**kwargs):
        self.pre_processing()
        self.docx.save(filename,*args,**kwargs)
        self.post_processing(filename)


class Subdoc(object):
    """ Class for subdocument to insert into master document """
    def __init__(self, tpl,docpath=None):
        self.tpl = tpl
        self.docx = tpl.get_docx()
        self.subdocx = Document(docpath)
        self.subdocx._part = self.docx._part

    def __getattr__(self, name) :
        return getattr(self.subdocx, name)

    def _get_xml(self):
        if self.subdocx._element.body.sectPr is not None:
            self.subdocx._element.body.remove(self.subdocx._element.body.sectPr)
        xml = re.sub(r'</?w:body[^>]*>','',etree.tostring(
            self.subdocx._element.body, encoding='unicode', pretty_print=False))
        return xml

    def __unicode__(self):
        return self._get_xml()

    def __str__(self):
        return self._get_xml()

    def __html__(self):
        return self._get_xml()

class RichText(object):
    """ class to generate Rich Text when using templates variables

    This is much faster than using Subdoc class,
    but this only for texts INSIDE an existing paragraph.
    """
    def __init__(self, text=None, **text_prop):
        self.xml = ''
        if text:
            self.add(text, **text_prop)

    def add(self, text, style=None,
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
                        url_id=None):


        # If not a string : cast to string (ex: int, dict etc...)
        if not isinstance(text, (six.text_type, six.binary_type)):
            text = six.text_type(text)
        if not isinstance(text, six.text_type):
            text = text.decode('utf-8',errors='ignore')
        text = ( escape(text)
                 .replace('\n', NEWLINE_XML)
                 .replace('\a', NEWPARAGRAPH_XML)
                 .replace('\t',TAB_XML) )

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
        if subscript:
            prop += u'<w:vertAlign w:val="subscript"/>'
        if superscript:
            prop += u'<w:vertAlign w:val="superscript"/>'
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
        if font:
            prop += ( u'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
                      .format(font=font) )


        xml = u'<w:r>'
        if prop:
            xml += u'<w:rPr>%s</w:rPr>' % prop
        xml += u'<w:t xml:space="preserve">%s</w:t></w:r>' % text
        if url_id:
            xml = ( u'<w:hyperlink r:id="%s" w:tgtFrame="_blank">%s</w:hyperlink>'
                    % (url_id, xml) )
        self.xml += xml


    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml

    def __html__(self):
        return self.xml

R = RichText

class Listing(object):
    r"""class to manage \n and \a without to use RichText,
    by this way you keep the current template styling

    use {{ mylisting }} in your template and
    context={ mylisting:Listing(the_listing_with_newlines) }
    """
    def __init__(self, text):
        # If not a string : cast to string (ex: int, dict etc...)
        if not isinstance(text, (six.text_type, six.binary_type)):
            text = six.text_type(text)
        self.xml = ( escape(text)
                     .replace('\n', NEWLINE_XML)
                     .replace('\a', NEWPARAGRAPH_XML) )

    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml

    def __html__(self):
        return self.xml


class InlineImage(object):
    """Class to generate an inline image

    This is much faster than using Subdoc class.
    """
    tpl = None
    image_descriptor = None
    width = None
    height = None

    def __init__(self, tpl, image_descriptor, width=None, height=None):
        self.tpl, self.image_descriptor = tpl, image_descriptor
        self.width, self.height = width, height

    def _insert_image(self):
        pic = self.tpl.docx._part.new_pic_inline(
            self.image_descriptor,
            self.width,
            self.height
        ).xml
        return '</w:t></w:r><w:r><w:drawing>%s</w:drawing></w:r><w:r>' \
               '<w:t xml:space="preserve">' % pic

    def __unicode__(self):
        return self._insert_image()

    def __str__(self):
        return self._insert_image()

    def __html__(self):
        return self._insert_image()


