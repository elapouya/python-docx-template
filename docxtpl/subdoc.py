# -*- coding: utf-8 -*-
"""
Created : 2021-07-30

@author: Eric Lapouyade
"""

from docx import Document
from docx.oxml import CT_SectPr
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docxcompose.properties import CustomProperties
from docxcompose.utils import xpath
from docxcompose.composer import Composer
from docxcompose.utils import NS
from lxml import etree
import re


W14_ANCHOR_ID = (
    "{http://schemas.microsoft.com/office/word/2010/wordml}anchorId"
)


class SubdocComposer(Composer):
    def attach_parts(self, doc, remove_property_fields=True):
        """Attach docx parts instead of appending the whole document
        thus subdoc insertion can be delegated to jinja2"""
        self.reset_reference_mapping()

        # Remove custom property fields but keep the values
        if remove_property_fields:
            cprops = CustomProperties(doc)
            for name in cprops.keys():
                cprops.dissolve_fields(name)

        self._create_style_id_mapping(doc)
        self.renumber_ole_objects(doc)

        for element in doc.element.body:
            if isinstance(element, CT_SectPr):
                continue
            self.add_referenced_parts(doc.part, self.doc.part, element)
            self.add_styles(doc, element)
            self.add_numberings(doc, element)
            self.restart_first_numbering(doc, element)
            self.add_images(doc, element)
            self.add_diagrams(doc, element)
            self.add_shapes(doc, element)
            self.add_footnotes(doc, element)
            self.remove_header_and_footer_references(doc, element)

        self.add_styles_from_other_parts(doc)
        self.renumber_bookmarks()
        self.renumber_docpr_ids()
        self.renumber_nvpicpr_ids()
        self.fix_section_types(doc)

    def renumber_ole_objects(self, doc):
        state = getattr(self.doc.part, "_docxtpl_ole_id_state", None)
        if state is None:
            body = self.doc.element.body
            state = {
                "shape_ids": set(xpath(body, ".//v:shape/@id")),
                "shape_index": 1025,
                "object_ids": set(xpath(body, ".//o:OLEObject/@ObjectID")),
                "object_index": 1,
                "anchor_ids": {
                    element.get(W14_ANCHOR_ID).upper()
                    for element in body.iter()
                    if element.get(W14_ANCHOR_ID) is not None
                },
                "anchor_index": 1,
            }
            self.doc.part._docxtpl_ole_id_state = state

        state["shape_ids"].update(
            xpath(doc.element.body, ".//v:shape/@id")
        )
        for ole_object in xpath(doc.element.body, ".//o:OLEObject"):
            old_shape_id = ole_object.get("ShapeID")
            matching_shapes = [
                shape
                for shape in xpath(ole_object.getparent(), "./v:shape")
                if (
                    old_shape_id is not None
                    and shape.get("id") == old_shape_id
                )
            ]
            if len(matching_shapes) == 1:
                shape_id = self._next_ole_id(
                    state, "shape", lambda index: "_x0000_i%d" % index
                )
                matching_shapes[0].set("id", shape_id)
                ole_object.set("ShapeID", shape_id)
            if ole_object.get("ObjectID") is not None:
                ole_object.set(
                    "ObjectID",
                    self._next_ole_id(
                        state, "object", lambda index: "_%d" % index
                    ),
                )

        for element in doc.element.body.iter():
            if element.get(W14_ANCHOR_ID) is not None:
                element.set(
                    W14_ANCHOR_ID,
                    self._next_ole_id(
                        state, "anchor", lambda index: "%08X" % index
                    ),
                )

    @staticmethod
    def _next_ole_id(state, kind, formatter):
        index_key = "%s_index" % kind
        ids_key = "%s_ids" % kind
        while formatter(state[index_key]) in state[ids_key]:
            state[index_key] += 1
        value = formatter(state[index_key])
        state[index_key] += 1
        state[ids_key].add(value)
        return value

    def add_diagrams(self, doc, element):
        # While waiting docxcompose 1.3.3
        dgm_rels = xpath(element, ".//dgm:relIds[@r:dm]")
        for dgm_rel in dgm_rels:
            for item, rt_type in (
                ("dm", RT.DIAGRAM_DATA),
                ("lo", RT.DIAGRAM_LAYOUT),
                ("qs", RT.DIAGRAM_QUICK_STYLE),
                ("cs", RT.DIAGRAM_COLORS),
            ):
                dm_rid = dgm_rel.get("{%s}%s" % (NS["r"], item))
                dm_part = doc.part.rels[dm_rid].target_part
                new_rid = self.doc.part.relate_to(dm_part, rt_type)
                dgm_rel.set("{%s}%s" % (NS["r"], item), new_rid)


class Subdoc(object):
    """Class for subdocument to insert into master document"""

    def __init__(self, tpl, docpath=None):
        self.tpl = tpl
        self.docx = tpl.get_docx()
        self.subdocx = Document(docpath)
        if docpath:
            compose = SubdocComposer(self.docx)
            compose.attach_parts(self.subdocx)
        else:
            self.subdocx._part = self.docx._part

    def __getattr__(self, name):
        return getattr(self.subdocx, name)

    def _get_xml(self):
        if self.subdocx.element.body.sectPr is not None:
            self.subdocx.element.body.remove(self.subdocx.element.body.sectPr)
        xml = re.sub(
            r"</?w:body[^>]*>",
            "",
            etree.tostring(
                self.subdocx.element.body, encoding="unicode", pretty_print=False
            ),
        )
        return xml

    def __unicode__(self):
        return self._get_xml()

    def __str__(self):
        return self._get_xml()

    def __html__(self):
        return self._get_xml()
