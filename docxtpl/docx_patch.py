# -*- coding: utf-8 -*-
"""
Created : 2021-08-14

@author: Eric Lapouyade
"""


from docx.parts.story import BaseStoryPart


class DocxTplBaseStoryPart:
    next_id_cptr = 0

    @property
    def docx_next_id(self):
        id_str_lst = self._element.xpath('//@id')
        used_ids = [int(id_str) for id_str in id_str_lst if id_str.isdigit()]
        if not used_ids:
            DocxTplBaseStoryPart.next_id_cptr += 1  # global counter
            return DocxTplBaseStoryPart.next_id_cptr
        return max(used_ids) + 1


# Patch python_docx next_id() to have no collision on default value
BaseStoryPart.next_id = DocxTplBaseStoryPart.docx_next_id
