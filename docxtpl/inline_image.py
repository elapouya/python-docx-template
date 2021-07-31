# -*- coding: utf-8 -*-
"""
Created : 2021-07-30

@author: Eric Lapouyade
"""


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
        pic = self.tpl.current_rendering_part.new_pic_inline(
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
