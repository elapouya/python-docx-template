# -*- coding: utf-8 -*-
"""
Created : 2021-07-30

@author: Eric Lapouyade
"""
import six
try:
    from html import escape
except ImportError:
    # cgi.escape is deprecated in python 3.7
    from cgi import escape


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
        self.xml = escape(text)

    def __unicode__(self):
        return self.xml

    def __str__(self):
        return self.xml

    def __html__(self):
        return self.xml
