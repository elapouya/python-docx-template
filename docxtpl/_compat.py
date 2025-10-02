# -*- coding: utf-8 -*-
try:
    from html import escape
except ImportError:
    # cgi.escape is deprecated in python 3.7
    from cgi import escape  # type:ignore[attr-defined,no-redef]


__all__ = ("escape",)
