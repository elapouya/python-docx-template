# -*- coding: utf-8 -*-
"""
Created : 2017-01-14

@author: Eric Lapouyade
"""

import zipfile
from xml.etree import ElementTree

from docxtpl import DocxTemplate, InlineImage

# for height and width you have to use millimeters (Mm), inches or points(Pt) class :
from docx.shared import Mm
import jinja2

tpl = DocxTemplate("templates/inline_image_tpl.docx")

context = {
    "myimage": InlineImage(
        tpl,
        "templates/python_logo.png",
        width=Mm(20),
        title="python-logo-title",
        descr="python-logo-description",
    ),
    "myimageratio": InlineImage(
        tpl, "templates/python_jpeg.jpg", width=Mm(30), height=Mm(60)
    ),
    "frameworks": [
        {
            "image": InlineImage(tpl, "templates/django.png", height=Mm(10)),
            "desc": "The web framework for perfectionists with deadlines",
        },
        {
            "image": InlineImage(tpl, "templates/zope.png", height=Mm(10)),
            "desc": "Zope is a leading Open Source Application Server "
            "and Content Management Framework",
        },
        {
            "image": InlineImage(tpl, "templates/pyramid.png", height=Mm(10)),
            "desc": "Pyramid is a lightweight Python web framework aimed at taking "
            "small web apps into big web apps.",
        },
        {
            "image": InlineImage(tpl, "templates/bottle.png", height=Mm(10)),
            "desc": "Bottle is a fast, simple and lightweight WSGI micro web-framework "
            "for Python",
        },
        {
            "image": InlineImage(tpl, "templates/tornado.png", height=Mm(10)),
            "desc": "Tornado is a Python web framework and asynchronous networking "
            "library.",
        },
    ],
}
# testing that it works also when autoescape has been forced to True
jinja_env = jinja2.Environment(autoescape=True)
tpl.render(context, jinja_env)
tpl.save("output/inline_image.docx")


with zipfile.ZipFile("output/inline_image.docx") as docx_file:
    document = ElementTree.fromstring(docx_file.read("word/document.xml"))

namespaces = {
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}

picture = document.find(".//pic:cNvPr[@name='python_logo.png']", namespaces)
if picture is None:
    raise ValueError("Missing python_logo.png image metadata")
if picture.get("title") != "python-logo-title":
    raise ValueError("Missing picture title metadata")
if picture.get("descr") != "python-logo-description":
    raise ValueError("Missing picture description metadata")

doc_pr = document.find(".//wp:docPr[@title='python-logo-title']", namespaces)
if doc_pr is None:
    raise ValueError("Missing wp:docPr title metadata")
if doc_pr.get("descr") != "python-logo-description":
    raise ValueError("Missing wp:docPr description metadata")
