# -*- coding: utf-8 -*-
'''
Created : 2017-01-14

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate, InlineImage
# for height and width you have to use millimeters (Mm), inches or points(Pt) class :
from docx.shared import Mm, Inches, Pt
import jinja2
from jinja2.utils import Markup

tpl=DocxTemplate('templates/inline_image_tpl.docx')

context = {
    'myimage' : InlineImage(tpl,'templates/python_logo.png',width=Mm(20)),
    'myimageratio': InlineImage(tpl, 'templates/python_jpeg.jpg', width=Mm(30), height=Mm(60)),

    'frameworks' : [{'image' : InlineImage(tpl,'templates/django.png',height=Mm(10)),
                      'desc' : 'The web framework for perfectionists with deadlines'},

                    {'image' : InlineImage(tpl,'templates/zope.png',height=Mm(10)),
                     'desc' : 'Zope is a leading Open Source Application Server and Content Management Framework'},

                    {'image': InlineImage(tpl, 'templates/pyramid.png', height=Mm(10)),
                     'desc': 'Pyramid is a lightweight Python web framework aimed at taking small web apps into big web apps.'},

                    {'image' : InlineImage(tpl,'templates/bottle.png',height=Mm(10)),
                     'desc' : 'Bottle is a fast, simple and lightweight WSGI micro web-framework for Python'},

                    {'image': InlineImage(tpl, 'templates/tornado.png', height=Mm(10)),
                     'desc': 'Tornado is a Python web framework and asynchronous networking library.'},
                    ]
}
# testing that it works also when autoescape has been forced to True
jinja_env = jinja2.Environment(autoescape=True)
tpl.render(context, jinja_env)
tpl.save('output/inline_image.docx')
