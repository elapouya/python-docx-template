# -*- coding: utf-8 -*-
"""
Created : 2015-03-12

@author: sandeeprah, Eric Lapouyade
"""

from docxtpl import DocxTemplate
import jinja2

jinja_env = jinja2.Environment()


# to create new filters, first create functions that accept the value to filter
# as first argument, and filter parameters as next arguments
def my_filterA(value, my_string_arg):
    return_value = value + " " + my_string_arg
    return return_value


def my_filterB(value, my_float_arg):
    return_value = value + my_float_arg
    return return_value


# Then, declare them to jinja like this :
jinja_env.filters["my_filterA"] = my_filterA
jinja_env.filters["my_filterB"] = my_filterB


context = {"base_value_string": " Hello", "base_value_float": 1.5}

tpl = DocxTemplate("templates/custom_jinja_filters_tpl.docx")
tpl.render(context, jinja_env)
tpl.save("output/custom_jinja_filters.docx")
