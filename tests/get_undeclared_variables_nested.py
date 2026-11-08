#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test for the nested option of get_undeclared_template_variables
"""

import os

from docx import Document

from docxtpl import DocxTemplate

TEMPLATE_PATH = "output/get_undeclared_variables_nested_tpl.docx"


def make_template():
    os.makedirs("output", exist_ok=True)
    doc = Document()
    doc.add_paragraph("{{ foo.bar }} {{ foo['baz'] }} {{ plain }}")
    doc.add_paragraph("{{ order.customer.name }} {{ items[0] }}")
    doc.add_paragraph("{% for row in rows %}{{ row.id }} {{ other.value }}{% endfor %}")
    doc.save(TEMPLATE_PATH)


def test_nested():
    make_template()
    tpl = DocxTemplate(TEMPLATE_PATH)
    expected = {
        "foo.bar",
        "foo.baz",
        "plain",
        "order.customer.name",
        "items",
        "rows",
        "other.value",
    }
    found = tpl.get_undeclared_template_variables(nested=True)
    assert found == expected, found

    # the default behavior still returns the top level names only
    assert tpl.get_undeclared_template_variables() == {
        "foo",
        "plain",
        "order",
        "items",
        "rows",
        "other",
    }

    # variables given in the context are left out
    found = tpl.get_undeclared_template_variables(context={"foo": {}}, nested=True)
    assert found == expected - {"foo.bar", "foo.baz"}, found


if __name__ == "__main__":
    test_nested()
    print("PASS: nested undeclared variables")
