# -*- coding: utf-8 -*-
import os
import re
import shutil
import zipfile
from xml.sax.saxutils import escape as xml_escape

from jinja2 import Environment

from docxtpl import DocxTemplate

TEMPLATE = "templates/hyperlink_tpl.docx"
OUTPUT_DIR = "output"


def _patch_hyperlink_docx(src, dest, target, label=None):
    """Copy a docx and rewrite the external hyperlink Target (and optional label)."""
    shutil.copy(src, dest)
    with zipfile.ZipFile(dest, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}

    rels_name = "word/_rels/document.xml.rels"
    rels = parts[rels_name].decode()
    target_xml = xml_escape(target, {"\"": "&quot;"})

    def replace_target(match):
        return match.group(1) + target_xml + match.group(2)

    rels, count = re.subn(
        r'(Type="[^"]*hyperlink"[^>]*\bTarget=")[^"]*(")',
        replace_target,
        rels,
        count=1,
    )
    if count == 0:
        rels, count = re.subn(
            r'(\bTarget=")[^"]*("[^>]*Type="[^"]*hyperlink")',
            replace_target,
            rels,
            count=1,
        )
    assert count == 1, "hyperlink Target not found in relationships"
    parts[rels_name] = rels.encode()

    if label is not None:
        document = parts["word/document.xml"].decode()

        def replace_label(match):
            return match.group(1) + label + match.group(3)

        document, count = re.subn(
            r"(<w:hyperlink[^>]*>.*?<w:t[^>]*>)(.*?)(</w:t>)",
            replace_label,
            document,
            count=1,
            flags=re.DOTALL,
        )
        assert count == 1, "hyperlink label not found in document"
        parts["word/document.xml"] = document.encode()

    os.remove(dest)
    with zipfile.ZipFile(dest, "w") as zout:
        for name, data in parts.items():
            zout.writestr(name, data)


def _read_rels_and_document(path):
    with zipfile.ZipFile(path) as docx:
        return (
            docx.read("word/_rels/document.xml.rels").decode(),
            docx.read("word/document.xml").decode(),
        )


def test_basic_hyperlink():
    tpl = DocxTemplate(TEMPLATE)
    context = {"foo": "https://example.com"}
    tpl.render(context)
    output_path = os.path.join(OUTPUT_DIR, "hyperlink.docx")
    tpl.save(output_path)

    rels, document = _read_rels_and_document(output_path)
    assert 'Target="https://example.com"' in rels, rels
    assert "{{ foo }}" not in rels, rels
    assert re.search(r"<w:t[^>]*>https://example.com</w:t>", document), document

    undeclared = DocxTemplate(TEMPLATE).get_undeclared_template_variables(
        context=context
    )
    assert undeclared == set(), undeclared


def test_custom_delimiters():
    patched = os.path.join(OUTPUT_DIR, "hyperlink_custom_delim_tpl.docx")
    _patch_hyperlink_docx(TEMPLATE, patched, "[[ foo ]]", label="[[ foo ]]")

    env = Environment(variable_start_string="[[", variable_end_string="]]")
    context = {"foo": "https://example.com"}

    undeclared = DocxTemplate(patched).get_undeclared_template_variables(jinja_env=env)
    assert "foo" in undeclared, undeclared

    undeclared = DocxTemplate(patched).get_undeclared_template_variables(
        jinja_env=env, context=context
    )
    assert undeclared == set(), undeclared

    tpl = DocxTemplate(patched)
    tpl.render(context, jinja_env=env)
    output_path = os.path.join(OUTPUT_DIR, "hyperlink_custom_delim.docx")
    tpl.save(output_path)

    rels, document = _read_rels_and_document(output_path)
    assert 'Target="https://example.com"' in rels, rels
    assert "[[ foo ]]" not in rels, rels
    assert re.search(r"<w:t[^>]*>https://example.com</w:t>", document), document


def test_mixed_percent_and_braces():
    url_tpl = "https://example.com/{{ path }}?q=%7Bbar%7D&id={{ id }}"
    patched = os.path.join(OUTPUT_DIR, "hyperlink_mixed_tpl.docx")
    _patch_hyperlink_docx(TEMPLATE, patched, url_tpl, label="{{ path }}")

    context = {"path": "docs/page", "id": "42"}
    tpl = DocxTemplate(patched)
    tpl.render(context)
    output_path = os.path.join(OUTPUT_DIR, "hyperlink_mixed.docx")
    tpl.save(output_path)

    rels, _document = _read_rels_and_document(output_path)
    # & is stored as &amp; in relationship XML attributes
    expected = "https://example.com/docs/page?q=%7Bbar%7D&amp;id=42"
    assert 'Target="%s"' % expected in rels, rels
    assert "{{ path }}" not in rels, rels
    assert "{{ id }}" not in rels, rels
    assert "%7Bbar%7D" in rels, rels


if __name__ == "__main__":
    if not os.path.exists(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)
    test_basic_hyperlink()
    test_custom_delimiters()
    test_mixed_percent_and_braces()
    print("hyperlink test passed")
