# -*- coding: utf-8 -*-
"""
Regression test for issue #516 : docxtpl must read and write text files as
UTF-8 whatever the locale the interpreter runs under.

This test re-runs itself in a child interpreter whose default encoding is NOT
UTF-8 (PYTHONUTF8=0 + "C" locale), which is what Windows users get with a
legacy code page. That is why the bug could not be reproduced on Linux/macOS,
where the default encoding already is UTF-8.

Without an explicit encoding :
  - `python -m docxtpl` decodes the json data with the locale code page, so
    non-ASCII characters end up mangled in the generated docx (cp1252) or the
    CLI dies with UnicodeDecodeError (cp936, C locale),
  - DocxTemplate.write_xml() dies with UnicodeEncodeError as soon as the
    document contains a character the code page cannot represent.
"""

import locale
import os
import subprocess
import sys

CHILD_ENV_FLAG = "DOCXTPL_NON_UTF8_LOCALE_CHILD"

TEMPLATE_PATH = "templates/module_execute_tpl.docx"
JSON_PATH = "templates/module_execute_utf8.json"
XML_TEMPLATE_PATH = "templates/richtext_eastAsia_tpl.docx"
OUTPUT_FILENAME = "output/utf8_locale.docx"
XML_OUTPUT_FILENAME = "output/utf8_locale.xml"
LATIN1_JSON_FILENAME = "output/utf8_locale_latin1.json"

# Same value as the one stored in templates/module_execute_utf8.json
EXPECTED_TEXT = "äöü 世界 Привет"


def rerun_with_non_utf8_locale():
    env = dict(os.environ)
    env[CHILD_ENV_FLAG] = "1"
    env["PYTHONUTF8"] = "0"  # disable UTF-8 mode (PEP 540)
    env["PYTHONCOERCECLOCALE"] = "0"  # disable C locale coercion (PEP 538)
    env["LC_ALL"] = "C"
    env["LANG"] = "C"
    # Give the bare file name : the child filesystem encoding is ASCII, so an
    # accented character in the path would not survive as an argument.
    return subprocess.call(
        [sys.executable, os.path.basename(__file__)],
        cwd=os.path.dirname(os.path.abspath(__file__)),
        env=env,
    )


def check_json_data_is_read_as_utf8():
    if os.path.exists(OUTPUT_FILENAME):
        os.unlink(OUTPUT_FILENAME)

    subprocess.check_call(
        [sys.executable, "-m", "docxtpl"]
        + [TEMPLATE_PATH, JSON_PATH, OUTPUT_FILENAME, "-o", "-q"]
    )

    from docx import Document

    text = "\n".join(p.text for p in Document(OUTPUT_FILENAME).paragraphs)
    assert EXPECTED_TEXT in text, (
        "python -m docxtpl did not read the json data as UTF-8, got : %r" % text
    )
    print("    --> %s has been generated with UTF-8 data." % OUTPUT_FILENAME)


def check_non_utf8_json_data_is_reported():
    # A json file that is not UTF-8 encoded must be reported like any other
    # bad input, not crash the command line with a UnicodeDecodeError.
    from docxtpl.__main__ import get_json_data

    with open(LATIN1_JSON_FILENAME, "wb") as fh:
        fh.write('{"json_string_var": "caf\xe9"}'.encode("latin-1"))

    try:
        get_json_data(LATIN1_JSON_FILENAME)
    except RuntimeError:
        print("    --> a non UTF-8 json file is reported as a normal error.")
    else:
        raise AssertionError("reading a non UTF-8 json file should have failed")


def check_xml_is_written_as_utf8():
    from docxtpl import DocxTemplate

    tpl = DocxTemplate(XML_TEMPLATE_PATH)
    tpl.init_docx()
    tpl.write_xml(XML_OUTPUT_FILENAME)

    with open(XML_OUTPUT_FILENAME, "rb") as fh:
        written = fh.read().decode("utf-8")
    assert written == tpl.get_xml(), "write_xml() did not write UTF-8"
    print("    --> %s has been written as UTF-8." % XML_OUTPUT_FILENAME)


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    if not os.path.exists("output"):
        os.mkdir("output")
    if os.environ.get(CHILD_ENV_FLAG) != "1":
        print(
            "Re-running %s with a non UTF-8 locale ..." % os.path.basename(__file__),
            flush=True,
        )
        sys.exit(rerun_with_non_utf8_locale())
    if locale.getpreferredencoding(False).lower().replace("-", "") in (
        "utf8",
        "cp65001",
    ):
        print("    --> skipped : the locale encoding already is UTF-8.")
        sys.exit(0)
    check_json_data_is_read_as_utf8()
    check_non_utf8_json_data_is_reported()
    check_xml_is_written_as_utf8()
