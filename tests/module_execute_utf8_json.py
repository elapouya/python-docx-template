import json
import os
import subprocess
import sys

from docx import Document


TEMPLATE_PATH = "templates/module_execute_tpl.docx"
JSON_PATH = "output/module_execute_utf8.json"
OUTPUT_FILENAME = "output/module_execute_utf8.docx"


os.chdir(os.path.dirname(__file__))

json_data = {
    "json_dict_var": {"json_dict_var": "successfully inserted"},
    "json_array_var": ["json", "array", "var", "successfully", "inserted"],
    "json_string_var": "cafe accented: café; cjk: 测试",
    "json_int_var": 123,
    "json_float_var": 1.234,
    "json_true_var": True,
    "json_false_var": False,
    "json_none_var": None,
}

with open(JSON_PATH, "w", encoding="utf-8") as fh:
    json.dump(json_data, fh, ensure_ascii=False)

if os.path.exists(OUTPUT_FILENAME):
    os.unlink(OUTPUT_FILENAME)

env = os.environ.copy()
env["LC_ALL"] = "C"
env["PYTHONCOERCECLOCALE"] = "0"
env["PYTHONUTF8"] = "0"

cmd = [
    sys.executable,
    "-m",
    "docxtpl",
    TEMPLATE_PATH,
    JSON_PATH,
    OUTPUT_FILENAME,
    "-o",
    "-q",
]

print('Executing "%s" ...' % " ".join(cmd))
subprocess.check_call(cmd, env=env)

doc_text = "\n".join(
    paragraph.text for paragraph in Document(OUTPUT_FILENAME).paragraphs
)
assert json_data["json_string_var"] in doc_text

print("    --> UTF-8 JSON data rendered into %s." % OUTPUT_FILENAME)
