import os
import subprocess
import sys

TEMPLATE_PATH = "templates/module_execute_tpl.docx"
JSON_PATH = "templates/module_execute.json"
OUTPUT_FILENAME = "output/module_execute.docx"
OVERWRITE = "-o"
QUIET = "-q"


if os.path.exists(OUTPUT_FILENAME):
    os.unlink(OUTPUT_FILENAME)

os.chdir(os.path.dirname(__file__))
cmd = [
    sys.executable,
    "-m",
    "docxtpl",
    TEMPLATE_PATH,
    JSON_PATH,
    OUTPUT_FILENAME,
    OVERWRITE,
    QUIET,
]
print('Executing "%s" ...' % " ".join(cmd))
subprocess.run(cmd, check=True)

if os.path.exists(OUTPUT_FILENAME):
    print("    --> File %s has been generated." % OUTPUT_FILENAME)
