import os

TEMPLATE_PATH = "templates/module_execute_tpl.docx"
JSON_PATH = "templates/module_execute.json"
OUTPUT_FILENAME = "output/module_execute.docx"
OVERWRITE = "-o"
QUIET = "-q"


if os.path.exists(OUTPUT_FILENAME):
    os.unlink(OUTPUT_FILENAME)

os.chdir(os.path.dirname(__file__))
cmd = f"python -m docxtpl {TEMPLATE_PATH} {JSON_PATH} {OUTPUT_FILENAME} {OVERWRITE} {QUIET}"
print(f'Executing "{cmd}" ...')
os.system(cmd)

if os.path.exists(OUTPUT_FILENAME):
    print(f"    --> File {OUTPUT_FILENAME} has been generated.")
