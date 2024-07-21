import os

TEMPLATE_PATH = "templates/module_execute_tpl.docx"
JSON_PATH = "templates/module_execute.json"
OUTPUT_FILENAME = "output/module_execute.docx"
OVERWRITE = "-o"
QUIET = "-q"


if os.path.exists(OUTPUT_FILENAME):
    os.unlink(OUTPUT_FILENAME)

os.chdir(os.path.dirname(__file__))
cmd = "python -m docxtpl %s %s %s %s %s" % (
    TEMPLATE_PATH,
    JSON_PATH,
    OUTPUT_FILENAME,
    OVERWRITE,
    QUIET,
)
print('Executing "%s" ...' % cmd)
os.system(cmd)

if os.path.exists(OUTPUT_FILENAME):
    print("    --> File %s has been generated." % OUTPUT_FILENAME)
