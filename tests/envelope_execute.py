import os
import sys

TEMPLATE_PATH = "templates/envelope_cli_tpl.docx"
JSON_PATH = "templates/envelope_cli.json"
OUTPUT_FILENAME = "output/envelope_cli.docx"
OVERWRITE = "-o"
QUIET = "-q"


os.chdir(os.path.dirname(__file__))
os.makedirs("output", exist_ok=True)

if os.path.exists(OUTPUT_FILENAME):
    os.unlink(OUTPUT_FILENAME)

cmd = "%s -m docxtpl %s %s %s %s %s" % (
    sys.executable,
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
