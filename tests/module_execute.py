import os
from pathlib import Path

TEMPLATE_PATH = 'templates/module_execute_tpl.docx'
JSON_PATH = 'templates/module_execute.json'
OUTPUT_FILENAME = 'output/module_execute.docx'
OVERWRITE = '-o'
QUIET = '-q'


output_path = Path(OUTPUT_FILENAME)
if output_path.exists():
    output_path.unlink()

os.chdir(Path(__file__).parent)
cmd = f'python -m docxtpl {TEMPLATE_PATH} {JSON_PATH} {OUTPUT_FILENAME} {OVERWRITE} {QUIET}'
print(f'Executing "{cmd}" ...')
os.system(cmd)

if output_path.exists():
    print(f'    --> File {output_path.resolve()} has been generated.')
