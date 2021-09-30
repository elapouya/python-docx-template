from docxtpl.__main__ import main
import sys
from pathlib import Path

TEMPLATE_PATH = 'templates/template_for_main_test.docx'
JSON_PATH = 'templates/json_for_main_test.json'
OUTPUT_FILENAME = 'output/output_for_main_test.docx'
OVERWRITE = '-o'


# Simulate command line arguments, running with overwrite flag so test can be
# run repeatedly without need for confirmation:
sys.argv[1:] = [TEMPLATE_PATH, JSON_PATH, OUTPUT_FILENAME, OVERWRITE]

main()

output_path = Path(OUTPUT_FILENAME)

if output_path.exists():
    print(f'File {output_path.resolve()} exists.')

