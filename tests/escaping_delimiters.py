from pathlib import Path

from docx import Document
from docxtpl import DocxTemplate


output_dir = Path(__file__).parent / "output"
output_dir.mkdir(exist_ok=True)
template_path = output_dir / "escaping_delimiters_split_tpl.docx"
result_path = output_dir / "escaping_delimiters_split.docx"

escaped_delimiters = ("{_%", "%_}", "{_{", "}_}")
expected_delimiters = ("{%", "%}", "{{", "}}")

template = Document()

for delimiter in escaped_delimiters:
    template.add_paragraph(delimiter)

template.add_paragraph()

for delimiter in escaped_delimiters:
    paragraph = template.add_paragraph()
    for character in delimiter:
        paragraph.add_run(character)

# These are not escaped delimiters and must not be changed.
template.add_paragraph("{a_{hello")
template.add_paragraph("{f_{")

# Empty and formatted runs do not add text.
paragraph = template.add_paragraph()
paragraph.add_run("left }")
paragraph.add_run()
paragraph.add_run("_").bold = True
paragraph.add_run("} right")

# A tab or line break is visible content, even when Word gives it a separate run.
for separator in ("\t", "\n"):
    paragraph = template.add_paragraph()
    paragraph.add_run("left }")
    paragraph.add_run(separator)
    paragraph.add_run("_} right")

template.save(template_path)

doc = DocxTemplate(template_path)
doc.render({})
doc.save(result_path)

paragraphs = [paragraph.text for paragraph in Document(result_path).paragraphs]
assert paragraphs == [
    *expected_delimiters,
    "",
    *expected_delimiters,
    "{a_{hello",
    "{f_{",
    "left }} right",
    "left }\t_} right",
    "left }\n_} right",
]
