from docxtpl import DocxTemplate

tpl = DocxTemplate("templates/less_cells_after_loop_tpl.docx")
tpl.render({})
tpl.save("output/less_cells_after_loop.docx")
