from docxtpl import DocxTemplate, DVM

tpl=DocxTemplate('templates/complex_vertical_merging_tpl.docx')

context = {
    't':range(1,10),
    'B':DVM([(1,3),(4,6),(7,9)]),
    'C':DVM([(3,6),(8,9)],["BBBBBB","EEEEEEE","FF0000","00FF00","0000FF"]),
    'D':DVM([(1,2),(3,7)],["FF0000","0000FF"]),
}

tpl.render(context)
tpl.save('output/complex_vertical_merging.docx')
