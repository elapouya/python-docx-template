from docxtpl import DocxTemplate

tpl=DocxTemplate('test_files/dynamic_table_tpl.docx')

context = {
    'col_labels' : ['fruit', 'vegetable', 'stone', 'thing'],
    'tbl_contents': [
        {'label': 'yellow', 'cols': ['banana', 'capsicum', 'pyrite', 'taxi']},
        {'label': 'red', 'cols': ['apple', 'tomato', 'cinnabar', 'doubledecker']},
        {'label': 'green', 'cols': ['guava', 'cucumber', 'aventurine', 'card']},
        ]
}

tpl.render(context)
tpl.save('test_files/dynamic_table.docx')
