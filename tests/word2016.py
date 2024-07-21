from docxtpl import DocxTemplate, RichText

tpl = DocxTemplate("templates/word2016_tpl.docx")
tpl.render(
    {
        "test_space": "          ",
        "test_tabs": 5 * "\t",
        "test_space_r": RichText("          "),
        "test_tabs_r": RichText(5 * "\t"),
    }
)
tpl.save("output/word2016.docx")
