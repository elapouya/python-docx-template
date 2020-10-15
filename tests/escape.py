from docxtpl import DocxTemplate, R, Listing

tpl = DocxTemplate('templates/escape_tpl.docx')

context = {
    'myvar': R(
        '"less than" must be escaped : <, this can be done with RichText() or R()'
    ),
    'myescvar': 'Without using Richtext, you can simply escape text with a "|e" jinja filter in the template : < ',
    'nlnp': R(
        'Here is a multiple\nlines\nstring\aand some\aother\aparagraphs\aNOTE: the current character styling is removed'
    ),
    'mylisting': Listing(
        'the listing\nwith\nsome\nlines\nand special chars : <>&\f ... and a page break'
    ),
    'page_break': R('\f'),
    'new_listing': """With the latest version of docxtpl,
there is no need to use Listing objects anymore.
Just use \\n for newline,\n\\t\t for tabulation and \\f for ...\f...page break
One can also use special chars : <>& but you have then to add "|e" jinja filter in the template
""",
}

tpl.render(context)
tpl.save('output/escape.docx')
