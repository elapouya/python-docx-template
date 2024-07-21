from docxtpl import DocxTemplate, R, Listing

tpl = DocxTemplate("templates/escape_tpl.docx")

context = {
    "myvar": R(
        '"less than" must be escaped : <, this can be done with RichText() or R()'
    ),
    "myescvar": 'It can be escaped with a "|e" jinja filter in the template too : < ',
    "nlnp": R(
        "Here is a multiple\nlines\nstring\aand some\aother\aparagraphs",
        color="#ff00ff",
    ),
    "mylisting": Listing("the listing\nwith\nsome\nlines\nand special chars : <>& ..."),
    "page_break": R("\f"),
    "new_listing": """
This is a new listing
Now, does not require Listing() Object
Here is a \t tab\a
Here is a new paragraph\a
Here is a page break : \f
That's it
""",
    "some_html": (
        "HTTP/1.1 200 OK\n"
        "Server: Apache-Coyote/1.1\n"
        "Cache-Control: no-store\n"
        "Expires: Thu, 01 Jan 1970 00:00:00 GMT\n"
        "Pragma: no-cache\n"
        "Content-Type: text/html;charset=UTF-8\n"
        "Content-Language: zh-CN\n"
        "Date: Thu, 22 Oct 2020 10:59:40 GMT\n"
        "Content-Length: 9866\n"
        "\n"
        "<html>\n"
        "<head>\n"
        "    <title>Struts Problem Report</title>\n"
        "    <style>\n"
        "    \tpre {\n"
        "\t    \tmargin: 0;\n"
        "\t        padding: 0;\n"
        "\t    }    "
        "\n"
        "    </style>\n"
        "</head>\n"
        "<body>\n"
        "...\n"
        "</body>\n"
        "</html>"
    ),
}

tpl.render(context)
tpl.save("output/escape.docx")
