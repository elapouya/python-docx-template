.. python-docx-template documentation master file, created by
   sphinx-quickstart on Thu Mar 12 17:32:17 2015.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to python-docx-template's documentation!
================================================

.. rubric:: Quickstart

To install::

    pip install docxtpl

Usage::

    from python-docx-template import DocxTemplate

    doc = DocxTemplate("my_word_template.docx")
    context = { 'company_name' : 'World company" }
    doc.render(context)
    doc.save("generated_doc.docx")

Introduction
------------

This package uses 2 major packages :

- python-docx for reading, writing and creating sub documents
- jinja2 for managing tags inserted into the template docx

python-docx-template has been created because python-docx is powerful for creating documents but not for modifying them.

The idea is to begin to create an exemple of the document you want to generate with microsoft word, it can be as complex as you want :
pictures, index tables, footer, header, variables, anything you can do with word.
Then, as you are still editing the document with microsoft word, you insert jinja2-like tags directly in the document.
You save the document as a .docx file (xml format) : it will be your .docx template file.

Now you can use python-docx-template to generate as many word documents you want from this .docx template and context variables you will associate.

Jinja2-like syntax
------------------

As the Jinja2 package is used, one can use all jinja2 tags and filters inside the word document.
Nevertheless there are some restrictions and extensions to make it work inside a word document:

Restrictions
++++++++++++

The usual jinja2 tags, are only to be used inside a same run of a same paragraph, it can not be used across several paragraphs, table rows, runs.

Note : a 'run' for microsoft word is a sequence of characters with the same style. For exemple, if you create a paragraph with all characters the same style :
word will create internally only one 'run' in the paragraph. Now, if you put in bold a text in the middle of this paragraph, word will transform the previous 'run' into 3 'runs' (normal - bold - normal).

Extensions
++++++++++

Tags
....

In order to manage paragraphs, table rows, runs, special syntax has to be used ::

   {%p jinja2_tag %} for paragraphs
   {%tr jinja2_tag %} for table rows
   {%r jinja2_tag %} for runs

By using these tags, python-docx-template will take care to put the real jinja2 tags at the right place into the document's xml source code.
In addition, these tags also tells python-docx-template to remove the paragraph, table row or run where are located the begin and ending tags and only takes care about what is in between.

Display variables
.................

As part of jinja2, one can used double braces::

   {{ <var> }}

But if `<var>` is an RichText object, you must specify that you are changing the actual 'run' ::

   {{r <var> }}

Note the 'r' right after the openning braces

Cell color
..........

There is a special case when you want to change the background color of a table cell, you must put the following tag at the very beginning of the cell ::

   {% cellbg <var> %}

`<var>` must contain the color's hexadecimal code *without* the hash sign

Escaping
........

In order to display `{%`, `%}`, `{{` or `}}`, one can use ::

   {_%, %_}, {_{ or  }_}

.. rubric:: Functions index

.. currentmodule:: docxtpl

.. to list all function : grep "def " *.py | sed -e 's,^def ,,' -e 's,(.*,,' | sort

.. rubric:: Functions documentation

.. automodule:: docxtpl
   :members:


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

