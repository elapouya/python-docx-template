.. python-docx-template documentation master file, created by
   sphinx-quickstart on Thu Mar 12 17:32:17 2015.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to python-docx-template's documentation!
================================================

.. rubric:: Quickstart

To install::
    
    pip install python-docx-template

Usage::

    from python-docx-template import DocxTemplate

    doc = DocxTemplate("my_word_template.docx")
    context = { 'company_name' : 'World company" }
    doc.render(context)
    doc.save("generated_doc.docx") 
    
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

