0.20.2 *(Unreleased)*
-------------------
- Move docxcompose to optional dependency (Thanks to Waket Zheng)

0.20.1 (2025-07-15)
-------------------
- Fix and improve get_undeclared_template_variables() method (Thanks to Pablo Esteban)

0.20.0 (2024-12-29)
-------------------
- Add RichTextParagraph (Thanks to ST-Imrie)
- Add RTL support for bold/italic (Thanks to bm-rana)
- Update documentation

0.19.1 (2024-12-29)
-------------------
- PR #575 : fix unicode in footnotes (Thanks to Jonathan Pyle)

0.19.0 (2024-11-12)
-------------------
- Support rendering variables in footnotes (Thanks to Bart Broere)

0.18.0 (2024-07-21)
-------------------
- IMPORTANT : Remove Python 2.x support
- Add hyperlink option in InlineImage (Thanks to Jean Marcos da Rosa)
- Update index.rst (Thanks to jkpet)
- Add poetry env
- Black all files

0.17.0 (2024-05-01)
-------------------
- Add support to python-docx 1.1.1

0.16.8 (2024-02-23)
-------------------
- PR #527 : upgrade Jinja2 in Pipfile.lock

0.16.7 (2023-05-08)
-------------------
- PR #493 - thanks to AdrianVorobel

0.16.6 (2023-03-12)
-------------------
- PR #482 - thanks to dreizehnutters

0.16.5 (2023-01-07)
-------------------
- PR #467 - thanks to Slarag
- fix #465
- fix #464

0.16.4 (2022-08-04)
-------------------
- Regional fonts for RichText
- Reorganize documentation

0.16.3 (2022-07-14)
-------------------
- fix #448

0.16.2 (2022-07-14)
-------------------
- fix #444
- fix #443

0.16.1 (2022-06-12)
-------------------
- PR #442

0.16.0 (2022-04-16)
-------------------
- add jinja2 comment support - Thanks to staffanm

0.15.2 (2022-01-12)
-------------------
- fix #408
- Multi-rendering with same DocxTemplate object is now possible
  see tests/multi_rendering.py
- fix #392
- fix #398

0.14.1 (2021-10-01)
-------------------
- One can now use python -m docxtpl on command line
  to generate a docx from a template and a json file as a context
  Thanks to Lcrs123@github

0.12.0 (2021-08-15)
-------------------
- Code has be split into many files for better readability
- Use docxcomposer to attach parts when a docx file is given to create a subdoc
  Images, styles etc... must now be taken in account in subdocs
- Some internal XML IDs are now renumbered to avoid collision, thus images are not randomly disapearing anymore.
- fix #372
- fix #374
- fix #375
- fix #369
- fix #368
- fix #347
- fix #181
- fix #61

0.11.5 (2021-05-09)
-------------------
- PR #351
- It is now possible to put InlineImage in header/footer
- fix #323
- fix #320
- \\n, \\a, \\t and \\f are now accepted in simple context string. Thanks to chabErch@github

0.10.5 (2020-10-15)
-------------------
- Remove extension testing (#297)
- Fix spaces missing in some cases (#116, #227)

0.9.2 (2020-04-26)
-------------------
- Fix #271
- Code styling

0.8.1 (2020-04-14)
-------------------
- fix #266
- docxtpl is now able to use latest python-docx (0.8.10). Thanks to Dutchy-@github.

0.7.0 (2020-04-09)
-------------------
- Add replace_zipname() method to replace Excel and PowerPoint embedded files

0.6.4 (2020-04-06)
-------------------
- Add the possibility to add RichText to a Richtext
- Prevent lxml from attempting to parse None
- PR #207 and #209
- Handle spaces correctly when run are split by Jinja code (#205)
- PR #203
- DocxTemplate now accepts file-like objects (Thanks to edufresne)

0.5.20 (2019-05-23)
-------------------
- Fix #199
- Add support for file-like objects for replace_media (#197)
- Fix  #176
- Delegated autoescaping to Jinja2 Environment (#175)
- Force to use python-docx 0.8.7 (#170)
- Add getting undeclared variables in the template (#171)
- Added `PAGE_BREAK` feature (#168)
- Fixed issue #159: autoescaped values for both str and unicode.
- Fix tables with gridSpan that have less cells after the tc forloop (#164)
- Smart double quotes in jinja tags are now converted into simple double quotes
- Smart quotes in jinja tags are now converted into simple quotes
- Add custom jinja filter example in tests/
- Reformat the code to be a little more PEP8 compliant
- Add {% hm %} tag for table columns horizontal merging (Thanks to nickgashkov)
- Split tests/tests_files dir into templates and output dirs
- autoescape support for python 2.7
- fix issue #154
- Render can now autoescape context dict
- Fix invalid xml parse because using {% vm %}
- Cast to string non-string value given to RichText or Listing objects
- Import html.escape instead of cgi.escape (deprecated)
- Declare package as python2 and python3 compatible for wheel distrib
- Add sub/superscript in RichText
- Fix table vertical merge
- An hyperlink can now be used in RichText

0.4.13 (2018-06-21)
-------------------
- Subdocument can now be based on an existing docx
- Add font option in RichText
- Better tabs and spaces management for MS Word 2016
- Wheel distribution
- Manage autoscaping on InlineImage, Richtext and Subdoc
- Purge MANIFEST.in file
- Accept variables starting with 'r' in {{}} when no space after {{
- Remove debug traces
- Add {% vm %} to merge cell vertically within a loop (Thanks to Arthaslixin)
- use six.iteritems() instead of iteritems for python 3 compatibility
- Fixed Bug #95 on replace_pic() method
- Add replace_pic() method to replace pictures from its filename (Thanks to Riccardo Gusmeroli)
- Improve image attachment for InlineImage ojects
- Add replace_media() method (useful for header/footer images)
- Add replace_embedded() method (useful for embedding docx)

0.3.9 (2017-06-27)
------------------
- Fix exception in fix_table()
- Fix bug when using more than one {{r }} or {%r %} in the same run
- Fix git tag v0.3.6 was in fact for 0.3.5 package version
  so create a tag 0.3.7 for 0.3.7 package version
- Better head/footer jinja2 handling (Thanks to hugokernel)
- Fix bug where one is using '%' (modulo operator) inside a tag
- Add Listing class to manage \n and \a (new paragraph) and escape text AND keep current styling
- Add {%tc } tags for dynamic table columns (Thanks to majkls23)
- Remove version limitation over sphinx package in setup.py
- Add PNG & JPEG in tests/test_files/
- You can now add images directly without using subdoc, it is much more faster.

0.2.5 (2017-01-14)
------------------
- Add dynamic colspan tag for tables
- Fix /n in RichText class
- Add Python 3 support for footer and header
- Fix bug when using utf-8 chracters inside footer or header in .docx template
  It now detects header/footer encoding automatically
- Fix bug where using subdocs is corrupting header and footer in generated docx
  Thanks to Denny Weinberg for his help.
- Add Header and Footer support (Thanks to Denny Weinberg)

0.1.11 (2016-03-1)
------------------
- '>' and '<' can now be used inside jinja tags
- render() accepts optionnal jinja_env argument :
  useful to set custom filters and other things
- better subdoc management : accept tables
- better xml code cleaning around Jinja2 tags
- python 3 support
- remove debug code
- add lxml dependency
- fix template filter with quote
- add RichText support
- add subdoc support
- add some exemples in tests/
- First running version
