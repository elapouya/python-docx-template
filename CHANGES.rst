0.5.6 (2018-10-18)
------------------
- Fix invalid xml parse because using {% vm %}

0.5.5 (2018-10-05)
------------------
- Cast to string non-string value given to RichText or Listing objects
- Import html.escape instead of cgi.escape (deprecated)

0.5.4 (2018-09-19)
------------------
- Declare package as python2 and python3 compatible for wheel distrib

0.5.3 (2018-09-19)
------------------
- Add sub/superscript in RichText

0.5.2 (2018-09-13)
------------------
- Fix table vertical merge

0.5.0 (2018-08-03)
------------------
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
