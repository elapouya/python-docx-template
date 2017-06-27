0.3.9 (2017-06-27)
------------------
- Fix exception in fix_table()

0.3.8 (2017-06-20)
------------------
- Fix bug when using more than one {{r }} or {%r %} in the same run

0.3.7 (2017-06-13)
------------------
- Fix git tag v0.3.6 was in fact for 0.3.5 package version
  so create a tag 0.3.7 for 0.3.7 package version

0.3.6 (2017-06-10)
------------------
- Better head/footer jinja2 handling (Thanks to hugokernel)

0.3.5 (2017-02-20)
------------------
- Fix bug where one is using '%' (modulo operator) inside a tag

0.3.4 (2017-02-14)
------------------
- Add Listing class to manage \n and \a (new paragraph) and escape text AND keep current styling

0.3.3 (2017-02-07)
------------------
- Add {%tc } tags for dynamic table columns (Thanks to majkls23)

0.3.2 (2017-01-16)
------------------
- Remove version limitation over sphinx package in setup.py

0.3.1 (2017-01-16)
------------------
- Add PNG & JPEG in tests/test_files/

0.3.0 (2017-01-15)
------------------
- You can now add images directly without using subdoc, it is much more faster.

0.2.5 (2017-01-14)
------------------
- Add dynamic colspan tag for tables

0.2.4 (2016-11-30)
------------------
- Fix /n in RichText class

0.2.3 (2016-08-09)
------------------
- Add Python 3 support for footer and header

0.2.2 (2016-06-11)
------------------
- Fix bug when using utf-8 chracters inside footer or header in .docx template
  It now detects header/footer encoding automatically

0.2.1 (2016-06-11)
------------------
- Fix bug where using subdocs is corrupting header and footer in generated docx
  Thanks to Denny Weinberg for his help.

0.2.0 (2016-03-17)
------------------
- Add Header and Footer support (Thanks to Denny Weinberg)

0.1.11 (2016-03-1)
------------------
- '>' and '<' can now be used inside jinja tags

0.1.10 (2016-02-11)
-------------------
- render() accepts optionnal jinja_env argument :
  useful to set custom filters and other things

0.1.9 (2016-01-18)
------------------
- better subdoc management : accept tables

0.1.8 (2015-11-05)
------------------
- better xml code cleaning around Jinja2 tags

0.1.7 (2015-09-09)
------------------
- python 3 support

0.1.6 (2015-05-11)
------------------
- remove debug code
- add lxml dependency

0.1.5 (2015-05-11)
------------------
- fix template filter with quote

0.1.4 (2015-03-27)
------------------
- add RichText support

0.1.3 (2015-03-13)
------------------
- add subdoc support
- add some exemples in tests/

0.1.2 (2015-03-12)
------------------
- First running version
