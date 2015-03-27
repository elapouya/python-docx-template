from setuptools import setup
import os

def read(*names):
    values = dict()
    for name in names:
        filename = name + '.rst'
        if os.path.isfile(filename):
            fd = open(filename)
            value = fd.read()
            fd.close()
        else:
            value = ''
        values[name] = value
    return values


long_description = """
%(README)s

News
====
%(CHANGES)s
""" % read('README', 'CHANGES')

setup(name='docxtpl',
      version='0.1.4',
      description='Python docx template engine',
      long_description=long_description,
      classifiers=[
          "Intended Audience :: Developers",
          "Development Status :: 4 - Beta",
          "Programming Language :: Python :: 2",
          "Programming Language :: Python :: 2.7",
      ],
      keywords='jinja2',
      url='https://github.com/elapouya/python-docx-template',
      author='Eric Lapouyade',
      author_email='elapouya@gmail.com',
      license='LGPL 2.1',
      packages=['docxtpl'],
      install_requires = ['Sphinx<1.3b', 'sphinxcontrib-napoleon', 'python-docx','jinja2'],
      eager_resources = ['docs'],
      zip_safe=False)
