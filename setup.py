from setuptools import setup
import os
import re
import sys

# To register onto Pypi :
# python setup.py sdist bdist_wheel upload


def read(*names):
    values = dict()
    for name in names:
        filename = name + ".rst"
        if os.path.isfile(filename):
            fd = open(filename)
            value = fd.read()
            fd.close()
        else:
            value = ""
        values[name] = value
    return values


long_description = """
%(README)s

News
====
%(CHANGES)s
""" % read(
    "README", "CHANGES"
)


def get_version(pkg):
    path = os.path.join(os.path.dirname(__file__), pkg, "__init__.py")
    if sys.version_info >= (3, 0):
        fh = open(path, encoding="utf-8")  # required to read utf-8 file on windows
    else:
        fh = open(path)  # encoding parameter does not exist in python 2
    with fh:
        m = re.search(r'^__version__\s*=\s*[\'"]([^\'"]+)[\'"]', fh.read(), re.M)
    if m:
        return m.group(1)
    raise RuntimeError("Unable to find __version__ string in %s." % path)


setup(
    name="docxtpl",
    version=get_version("docxtpl"),
    description="Python docx template engine",
    long_description=long_description,
    long_description_content_type="text/x-rst",
    classifiers=[
        "Intended Audience :: Developers",
        "Development Status :: 4 - Beta",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    keywords="jinja2",
    url="https://github.com/elapouya/python-docx-template",
    author="Eric Lapouyade",
    license="LGPL-2.1-only",
    license_files=[],
    packages=["docxtpl"],
    install_requires=["python-docx>=1.1.1", "docxcompose", "jinja2", "lxml"],
    extras_require={"docs": ["Sphinx", "sphinxcontrib-napoleon"]},
    eager_resources=["docs"],
    zip_safe=False,
)
