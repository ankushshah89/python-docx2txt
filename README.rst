docxpy
======

|image0| |PyPI|

This project is forked from
`ankushshah89/python-docx2txt <https://github.com/ankushshah89/python-docx2txt/pull/10/files>`__.
A new feature is added: extract the hyperlinks and its corresponding
texts.

It is a pure python-based utility to extract text from docx files. The
code is taken and adapted from
`python-docx <https://github.com/python-openxml/python-docx>`__. It can
however also extract **text** from header, footer and **hyperlinks**. It
can now also extract **images**.

How to install?
---------------

.. code:: bash

    pip install docxpy

How to run?
-----------

a. From command line:

.. code:: bash

    # extract text
    docx2txt file.docx
    # extract text and images
    docx2txt -i /tmp/img_dir file.docx

b. From python:

.. code:: python

    import docxpy

    c = 'file.docx'

    # extract text
    text = docxpy.process(file)

    # extract text and write images in /tmp/img_dir
    text = docxpy.process(file, "/tmp/img_dir")


    # if you want the hyperlinks
    doc = docxpy.DOCReader(file)
    doc.process()  # process file
    hyperlinks = doc.data['links']

.. |image0| image:: https://travis-ci.org/badbye/docxpy.svg?branch=master
.. |PyPI| image:: https://img.shields.io/pypi/pyversions/scrapy-corenlp.svg?style=flat-square
