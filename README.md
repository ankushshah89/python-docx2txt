# python-docx2txt

## Introduction

A pure Python-based utility to extract text from docx files.

The code is taken and adapted from [python-docx](https://github.com/python-openxml/python-docx).
It can however also extract text from header, footer and hyperlinks.
__It can now also extract images and properties.__

It can be used as a [Python library](#Python%20Library)
or from the [command line](#Command%20Line%20Utility).

## Python Library

### Library Installation

```sh
pip install docx2txt
```

### Library Usage

#### Procedural

The library is easy to use procedurally.

```py
>>> import docx2txt
>>> # get document text
>>> docx2txt.process('file.docx')
'header_textmain_textfooter_text'
>>> # or
>>> # get document text, extract images to /tmp/img_dir
>>> process('file.docx', img_dir='/tmp/img_dir/')
'header_textmain_textfooter_text'
```

#### Object Oriented

The DocxFile class provides more granularity.
Its argument list and accompanying behaviors are identical to `process()`.
Document properties are stored as a dictionary.
No keys are guaranteed, so the get() method is recommended.

```py
>>> import docx2txt
>>> # parse Word doc
>>> document = docx2txt.DocxFile('file.docx', img_dir='/tmp/img_dir/')
>>> # path to file
>>> document.path
'/absolute/path/to/file.docx'
>>> # all document text
>>> document.text
'header_textmain_textfooter_text'
>>> # image directory
>>> document.img_dir
>>> '/tmp/img_dir'
>>> # text components
>>> '||'.join([document.header, document.main, document.footer])
'header_text||main_text||footer_text'
>>> # images (filename only if not extracted)
>>> document.images
['/tmp/img_dir/image1.jpg', '/tmp/img_dir/image2.jpg']
>>> # document properties
>>> document.properties
{'property_name': 'property value', ...}
```

## Command Line Utility

### Utility Installation

With this README file as the working directory:

```sh
python setup.py install
```

### Utility Usage

```sh
# simple text extraction
docx2txt file.docx
# get text, extract images to /tmp/img_dir
docx2txt -i /tmp/img_dir file.docx
# get all document data
docx2txt -d file.docx
# get all data, extract images to /tmp/img_dir
docx2txt -d -i /tmp/img_dir file.docx
# same as previous, more simply:
docx2txt -di /tmp/img_dir file.docx
```