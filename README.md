# docxpy #

This project is forked from [ankushshah89/python-docx2txt](https://github.com/ankushshah89/python-docx2txt/pull/10/files). 
A new feature is added: extract the hyperlinks and its corresponding texts.

It is a pure python-based utility to extract text from docx files. The code is taken and adapted from [python-docx](https://github.com/python-openxml/python-docx). It can however also extract **text** from header, footer and **hyperlinks**. It can now also extract **images**.

## How to install? ##
```bash
pip install docxpy
```

## How to run? ##

a. From command line:
```bash
# extract text
docx2txt file.docx
# extract text and images
docx2txt -i /tmp/img_dir file.docx
```


b. From python:
```python
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
```
