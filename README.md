# python-docx2txt #

A pure python-based utility to extract text from docx files. 

The code is taken and adapted from [python-docx](https://github.com/python-openxml/python-docx). It can however also extract text from header, footer and hyperlinks. __It can now also extract images.__ 

## How to install? ##
```bash
pip install docx2txt
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
import docx2txt

# extract text
text = docx2txt.process("file.docx")

# extract text and write images in /tmp/img_dir
text = docx2txt.process("file.docx", "/tmp/img_dir") 
```
