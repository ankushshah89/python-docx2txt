#! /usr/bin/env python

import glob
import os
import shutil
import xml.etree.ElementTree as ET
import zipfile

from tempfile import mkdtemp

nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{%s}%s' % (uri, tagroot)


def xml2text(fname_xml):
    """
    A string representing the textual content of this run, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    Adapted from: https://github.com/python-openxml/python-docx/
    """
    text = u''
    xml = ET.parse(fname_xml)
    root = xml.getroot()
    for child in root.iter():
        if child.tag == qn('w:t'):
            t_text = child.text
            text += t_text if t_text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag in (qn('w:br'), qn('w:cr')):
            text += '\n'
        elif child.tag == qn("w:p"):
            text += '\n\n'
    return text


def process(docx):
    text = ''

    # unzip the docx into a temp directory
    temp_dir = mkdtemp()
    with zipfile.ZipFile(docx) as zipf:
        zipf.extractall(temp_dir)

    # get header text
    # there can be 3 header files in the zip
    header_xmls = glob.glob(os.path.join(temp_dir, 'word', 'header*.xml'))
    for header_xml in header_xmls:
        text += xml2text(header_xml)

    # get main text
    doc_xml = os.path.join(temp_dir, 'word', 'document.xml')
    text += xml2text(doc_xml)

    # get footer text
    # there can be 3 footer files in the zip
    footer_xmls = glob.glob(os.path.join(temp_dir, 'word', 'footer*.xml'))
    for footer_xml in footer_xmls:
        text += xml2text(footer_xml)

    shutil.rmtree(temp_dir)
    return text.strip()

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print 'Filename missing.'
        sys.exit(1)

    docx = sys.argv[1]
    if not os.path.exists(docx):
        print 'File %s do not exists.' % (docx)
        sys.exit(1)

    text = process(docx)
    sys.stdout.write(text.encode('utf-8'))
