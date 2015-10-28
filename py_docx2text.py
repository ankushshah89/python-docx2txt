#! /usr/bin/env python

import HTMLParser
import glob
import os
import re
import shutil
import sys
import zipfile

from tempfile import mkdtemp


def xml2text(fname_xml):
    with open(fname_xml, 'rb') as f:
        xml = f.read()
        xml = unicode(xml, encoding="utf-8")

    xml = re.sub("</w:p>", "\n\n", xml)  # \n\n for end of paras
    xml = re.sub("<w:tab/>", "\t", xml)  # \t for tabs
    xml = re.sub("<w:br/>|<w:cr/>", "\n", xml)  # \n for new line
    xml = re.sub("<[^>]*>", "", xml)  # remove all xml elements

    # # convert xml entities like &amp; to "&", &lt; to "<", etc.
    text = HTMLParser.HTMLParser().unescape(xml)  # only for python2.x

    return text


def get_text(docx):
    text = ""

    # unzip the docx into a temp directory
    temp_dir = mkdtemp()
    with zipfile.ZipFile(docx) as zipf:
        zipf.extractall(temp_dir)

    # get header text
    # there can be 3 header files in the zip
    header_xmls = glob.glob(os.path.join(temp_dir, "word", "header*.xml"))
    for header_xml in header_xmls:
        text += xml2text(header_xml)

    # get main text
    doc_xml = os.path.join(temp_dir, "word", "document.xml")
    text += xml2text(doc_xml)

    # get footer text
    # there can be 3 footer files in the zip
    footer_xmls = glob.glob(os.path.join(temp_dir, "word", "footer*.xml"))
    for footer_xml in footer_xmls:
        text += xml2text(footer_xml)

    shutil.rmtree(temp_dir)
    return text.strip()

if __name__ == "__main__":
    docx = sys.argv[1]
    text = get_text(docx)
    sys.stdout.write(text.encode("utf-8"))
