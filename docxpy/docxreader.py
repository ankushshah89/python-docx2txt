#! /usr/bin/env python

import argparse
import xml.etree.ElementTree as ET
import zipfile
import os
import sys
import re


def process_args():
    parser = argparse.ArgumentParser(description='A pure python-based utility '
                                                 'to extract text and images '
                                                 'from docx files.')
    parser.add_argument("docx", help="path of the docx file")
    parser.add_argument('-i', '--img_dir', help='path of directory '
                                                'to extract images')

    args = parser.parse_args()

    if not os.path.exists(args.docx):
        print('File {} does not exist.'.format(args.docx))
        sys.exit(1)

    if args.img_dir is not None:
        if not os.path.exists(args.img_dir):
            try:
                os.makedirs(args.img_dir)
            except OSError:
                print("Unable to create img_dir {}".format(args.img_dir))
                sys.exit(1)
    return args


def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)


class DOCReader(object):
    def __init__(self, docx, img_dir=None):
        if not os.path.exists(docx):
            raise Exception('Can not file document: %s' % docx)
        self.file = docx
        self.img_dir = img_dir
        self.data = {'links': []}  # save header, footer, document, links
        self.links = {}

        # read file
        self.zipf = zipfile.ZipFile(self.file)
        self.filelist = self.zipf.namelist()

        # parse hyperlinks
        hyperlink_document = 'word/_rels/document.xml.rels'
        if hyperlink_document in self.filelist:
            self.process_hyperlink(self.zipf.read(hyperlink_document))

    def process_hyperlink(self, doc):
        """
        external hyperlink from a string of xml document(typically the `word/_rels/document.xml.rels` file)
        """
        root = ET.fromstring(doc)
        nodes = [node.attrib for node in root]
        nodes = filter(lambda x: x.get('TargetMode', '') == 'External', nodes)
        self.links = {node['Id']: node['Target'] for node in nodes}

    def xml2text(self, xml):
        """
        A string representing the textual content of this run, with content
        child elements like ``<w:tab/>`` translated to their Python
        equivalent.
        Adapted from: https://github.com/python-openxml/python-docx/
        """
        text = u''
        root = ET.fromstring(xml)
        for child in root.iter():
            attr = child.attrib
            for k, v in attr.items():
                if k.endswith('id') and v in self.links:
                    self.data['links'].append((ET.tostring(child, encoding='utf-8', method='text'), self.links[v]))
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

    def process(self):
        text = u''
        # get header text
        # there can be 3 header files in the zip
        header_xmls = re.compile('word/header[0-9]*.xml')
        self.data['header'] = [self.xml2text(self.zipf.read(fname)) for fname in self.filelist if header_xmls.match(fname)]
        text += '\n'.join(self.data['header'])

        # get main text
        doc_xml = 'word/document.xml'
        self.data['document'] = self.xml2text(self.zipf.read(doc_xml))
        text += self.data['document']

        # get footer text
        # there can be 3 footer files in the zip
        footer_xmls = re.compile('word/footer[0-9]*.xml')
        self.data['footer'] = [self.xml2text(self.zipf.read(fname)) for fname in self.filelist if footer_xmls.match(fname)]
        text += '\n'.join(self.data['footer'])

        if self.img_dir is not None:
            # extract images
            for fname in self.filelist:
                _, extension = os.path.splitext(fname)
                if extension in [".jpg", ".jpeg", ".png", ".bmp"]:
                    dst_fname = os.path.join(self.img_dir, os.path.basename(fname))
                    with open(dst_fname, "w") as dst_f:
                        dst_f.write(self.zipf.read(fname))
        self.zipf.close()
        return text.strip()


def process(docx, img_dir=None):
    obj = DOCReader(docx, img_dir=img_dir)
    res = obj.process()
    return res


if __name__ == '__main__':
    args = process_args()
    text = process(args.docx, args.img_dir)
    print(text.encode('utf-8'))
