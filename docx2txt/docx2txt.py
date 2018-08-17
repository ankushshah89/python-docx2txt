#! /usr/bin/env python

import argparse
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile


def process_args():
    """Parse command line arguments if invoked directly

    Returns:
        object -- .img_dir: output directory, .details: get document details
    """
    desc = 'A pure Python-based utility to extract data from docx files.'
    id_help = 'path of directory to extract images'
    ad_help = 'get all document data'

    parser = argparse.ArgumentParser(description=desc)
    parser.add_argument('docx', help='path of the docx file')
    parser.add_argument('-i', '--img_dir', help=id_help)
    parser.add_argument('-d', '--details', help=ad_help, action='store_true')

    args = parser.parse_args()

    if not os.path.exists(args.docx):
        sys.stderr.write('File {!r} does not exist.'.format(args.docx))
        sys.exit(1)

    if args.img_dir is not None:
        if not os.path.exists(args.img_dir):
            try:
                os.makedirs(args.img_dir)
            except OSError:
                sys.stderr.write(
                    'Unable to create img_dir {!r}'.format(args.img_dir))
                sys.exit(1)
    return args


def get_rel_key(attrib):
    # type: (dict) -> str
    """Get dictionary key for XML node

    Arguments:
        attrib {dict} -- relationship node attributes

    Returns:
        str -- simplified key name
    """
    node_type = attrib.get('Type', '')
    key = str(re.sub(r'.+[\/\-]+', '', node_type))
    return key


def get_rel_path(parent, attrib):
    # type: (str, dict) -> str
    """Get path to relationship in REL file

    Arguments:
        parent {str} -- parent directory of relationship
        attrib {dict} -- relationship node attributes

    Returns:
        str -- full path to relationship
    """
    target = attrib.get('Target', '')
    path = (parent + target).lstrip('/')

    return path


def load_rels(xml, fname):
    # type: (bytes, str) -> dict
    """Parse document REL file

    Arguments:
        xml {bytes} -- contents of XML file
        fname {str} -- path to XML file

    Returns:
        dict -- dictionary of XML data
    """
    root = ET.fromstring(xml)
    base_path = str(re.sub(r'_rels/.+', '', fname))
    data = {}  # type: dict

    for node in root.iter():
        key = get_rel_key(node.attrib)
        path = get_rel_path(base_path, node.attrib)
        data[key] = data.get(key, []) + [path]

    return data


def extract_image(img_bytes, img_dir, fname):
    # type: (bytes, str, str) -> str
    """Write image data to disk

    Arguments:
        img_bytes {bytes} -- image data
        img_dir {str} -- output directory
        fname {str} -- name of source file

    Returns:
        str -- absolute path to extracted image
    """
    dst_fname = os.path.join(img_dir, os.path.basename(fname))

    with open(dst_fname, 'wb') as dst_f:
        dst_f.write(img_bytes)

    return os.path.abspath(dst_fname)


def un(tag):
    # type: (str) -> str
    """Stands for 'unqualified name'. Removes namespace from prefixed tag.

    See: [Python issue 18304](https://bugs.python.org/issue18304)

    Arguments:
        tag {str} -- (possibly-)namespaced tag

    Returns:
        str -- tag name without namespace
    """
    return tag.split('}').pop()


def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    nsmap = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)


def xml2text(xml):
    """
    A string representing the textual content of this run, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    Adapted from: https://github.com/python-openxml/python-docx/
    """
    text = u''
    root = ET.fromstring(xml)
    whitespace_tags = {
        qn('w:tab'): '\t',
        qn('w:br'): '\n',
        qn('w:cr'): '\n',
        qn('w:p'): '\n\n', }
    text_tag = qn('w:t')
    for child in root.iter():
        text += whitespace_tags.get(child.tag, '')
        if child.tag == text_tag and child.text is not None:
            text += child.text
    return text


def xml2dict(xml):
    # type: (bytes) -> dict
    """Get dictionary of values from ``xml``

    Arguments:
        xml {bytes} -- contents of XML file

    Returns:
        dict -- dictionary of {node.tagName: node.text}
    """
    root = ET.fromstring(xml)
    data = {
        un(child.tag): child.text
        for child in root.iter()}
    return data


def parse_docx(path, img_dir):
    # type: (str, str) -> dict
    """Load and parse contents of file at ``path``

    Arguments:
        path {str} -- path to DOCX file

    Keyword Arguments:
        img_dir {str} -- save images in specififed directory (default: {None})

    Returns:
        dict -- header, main, footer, images, and properties of DOCX file
    """
    TEXT_KEYS = ['header', 'officeDocument', 'footer']
    PROP_KEY = 'properties'
    IMG_KEY = 'image'

    zipf = zipfile.ZipFile(path)
    paths = {}
    for fname in ['_rels/.rels', 'word/_rels/document.xml.rels']:
        paths.update(load_rels(zipf.read(fname), fname))

    doc_data = {IMG_KEY: paths[IMG_KEY], PROP_KEY: {}}
    doc_data.update({
        key: ''.join([
            xml2text(zipf.read(fname))
            for fname in paths.get(key, [])])
        for key in TEXT_KEYS})

    if img_dir is not None:
        doc_data[IMG_KEY] = [
            extract_image(zipf.read(fname), img_dir, fname)
            for fname in paths[IMG_KEY]]

    for fname in paths[PROP_KEY]:
        doc_data[PROP_KEY].update(xml2dict(zipf.read(fname)))

    zipf.close()

    return {
        'header': doc_data[TEXT_KEYS[0]],
        'main': doc_data[TEXT_KEYS[1]],
        'footer': doc_data[TEXT_KEYS[2]],
        'images': doc_data[IMG_KEY],
        PROP_KEY: doc_data[PROP_KEY], }


class DocxFile(object):
    def __init__(self, path, img_dir=None):
        doc_data = parse_docx(path, img_dir)

        self.path = os.path.abspath(path)         # type: str
        self.img_dir = img_dir                    # type: str
        self.header = doc_data['header']          # type: str
        self.main = doc_data['main']              # type: str
        self.footer = doc_data['footer']          # type: str
        self.images = doc_data['images']          # type: list
        self.properties = doc_data['properties']  # type: dict

    def __str__(self):
        str_val = ''.join([self.header, self.main, self.footer])

        if sys.version_info[0] < 3:
            return str_val.encode('utf-8')

        return str_val

    def __repr__(self):
        return 'DocxFile({!r}, {!r})'.format(self.path, self.img_dir)

    def __getattr__(self, attr_name):
        if attr_name == 'text':
            return str(self).strip()


def process(docx, img_dir=None):
    document = DocxFile(docx, img_dir)
    return document


def detail_text(prop_name, prop_val):
    return '{:10s}: {!r}\n'.format(prop_name, prop_val)


def get_output():
    args = process_args()
    document = process(args.docx, args.img_dir)

    if args.details:
        yield detail_text('path', document.path)
        yield detail_text('header', document.header)
        yield detail_text('main', document.main)
        yield detail_text('footer', document.footer)
        yield detail_text('images', document.images)
        yield detail_text('properties', document.properties)
    else:
        yield document.text


if __name__ == '__main__':
    for line in get_output():
        sys.stdout.write(line)
