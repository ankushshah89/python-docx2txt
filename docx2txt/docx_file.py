import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile


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


def un_qn(tag):
    # type: (str) -> str
    """Stands for 'unqualified name'. Removes namespace from prefixed tag.

    See: [Python issue 18304](https://bugs.python.org/issue18304)

    Arguments:
        tag {str} -- (possibly-)namespaced tag

    Returns:
        str -- tag name without namespace
    """
    return tag.split('}').pop()


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
        un_qn(child.tag): child.text
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

    doc_data = {
        key: ''.join([
            xml2text(zipf.read(fname))
            for fname in paths.get(key, [])])
        for key in TEXT_KEYS}  # type: dict

    if img_dir is None:
        doc_data[IMG_KEY] = [
            os.path.basename(fname)
            for fname in paths.get(IMG_KEY, [])]
    else:
        doc_data[IMG_KEY] = [
            extract_image(zipf.read(fname), img_dir, fname)
            for fname in paths.get(IMG_KEY, [])]

    doc_data[PROP_KEY] = {}
    for fname in paths[PROP_KEY]:
        doc_data[PROP_KEY].update(xml2dict(zipf.read(fname)))

    zipf.close()

    return {
        'header': doc_data[TEXT_KEYS[0]],
        'main': doc_data[TEXT_KEYS[1]],
        'footer': doc_data[TEXT_KEYS[2]],
        'images': doc_data[IMG_KEY],
        PROP_KEY: doc_data[PROP_KEY], }


def get_path(path):
    # type: (object) -> str
    """Get absolute path to document

    Arguments:
        path {str} -- path to DOCX file (nominal)

    Returns:
        str -- path to document (absolute)
    """
    # simple filesystem path string
    try:
        return os.path.abspath(str(path))
    except TypeError:
        pass

    # TextIOWrapper, addinfourl, HTTPResponse... and more?
    for attr in (getattr(path, key) for key in ('name', 'url')):
        if attr is not None:
            return str(attr)

    return ''


class DocxFile(object):
    def __init__(self, file, img_dir=None):
        doc_data = parse_docx(file, img_dir)

        self.path = get_path(file)                # type: str
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

    @property
    def text(self):
        return str(self).strip()
