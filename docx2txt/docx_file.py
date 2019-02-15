import errno
import os.path as os_path
import sys
from os import makedirs
from zipfile import ZipFile

from . import dict_util, xml_util


def simplify_rel(attribs):
    # type: (dict) -> str
    """Simplify Type of a Relationship node

    Arguments:
        attribs {dict} -- attributes of Relationship node

    Returns:
        str - salient portion of rel type (officeDocument, image, etc.)
    """
    attr = attribs.get('Type', '')

    return os_path.basename(attr)


def locate_rel(attribs):
    # type: (dict) -> str
    """Get path to file in Relationship node

    Arguments:
        attribs {dict} -- attributes of Relationship node

    Returns:
        str -- path of Target within package
    """
    attr = attribs.get('Target', '')

    return attr.lstrip('/')


def get_package_rels(pkg_xml):
    # type: (bytes) -> dict
    """Get package relationships

    Arguments:
        pkg_xml {bytes} -- top level relationships XML (_rels/.rels)

    Returns:
        dict -- property and document paths
    """
    rels = xml_util.parse(pkg_xml)

    return {
        simplify_rel(rel.attrib): locate_rel(rel.attrib)
        for rel
        in rels.iter()}


def parse_properties(prop_xml):
    # type: (bytes) -> dict
    """Parse XML for document metadata

    Arguments:
        prop_xml {bytes} -- property XML file (docProps XML)

    Returns:
        dict -- document metadata
    """
    props = xml_util.parse(prop_xml)

    return {xml_util.unquote(prop.tag): prop.text for prop in props.iter()}


def is_property_rel(rel_type):
    # type: (str) -> bool
    """Test for string indicating a property relationship

    Arguments:
        rel_type {str} -- relationship type

    Returns:
        bool -- relationship is a property
    """
    return rel_type.endswith('-properties')


def get_package_props(pkg, pkg_rels):
    # type: (ZipFile, dict) -> dict
    """Get all properties of package

    Arguments:
        pkg {ZipFile} -- package as ZipFile
        pkg_rels {dict} -- all package relationships

    Returns:
        dict -- properties of package
    """
    prop_dicts = [
        parse_properties(pkg.read(path))
        for path
        in dict_util.filter_key(pkg_rels, is_property_rel)]

    return dict_util.merge(prop_dicts)


def get_part_rels_path(part_path):
    # type: (str) -> str
    """Get path to document relationships

    Arguments:
        part_path {str} -- path to officeDocument relationship

    Returns:
        str -- path to relationships for ``part_path``
    """
    path_comps = [
        os_path.dirname(part_path).lstrip('/'),
        '_rels',
        os_path.basename(part_path) + '.rels']

    return '/'.join(path_comps)


def get_part_rels(pkg, part_key, part_path):
    # type: (ZipFile, str, str) -> dict
    """Parse relationships of part (document)

    Arguments:
        pkg {zipfile.ZipFile} -- package ZipFile
        part_key {str} -- key to store path of officeDocument part
        part_path {str} -- path to officeDocument part in package

    Returns:
        dict -- dictionary of XML data
    """
    base_path = os_path.dirname(part_path).lstrip('/')
    rels_path = get_part_rels_path(part_path)
    rel_nodes = xml_util.parse(pkg.read(rels_path))

    rels = {}  # type: dict
    for rel_node in rel_nodes.iter():
        key = simplify_rel(rel_node.attrib)
        path = '/'.join([base_path, rel_node.attrib.get('Target', '')])

        rels[key] = rels.get(key, []) + [path]

    rels.update({part_key: [part_path]})

    return rels


def get_all_rels(pkg, part_type):
    # type: (ZipFile, str) -> tuple
    """Get relationships for package and part of ``part_type``

    Arguments:
        pkg {ZipFile} -- package as ZipFile
        part_type {str} -- type of 'part' to locate (officeDocument)

    Returns:
        tuple -- package relationships, part relationships
    """
    pkg_rels = get_package_rels(pkg.read('_rels/.rels'))
    part_path = pkg_rels.get(part_type, 'word/document.xml')
    part_rels = get_part_rels(pkg, part_type, part_path)

    return pkg_rels, part_rels


def mkdir_p(path):
    # type: (str) -> None
    """Recursively create directory at ``path``

    Arguments:
        path {str} -- directory to create
    """
    try:
        makedirs(path)
    except OSError as err:
        if err.errno != errno.EEXIST or not os_path.isdir(path):
            raise


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
    dst_fname = os_path.join(img_dir, os_path.basename(fname))

    with open(dst_fname, 'wb') as dst_f:
        dst_f.write(img_bytes)

    return os_path.abspath(dst_fname)


def xml2text(xml):
    """
    A string representing the textual content of this run, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    Adapted from: https://github.com/python-openxml/python-docx/
    """
    text = u''
    root = xml_util.parse(xml)
    whitespace_tags = {
        xml_util.quote('w:tab'): '\t',
        xml_util.quote('w:br'): '\n',
        xml_util.quote('w:cr'): '\n',
        xml_util.quote('w:p'): '\n\n', }
    text_tag = xml_util.quote('w:t')
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
    root = xml_util.parse(xml)
    data = {
        xml_util.unquote(child.tag): child.text
        for child in root.iter()}
    return data


def read_docx(path, img_dir):
    # type: (str, str) -> dict
    """Load and parse contents of file at ``path``

    Arguments:
        path {str} -- path to DOCX file

    Keyword Arguments:
        img_dir {str} -- save images in specififed directory

    Returns:
        dict -- header, main, footer, images, and properties of DOCX file
    """
    HEAD_KEY = 'header'
    MAIN_KEY = 'officeDocument'
    FOOT_KEY = 'footer'
    IMG_KEY = 'image'

    with ZipFile(path) as pkg:
        pkg_rels, doc_rels = get_all_rels(pkg, MAIN_KEY)

        text = {
            key: ''.join([
                xml2text(pkg.read(fname))
                for fname in doc_rels.get(key, [])])
            for key in [HEAD_KEY, MAIN_KEY, FOOT_KEY]}  # type: dict

        images = []  # type: list
        if img_dir is None:
            images += [
                os_path.basename(fname)
                for fname in doc_rels.get(IMG_KEY, [])]
        else:
            mkdir_p(img_dir)
            images += [
                extract_image(pkg.read(fname), img_dir, fname)
                for fname in doc_rels.get(IMG_KEY, [])]

        props = get_package_props(pkg, pkg_rels)

    return {
        'header': text.get(HEAD_KEY),
        'main': text.get(MAIN_KEY),
        'footer': text.get(FOOT_KEY),
        'images': images,
        'properties': props, }


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
        return os_path.abspath(str(path))
    except TypeError:
        pass

    # TextIOWrapper, addinfourl, HTTPResponse... and more?
    for attr in (getattr(path, key) for key in ('name', 'url')):
        if attr is not None:
            return str(attr)

    return ''


class DocxFile(object):
    def __init__(self, file, img_dir=None):
        doc_data = read_docx(file, img_dir)

        self._path = get_path(file)                # type: str
        self._img_dir = img_dir                    # type: str
        self._header = doc_data['header']          # type: str
        self._main = doc_data['main']              # type: str
        self._footer = doc_data['footer']          # type: str
        self._images = doc_data['images']          # type: list
        self._properties = doc_data['properties']  # type: dict

    def __str__(self):
        if sys.version_info[0] < 3:
            return self._main.encode('utf-8')

        return self._main

    def __repr__(self):
        return 'DocxFile({!r}, {!r})'.format(self._path, self._img_dir)

    @property
    def path(self):
        return self._path

    @property
    def img_dir(self):
        return self._img_dir

    @property
    def header(self):
        return self._header

    @property
    def main(self):
        return self._main

    @property
    def footer(self):
        return self._footer

    @property
    def images(self):
        return self._images

    @property
    def properties(self):
        return self._properties

    @property
    def text(self):
        return str(self).strip()
