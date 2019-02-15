"""XML Utilities"""

import xml.etree.ElementTree as ET


def quote(tag):
    """
    Turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    nsmap = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)


def unquote(tag):
    # type: (str) -> str
    """Remove namespace from prefixed tag.

    See: [Python issue 18304](https://bugs.python.org/issue18304)

    Arguments:
        tag {str} -- (possibly-)namespaced tag

    Returns:
        str -- tag name without namespace
    """
    return tag.split('}').pop()


def parse(xml_bytes):
    # type: (bytes) -> ET.Element
    return ET.fromstring(xml_bytes)
