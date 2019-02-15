#! /usr/bin/env python

import argparse
import os
import sys

from . import docx_file


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


def process(docx, img_dir=None):
    document = docx_file.DocxFile(docx, img_dir)
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
