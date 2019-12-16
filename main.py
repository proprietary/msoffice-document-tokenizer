#!/usr/bin/env python3.8

from xml.etree import ElementTree
from typing import Type
from zipfile import ZipFile
import re
import sys

def lines_from_xml(tree: ElementTree):
    for node in tree.findall('.//*'):
        text = node.text
        if text:
            yield text


def msoffice_document(path: str):
    PPTX_SLIDES = r'^ppt/slides/.*\.xml$'
    DOCX_DOCUMENT = r'^word/document.xml$'
    ext_matches = re.search(r'.*\.([a-z]+)$', path)
    if ext_matches is None:
        raise ValueError(f"Path \"{path}\" doesn't have an extension")
    ext = ext_matches[1]
    if ext in ['docx', 'DOCX']:
        document_path_re = DOCX_DOCUMENT
    elif ext in ['pptx', 'PPTX']:
        document_path_re = PPTX_SLIDES
    with ZipFile(path) as archive:
        for member in archive.namelist():
            if re.search(document_path_re, member):
                with archive.open(member) as f:
                    tree = ElementTree.fromstring(f.read())
                    for line in lines_from_xml(tree):
                        yield line


class Tokens:
    def __init__(self):
        self.tokens = set()


    def add(self, line):
        for token in re.findall(r'(\w+)', line):
            self.tokens.add(token.lower())

    def __iter__(self):
        return self.tokens.__iter__()


if __name__ == '__main__':
    if len(sys.argv) < 1:
        sys.exit(1)
    tokens = Tokens()
    for document_path in sys.argv[1:]:
        for line in msoffice_document(document_path):
            tokens.add(line)
    for token in tokens:
        print(token)
