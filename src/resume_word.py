
# Copyright 2013 Virantha Ekanayake All Rights Reserved.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import zipfile, re
import os,shutil
from lxml import etree
import logging
import itertools
import copy
from collections import OrderedDict

class WordResume(object):

    nsprefixes = {
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    'o':  'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv':  'urn:schemas-microsoft-com:mac:vml',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v':   'urn:schemas-microsoft-com:vml',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    # Properties (core and extended)
    'cp':  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc':  'http://purl.org/dc/elements/1.1/',
    'ep':  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
    # Dublin Core document properties
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms':  'http://purl.org/dc/terms/'}

    def __init__ (self, template_file, resume_data):
        self.resume_data = resume_data
        self.template = template_file
        self.template_filename = template_file.name
        self.zipfile = zipfile.ZipFile(self.template)
        self.doc_etree = self._get_doc_from_docx()

        self._test_func()
        self._write_and_close_docx(self.doc_etree)

    def _check_element_is(self, element, type_char):
        return element.tag == '{%s}%s' % (self.nsprefixes['w'],type_char)
    def _assert_element_is(self, element, type_char):
        assert self._check_element_is(element, type_char)

    def _get_all_text_in_node(self, node):
        all_txt = []
        for txt in node.itertext(tag=etree.Element):
            all_txt.append(txt)
        return ''.join(all_txt)

    def _find_subtags_in_loop(self, my_etree):
        mTag = r"""\[(?P<tag>[\s\w\_]+)\]"""
        tags = OrderedDict()

        # Get the parent paragraph 
        self._assert_element_is(my_etree, 't')
        run = my_etree.getparent()
        self._assert_element_is(run, 'r')
        paragraph = run.getparent()
        self._assert_element_is(paragraph, 'p')

        loop_done = False
        loop_tree = None
        for node, text in self._itersiblingtext(paragraph):
            if '<' in text:
                logging.debug("Found <")
                inside_loop = True
                ind = text.find('<') + 1
                text = text[ind:]
                loop_tree = etree.Element(copy.deepcopy(node.getparent().getparent()))
            if '>' in text:
                assert inside_loop
                ind = text.find( '>')
                text = text[:ind]
                loop_done = True
            if inside_loop:
                tag_text = re.findall(mTag, text)
                if tag_text:
                    logging.debug("Found grps %s" % (','.join(tag_text)))
                for tag in tag_text:
                    tag = tag.lower()
                    tags[tag] = node
                if loop_done:
                    loop_done = False
                    inside_loop = False
                    break
        return tags

    def _find_tags(self, my_etree, tags_to_find, char_to_stop_on=None):
        tags = {}
        logging.debug("Looking for tags: %s" % (','.join(tags_to_find)))
        mTag = r"""\[(?P<tag>[\s\w\_]+)\]"""
        for node,text in self._itertext(my_etree):
            tag_text = re.findall(mTag, text) 
            if tag_text:
                logging.debug("Found grps %s" % (','.join(tag_text)))
            for tag in tag_text:
                tag = tag.lower()
                if tag in tags_to_find:
                    tags[tag] = node
        return tags

    def _get_all_keys_in_list_of_dicts(self, mylist):
        mykeys = set()
        for e in mylist:
            for k in e.keys():
                mykeys.add(k)
        return list(mykeys)

    def _test_func(self):
        body = self.doc_etree.xpath('/w:document/w:body', namespaces=self.nsprefixes)[0]
        self._join_tags(body)

        # Get a list of all the top-level tags
        tags = self._find_tags(self.doc_etree, self.resume_data.keys())
        print tags
        for section_name, node in tags.items():
            logging.debug("Subtag search for %s" % section_name)
            subtag_list = self._get_all_keys_in_list_of_dicts(self.resume_data[section_name])
            logging.debug(subtag_list)
            subtags = self._find_subtags_in_loop(node)
            print subtags
        return

    def _itertext(self, my_etree):
        """Iterator to go through xml tree's text nodes"""
        for node in my_etree.iter(tag=etree.Element):
            if self._check_element_is(node, 't'):
                yield (node, node.text)

    def _itersiblingtext(self, my_etree):
        """Iterator to go through xml tree sibling text nodes"""
        for sib in my_etree.itersiblings(tag=etree.Element):
            for node in sib.iter(tag=etree.Element):
                if self._check_element_is(node, 't'):
                    yield (node, node.text)

    def _join_tags(self, my_etree):
        chars = []
        openbrac = False
        inside_openbrac_node = False

        for node,text in self._itertext(my_etree):
            # Scan through every node with text
            for i,c in enumerate(text):
                # Go through each node's text character by character
                if c == '[':
                    openbrac = True # Within a tag
                    inside_openbrac_node = True # Tag was opened in this node
                    openbrac_node = node # Save ptr to open bracket containing node
                    chars = []
                elif c== ']':
                    assert openbrac
                    if inside_openbrac_node:
                        # Open and close inside same node, no need to do anything
                        pass
                    else:
                        # Open bracket in earlier node, now it's closed
                        # So append all the chars we've encountered since the openbrac_node '['
                        # to the openbrac_node
                        chars.append(']')
                        openbrac_node.text += ''.join(chars)
                        # Also, don't forget to remove the characters seen so far from current node
                        node.text = text[i+1:] 
                    openbrac = False
                    inside_openbrac_node = False
                else:
                    # Normal text character
                    if openbrac and inside_openbrac_node:
                        # No need to copy text
                        pass
                    elif openbrac and not inside_openbrac_node:
                        chars.append(c)
                    else:
                        # outside of a open/close
                        pass
            if openbrac and not inside_openbrac_node:
                # Went through all text that is part of an open bracket/close bracket
                # in other nodes
                # need to remove this text completely
                node.text = ""
            inside_openbrac_node = False

    def _get_doc_from_docx (self):
        xml_content = self.zipfile.read('word/document.xml')
        return etree.fromstring(xml_content)

    def _write_and_close_docx (self, xml_content):
        # Extract all the files and zip it up
        tmp_dir = 'tmp'
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
        os.makedirs(tmp_dir)

        self.zipfile.extractall(tmp_dir)
        with open(os.path.join(tmp_dir,'word/document.xml'), 'w') as f:
            xmlstr = etree.tostring (xml_content, pretty_print=True)
            f.write(xmlstr)

        # Get a list of all the files in the original docx zipfile
        filenames = self.zipfile.namelist()

        # Copy the zip file
        zip_copy_filename = self.template_filename.replace(".docx", "-update.docx")
        with zipfile.ZipFile(zip_copy_filename, "w") as docx:
            for filename in filenames:
                docx.write(os.path.join(tmp_dir,filename), filename)


