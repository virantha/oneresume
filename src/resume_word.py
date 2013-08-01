
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

    def _collapse_tag_runs_in_paragraph(self, first_run, first_run_text_element):
        # Because we know a tag should really not be split at all
        self._assert_element_is(first_run, 'r')
        self._assert_element_is(first_run_text_element, 't')
        all_txt = [self._get_all_text_in_node(first_run)]
        for e in first_run.itersiblings(tag=etree.Element):
            if self._check_element_is(e, 'r'):
                for txt in e.itertext(tag=etree.Element):
                    if txt is not None:
                        all_txt.append(txt)
                    e.getparent().remove(e)
            else:
                # Probably a bookmark or some such nonsense from Word,
                # so remove all these nodes inside a run
                e.getparent().remove(e)
        logging.debug("Collapsing: found text in runs: %s" % all_txt)
        all_txt = ''.join(all_txt)

        #TODO: doesn't this preclude arbitrary text after a close tag?
        assert all_txt.endswith(']')
        first_run_text_element.text = all_txt


    def _find_all_enclosing_subtags(self, etree):
        # Find all [] loops and collaps all tags inside it, and return
        # a hash of tag text to text element containing it
        mTag = r"""\[(?P<tag>[\s\w\_]+)\]"""
        tags = {}
        open_bracket_elements = etree.xpath('.//w:p//w:t[text()[contains(., "[")]]', namespaces=self.nsprefixes)
        loop_open = open_bracket_elements[0] 
        for open_subtag in open_bracket_elements[1:]:
            # Need to collapse all runs containing text between [ and ] 
            # into one run.
            # See if this already holds true for this open_subtag:
             # [blah] [blk] 
            subtags = re.findall(mTag, open_subtab.text)
            if mTag.search(open_subtab.text):
                continue
            # Go to parent paragraph
            # INVARIANT: each subtag must be wholly contained withing a paragraph
            run = open_subtags.getparent()
            self._assert_element_is(run, 'r')
            paragraph = run.getparent()
            self._assert_element_is(paragraph, 'p')

            

    def _find_tag_bodies(self, etree, tags_to_find):
        # Find the [
        # Iterate through siblings until ]
        #   copy this structure, and insert into parent
        mTag = r"""\[(?P<tag>[\s\w\_]+)\]"""
        open_bracket_elements = etree.xpath('.//w:p//w:t[text()[contains(., "[")]]', namespaces=self.nsprefixes)
        loop_open = open_bracket_elements[0] 
        # Everything after this are bodies
        for e in open_bracket_elements[1:]:
            run = e.getparent()
            self._assert_element_is(run, 'r')
            paragraph = run.getparent()
            self._assert_element_is(paragraph, 'p')
            e_text = self._get_all_text_in_node(paragraph)
            grps = re.findall(mTag, e_text) 


    def _find_tags(self, etree, tags_to_find):
        tags = {}
        logging.debug("Looking for tags: %s" % (','.join(tags_to_find)))
        mTag = r"""\[(?P<tag>[\s\w\_]+)\]"""
        open_bracket_elements = etree.xpath('.//w:p//w:t[text()[contains(., "[")]]', namespaces=self.nsprefixes)
        # Each text entry must be part of a run
        # So, get the parent and iterate through it to get all the text
        # Then, do a match for [text], and return the tag to find
        for e in open_bracket_elements:
            # Get the parent run
            run = e.getparent()
            self._assert_element_is(run, 'r')
            paragraph = run.getparent()
            self._assert_element_is(paragraph, 'p')
            e_text = self._get_all_text_in_node(paragraph)
            grps = re.findall(mTag, e_text) 
            logging.debug("Found grps %s" % (','.join(grps)))
            for grp in grps:
                grp = grp.lower()
                if grp in tags_to_find:
                    self._collapse_tag_runs_in_paragraph(run, e)
                    e.text = e.text[1:-1] # Remove brackets from tag
                    print ("Collapsed tag %s" % grp)
                    tags[grp] = paragraph
        return tags



    def _test_func(self):
        body = self.doc_etree.xpath('/w:document/w:body', namespaces=self.nsprefixes)[0]
        self._find_tags(self.doc_etree, self.resume_data.keys())
        return
        # Iterate through the resume yaml
        for section,items in self.resume_data.items():
            element_in_template = self.doc_etree.xpath('.//w:t[text()[contains(., "[")]]', namespaces=self.nsprefixes)
            if len(element_in_template)>0:
                # Get the parent, which is a "w:r" run

                # Save a copy of this origin
                start_element = element_in_template[0]
                # Check if this [ ] tag is split up into multiple runs or not
                # If it is, we need to convert this into a single run
                if not ']' in start_element.text:
                    start_element_text = [start_element.text]
                    element_in_template = element_in_template[0].getparent()
                    assert element_in_template.tag == '{%s}r' % self.nsprefixes['w']

                    for element in element_in_template.itersiblings(tag=etree.Element):
                        # Keep looking for adjacent runs
                        if element.tag == '{%s}r' % self.nsprefixes['w']:
                            # Get all the text out of this run and concatanate it
                            # into the text o the first run
                            for txt in element.itertext(tag=etree.Element):
                                start_element_text.append(txt)
                                element.getparent().remove(element)
                        else:
                            element.getparent().remove(element)
                    start_element_text =  ''.join(start_element_text)
                    assert start_element_text.endswith("]")
                    start_element.text = start_element_text[1:-1]

            if len(element_in_template) != 0:
                print ("found %s" % section)

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

    def add_item_to_section(self, section_name, item): 
        pass

