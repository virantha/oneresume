#!/usr/bin/env python2.7
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


"""
    OneResume

        - Write your resume in YAML
        - Output it to word, html, txt, etc
"""

import argparse
import sys, os
import logging
import yaml

from resume_word import WordResume
from resume_text import TextResume

def error(text):
    print("ERROR: %s" % text)
    sys.exit(-1)


class OneResume(object):

    def __init__ (self):
        pass

    def getOptions(self, argv):
        p = argparse.ArgumentParser(prog="oneresume.py")

        allowed_filetypes = ['docx', 'mako']

        p.add_argument('-t', '--template-file', required=True, type=argparse.FileType('r'),
             dest='template_file', help='Template filename %s' % allowed_filetypes)

        p.add_argument('-y', '--yaml-resume-file', required=True, type=argparse.FileType('r'),
             dest='resume_file', help='Resume filename')

        p.add_argument('-d', '--debug', action='store_true',
            default=False, dest='debug', help='Turn on debugging')

        p.add_argument('-v', '--verbose', action='store_true',
            default=False, dest='verbose', help='Turn on verbose mode')

        p.add_argument('-s', '--skip-substitution', action='store_true',
            default=False, dest='skip', help='Skip the text substitution and just write out the template as is (useful for pretty-printing')


        args = p.parse_args(argv)

        self.debug = args.debug
        self.verbose = args.verbose
        self.skip = args.skip

        if args.debug:
            logging.basicConfig(level=logging.DEBUG, format='%(message)s')

        if args.verbose:
            logging.basicConfig(level=logging.INFO, format='%(message)s')

        # Normal options
        self.template_file = args.template_file
        logging.debug("template filename is %s" % (self.template_file.name)) 
        filebasename,filetype = os.path.splitext(self.template_file.name)
        if filetype[1:] not in allowed_filetypes:
            error("File type/extension %s is not one of following: %s" % (filetype,' '.join(allowed_filetypes)))

        self.resume = yaml.load(args.resume_file)
        args.resume_file.close()
        logging.debug(self.resume)

            

    
    def clean_up_files(self, files):
        for file in files:
            try:
                os.remove(file)
            except:
                logging.info("Error removing file %s .... continuing" % file)

    def go(self, argv):
        # Read the command line options
        self.getOptions(argv)

        #word = WordResume(self.template_file, self.resume, self.skip)
        text = TextResume(self.template_file, self.resume, self.skip)
        text.render("myresume.txt")
        #self.clean_up_files((tiff_filename, hocr_filename))

if __name__ == '__main__':
    script = OneResume()
    script.go(sys.argv[1:])


