OneResume
=============

This script can take a Resume (or any structured text, really) in a YAML file
and generate the following versions so far:

    - Word .docx (an example template file is included)
    - Plain text (an example mako file is included)

The plain text plugin currently uses the mako templating engine, while the word version
uses a simple made-up syntax.

Usage: 
------
    python oneresume.py single -y resume.yaml -t template.docx -o myresume_output.docx -f Word

    --> myresume_output.docx will be generated
    
Caveats
-------
This code is brand-new, and is barely commented with no unit-tests included.  I plan to improve 
things as time allows in the near-future.

Also, the word plugin is a little simplistic, so don't try to use tables or anything fancy like that
in your resume.  But if you stick to simple text flow and rely on styles for formatting for the most
part, you should be fine

Dependencies:
------------ 

    pip install lxml
    pip install yaml

