pywordform:
a python module to parse a Microsoft Word form in docx format, and
extract all field values with their tags into a dictionary.

Project website: http://www.decalage.info/python/pywordform



INSTALLATION:

- on Windows, launch install.bat
- on other systems, launch: setup.py install



HOW TO USE THIS MODULE:

Open sample_form.docx in MS Word, and edit field values.

From the shell, extract all fields with tags:

> python pywordform.py sample_form.docx
field1 = "hello, world."
field2 = "hello,"
field3 = "value B"
field4 = "04-03-2012"

In a python script:

import pywordform
fields = pywordform.parse_form('sample_form.docx')
print fields

=> this returns a dictionary of field values indexed by tags.

See http://www.decalage.info/python/pywordform
See main program at the end of the module, and also docstrings.



LICENSE:

See LICENSE.txt.
