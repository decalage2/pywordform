"""
pywordform

a python module to parse a Microsoft Word form in docx format, and
extract all field values with their tags into a dictionary.

Project website: http://www.decalage.info/python/pywordform

Copyright (c) 2012, Philippe Lagadec (http://www.decalage.info)
All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

 * Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
 * Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

__version__ = '0.01'

#------------------------------------------------------------------------------
# CHANGELOG:
# 2012-02-17 v0.01 PL: - first version

#------------------------------------------------------------------------------
#TODO:
# + extract multiline text fields using <w:t> and <w:br> tags
# - recognize date fields and extract fulldate
# - test docm files, add support if needed
# - support legacy fields
# - CSV output (option)
# - more advanced parser returning a list of field objects: keep order, and
#   get fields with no tag, extract other attributes such as title

#------------------------------------------------------------------------------
import zipfile, sys

try:
    # lxml: best performance for XML processing
    import lxml.etree as ET
except ImportError:
    try:
        # Python 2.5+: batteries included
        import xml.etree.cElementTree as ET
    except ImportError:
        try:
            # Python <2.5: standalone ElementTree install
            import elementtree.cElementTree as ET
        except ImportError:
            raise ImportError, "lxml or ElementTree are not installed, "\
                +"see http://codespeak.net/lxml "\
                +"or http://effbot.org/zone/element-index.htm"


# namespace for word XML tags:
NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# XML tags in Word forms:
TAG_FIELD         = NS_W+'sdt'
TAG_FIELDPROP     = NS_W+'sdtPr'
TAG_FIELDTAG      = NS_W+'tag'
ATTR_FIELDTAGVAL  = NS_W+'val'
TAG_FIELD_CONTENT = NS_W+'sdtContent'
TAG_RUN           = NS_W+'r'
TAG_TEXT          = NS_W+'t'


def parse_form(filename):
    fields = {}
    zfile = zipfile.ZipFile(filename)
    form = zfile.read('word/document.xml')
    xmlroot = ET.fromstring(form)
    for field in xmlroot.getiterator(TAG_FIELD):
        field_tag = field.find(TAG_FIELDPROP+'/'+TAG_FIELDTAG)
        if field_tag is not None:
            tag = field_tag.get(ATTR_FIELDTAGVAL, None)
            field_value = field.find(TAG_FIELD_CONTENT+'/'+TAG_RUN+'/'+TAG_TEXT)
            if field_value is not None:
                value = field_value.text
                fields[tag] = value
    zfile.close()
    return fields


if __name__ == '__main__':
    fields = parse_form(sys.argv[1])
    for tag, value in sorted(fields.items()):
        print '%s = "%s"' % (tag, value)
