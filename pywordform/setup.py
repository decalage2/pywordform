# Setup script for pywordform - Philippe Lagadec

# History:
# 2012-03-04 v0.01 PL: - first version

import distutils.core

from pywordform import __version__

kw = {
    'name': "pywordform",
    'version': __version__,
    'description': "a python module to parse a Microsoft Word form in docx format, and extract all field values with their tags into a dictionary.",
    'author': 'Philippe Lagadec',
    #'author_email': "decalage(a)laposte.net",
    'url': "http://www.decalage.info/python/pywordform",
    'license': "BSD (see source code or LICENCE.txt)",
    'py_modules': ['pywordform'],
    }


# If we're running Python 2.3, add extra information
if hasattr(distutils.core, 'setup_keywords'):
    if 'classifiers' in distutils.core.setup_keywords:
        kw['classifiers'] = [
            'Development Status :: 4 - Beta',
            'License :: OSI Approved :: BSD License',
            'Intended Audience :: Developers',
            'Operating System :: OS Independent',
            'Programming Language :: Python',
            'Topic :: Software Development :: Libraries :: Python Modules'
          ]
    if 'download_url' in distutils.core.setup_keywords:
        kw['download_url'] = "https://bitbucket.org/decalage/pywordform/downloads"

distutils.core.setup(**kw)
