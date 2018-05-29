LibreOfficeToPDF
================
This is a very simple utility that uses LibreOffice to do two things:

- Update the TOC of a given ``docx`` file (especially useful when creating ``docx`` files using `python-docx`_ library)
- Generate a PDF

Not tested for other text file formats, although itt might work.

.. _python-docx: https://github.com/python-openxml/python-docx

Installation
------------
Requirements:

- LibreOffice (it might also work on OpenOffice)

Installation Steps

- `Download latest release source zip file <https://github.com/typhoon-hil/LibreOfficeToPDF/releases>`_
- Unpack ``src\LibreOfficeToPDF`` folder to anywhere in your system
- Add the LibreOfficeToPDF folder to your PATH.

Windows/Mac OS X only:

- Create a environment variable called ``LIBREOFFICE_PROGRAM`` with the path to LibreOffice folder where ``soffice`` and ``python`` executables are present.

Usage
-----
Typical usage (opens file, updates indexes, saves, generates pdf):
``LibreOfficeToPDF C:\Users\john\file.docx``

Only update indexes and save:
``LibreOfficeToPDF --no-pdf C:\Users\john\file.docx``

For more options see:
``LibreOfficeToPDF --help``
