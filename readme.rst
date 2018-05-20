LibreOfficeToPDF
================
This is a very simple utility that uses LibreOffice to do two things:

- Update the TOC of a given ``docx`` file (especially useful when creating ``docx`` files using `python-docx`_ library)
- Generate a PDF

So far tested only in Windows 10 using Python 3.6, but it should work on other platforms and Python versions as well. Users are more than welcome to test it.

Also not tested for other text file formats. It might work.

.. _python-docx: https://github.com/python-openxml/python-docx

Installation
------------
For the moment it only works with a python installation:

- Clone the repository, go to project folder (where ``setup.py`` is) and install with ``pip install .``
- Create a environment variable called ``LIBREOFFICE_PROGRAM`` with the path to LibreOffice folder where ``soffice`` and ``python`` executables are present.

A binary executable file independent of python is being worked on, stay tuned.

Usage
-----
If your python and scripts folder are in path, then you can access from the command line directly.

Typical usage (opens file, updates indexes, saves, generates pdf):
``LibreOfficeToPDF C:\Users\john\file.docx``

For more options see:
``LibreOfficeToPDF --help``
