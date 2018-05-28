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

- Install LibreOffice
- Create a environment variable called ``LIBREOFFICE_PROGRAM`` with the path to LibreOffice folder where ``soffice`` and ``python`` executables are present.

Then use one of the two methods to use LibreOfficeToPDF:

Downloading and using the binaries directly
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
We publish windows standalone executable files. With them you can use it without having to install anything else (no python, etc..).

You can download them at: https://github.com/typhoon-hil/LibreOfficeToPDF/releases

Then you can use the executable directly (possibly adding it a folder added to PATH).

Using with Python
^^^^^^^^^^^^^^^^^
- Install directly from this git repository with ``pip install git+https://github.com/typhoon-hil/LibreOfficeToPDF.git``

The executable file will be present in `Python\Scripts` folder, make sure it is on your PATH.

Usage
-----

If your binary (or python and scripts folder) are in path, then you can access from the command line directly.

Typical usage (opens file, updates indexes, saves, generates pdf):
``LibreOfficeToPDF C:\Users\john\file.docx``

Only update indexes and save:
``LibreOfficeToPDF --no-pdf C:\Users\john\file.docx``

For more options see:
``LibreOfficeToPDF --help``

Building a standalone executable
--------------------------------
We use PyInstaller to create standalone executables. If you want to build an executable yourself, follow these steps:

- Create a new virtual environment with the proper python version (tested using python 3, 32 or 64 bit so far)
- Install using ONLY pip needed packages defined in `setup.py`. This prevents your executable to become too large with unnecessary dependencies
- Delete any previous `dist` folder from a previous PyInstaller run
- Run the `build_exe.cmd` command to run PyInstaller and create a single file executable.
