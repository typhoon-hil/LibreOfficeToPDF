# LibreOfficeToPDF
This is a very simple utility that uses LibreOffice to do two things:
- Update the TOC of a given `docx` file (especially useful when creating `docx` files using python-docx library)
- Generate a PDF

So far tested only in Windows 10 using Python 3.6, but it should work on other platforms and Python versions as well. Users are more than welcome to test it.

## Installation
- Clone the repository, go to project folder (where `setup.py` is) and install with `pip install .`
- Create a environment variable called `LIBREOFFICE_PROGRAM` with the path to LibreOffice folder where `soffice` and `python` are present.

## Usage
If your python and scripts folder are in path, then you can access from the command line directly:
`LibreOfficeToPDF --help`

You can enable/disable the TOC updates and PDF generation independently.