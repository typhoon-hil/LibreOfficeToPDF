from subprocess import check_output
import os
import click
import sys

OPENOFFICE_PATH    = os.environ["LIBREOFFICE_PROGRAM"]
OPENOFFICE_PYTHON  = os.path.join(OPENOFFICE_PATH, 'python')
print("LibreOffice path: {}".format(OPENOFFICE_PATH))
print("LibreOffice python: {}".format(OPENOFFICE_PYTHON))

script_dir = None
if getattr( sys, 'frozen', False ) :
    # running in a bundle
    script_dir = sys._MEIPASS
    print("Running from Executable")
else :
    # running live
    script_dir = os.path.dirname(os.path.realpath(__file__))
    print("Running from normal python installation")

cwd = os.getcwd()
print("Current working directory: {}".format(cwd))
script = os.path.join(script_dir, "script.py")
print("script path: {}".format(script))

@click.command()
@click.argument('source')
@click.option('--update_and_save', is_flag=True, default=True, help='Update indexes (e.g. Table of Contents) on the source file and save. (True or False)')
@click.option('--pdf', is_flag=True, default=True, help='Save a PDF file with same name as source. (True or False)')
def main(source, update_and_save, pdf):
    """source: Path (relative or absolute) to docx file."""
    if not os.path.isabs(source):
        source = os.path.join(cwd, source)
    print("Calling LibreOffice python\n")
    print(check_output(["{}".format(OPENOFFICE_PYTHON), script, source], shell=True).decode())

if __name__ == "__main__":
    main()
