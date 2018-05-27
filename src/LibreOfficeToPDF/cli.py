import subprocess
import os
import click
import sys


OPENOFFICE_PATH    = os.environ["LIBREOFFICE_PROGRAM"]
OPENOFFICE_PYTHON  = os.path.join(OPENOFFICE_PATH, 'python')
print("\nExecuting LibreOfficeToPDF")
print("LibreOffice path: {}".format(OPENOFFICE_PATH))
print("LibreOffice python: {}".format(OPENOFFICE_PYTHON))

script_dir = None
frozen = getattr(sys, 'frozen', False)
if frozen:
    # running in a bundle
    meipass = sys._MEIPASS
    script_dir = meipass
else :
    # running live
    meipass = "None" # Needs to be any string
    script_dir = os.path.dirname(os.path.realpath(__file__))

cwd = os.getcwd()
print("Current working directory: {}".format(cwd))
script = os.path.join(script_dir, "script.py")


@click.command()
@click.argument('source')
@click.option('--update/--no-update', default=True, help='Update indexes (e.g. Table of Contents) on the source file and save it.')
@click.option('--pdf/--no-pdf', default=True, help='Generate a PDF file with same name as source.')
def main(source, update, pdf):
    """source: Path (relative or absolute) to docx file."""
    if not os.path.isabs(source):
        source = os.path.join(cwd, source)
    print("Calling LibreOffice python\n")
    proc = subprocess.run([str(OPENOFFICE_PYTHON), script, meipass, source, str(update), str(pdf)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    print(proc.stdout.decode())
    sys.exit(proc.returncode)

if __name__ == "__main__":
    main()
