"""This script is passed to LibreOffice python interpreter to be executed."""
import sys
import os
import subprocess
import time
import atexit
import socket
import uno
import argparse


print("Executing LibreOffice python script using LibreOffice python")
OPENOFFICE_PORT = 8100 # 2002

OPENOFFICE_PATH    = os.environ["LIBREOFFICE_PROGRAM"]
OPENOFFICE_BIN     = os.path.join(OPENOFFICE_PATH, 'soffice')

NoConnectException = uno.getClass("com.sun.star.connection.NoConnectException")
PropertyValue = uno.getClass("com.sun.star.beans.PropertyValue")


# Adapted from: https://www.linuxjournal.com/content/starting-stopping-and-connecting-openoffice-python
class OORunner:
    """
    Start, stop, and connect to OpenOffice.
    """
    def __init__(self, port=OPENOFFICE_PORT):
        """ Create OORunner that connects on the specified port. """
        self.port = port


    def connect(self, no_startup=False):
        """
        Connect to OpenOffice.
        If a connection cannot be established try to start OpenOffice.
        """
        print("Connecting to LibreOffice at port {}".format(self.port))
        localContext = uno.getComponentContext()
        resolver     = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        context      = None
        did_start    = False

        n = 0
        while n < 6:
            try:
                context = resolver.resolve("uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % self.port)
                break
            except NoConnectException:
                pass

            # If first connect failed then try starting OpenOffice.
            if n == 0:
                print("Failed to connect. Trying to start LibreOffice instance.")
                # Exit loop if startup not desired.
                if no_startup:
                     break
                self.startup()
                did_start = True

            # Pause and try again to connect
            time.sleep(1)
            n += 1

        if not context:
            raise Exception("Failed to connect to LibreOffice on port %d" % self.port)

        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        dispatcher = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", context)

        if not desktop:
            raise Exception("Failed to create OpenOffice desktop on port %d" % self.port)

        if did_start:
            _started_desktops[self.port] = desktop

        return desktop, dispatcher


    def startup(self):
        """
        Start a headless instance of OpenOffice.
        """
        print("Starting headless LibreOffice")

        try:
            pid = subprocess.Popen([OPENOFFICE_BIN, '--norestore', '--nofirststartwizard', '--nologo', '--headless', '--invisible', '--accept=socket,host=localhost,port=%d;urp;' % self.port]).pid
        except Exception as e:
            raise Exception("Failed to start OpenOffice on port %d: %s" % (self.port, e.message))

        if pid <= 0:
            raise Exception("Failed to start OpenOffice on port %d" % self.port)

        print("LibreOffice started")


    def shutdown(self):
        """
        Shutdown OpenOffice.
        """
        print("Shutting down LibreOffice")
        try:
            if _started_desktops.get(self.port):
                print("Terminating instance at port {}".format(self.port))
                _started_desktops[self.port].terminate()
                del _started_desktops[self.port]
        except Exception as e:
            pass

# Keep track of started desktops and shut them down on exit.
_started_desktops = {}

def _shutdown_desktops():
    """ Shutdown all OpenOffice desktops that were started by the program. """
    for port, desktop in _started_desktops.items():
        print("Exit: Found instance found at port {}. Terminating...".format(port))
        try:
            if desktop:
                desktop.terminate()
        except Exception as e:
            pass


atexit.register(_shutdown_desktops)


def oo_shutdown_if_running(port=OPENOFFICE_PORT):
    """ Shutdown OpenOffice if it's running on the specified port. """
    oorunner = OORunner(port)
    try:
        desktop = oorunner.connect(no_startup=True)
        desktop.terminate()
    except Exception as e:
        pass


def run(source, update, pdf):
    fileUrl = uno.systemPathToFileUrl(os.path.realpath(source))
    filepath, ext = os.path.splitext(source)
    fileUrlPDF = uno.systemPathToFileUrl(os.path.realpath(filepath+".pdf"))

    runner = OORunner(2002)
    desktop, dispatcher = runner.connect()

    print("Loading document")
    struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    struct.Name = 'Hidden'
    struct.Value = True
    document = desktop.loadComponentFromURL(fileUrl, "_default", 0, ([struct]))
    doc = document.getCurrentController()
    #doc = desktop.getCurrentComponent().getCurrentController()

    if update:
        print("Updating Indexes and Saving")
        dispatcher.executeDispatch(doc, ".uno:UpdateAllIndexes", "", 0, ())
        struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        struct.Name = 'URL'
        struct.Value = fileUrl
        dispatcher.executeDispatch(doc, ".uno:SaveAs", "", 0, tuple([struct]))

    if pdf:
        print("Generating PDF")
        struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        struct.Name = 'URL'
        struct.Value = fileUrlPDF
        struct2 = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        struct2.Name = "FilterName"
        struct2.Value = "writer_pdf_Export"
        dispatcher.executeDispatch(doc, ".uno:ExportDirectToPDF", "", 0, tuple([struct, struct2]))

    runner.shutdown()


def main():

    parser = argparse.ArgumentParser(description='Run LibreOffice to update indexes in a document and convert to PDF')
    parser.add_argument('source', metavar='DOC_PATH',
                        help='Path to the source document file to be converted')
    parser.add_argument('--no-update', dest='update', action='store_false',
                        help="Do not update indexes (e.g. Table of Contents) on the source file and save it")
    parser.add_argument('--no-pdf', dest='pdf', action='store_false',
                        help='Do not generate a PDF file with same name as source')

    args = parser.parse_args()
    print(args)
    run(args.source, args.update, args.pdf)


if __name__ == "__main__":
    main()

