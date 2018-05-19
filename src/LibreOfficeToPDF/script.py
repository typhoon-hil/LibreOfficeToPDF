"""This script is passed to LibreOffice python interpreter to be executed."""
import socket
import uno
import sys
import os
import time
import atexit

OPENOFFICE_PORT = 8100 # 2002

OPENOFFICE_PATH    = os.environ["LIBREOFFICE_PROGRAM"]
OPENOFFICE_BIN     = os.path.join(OPENOFFICE_PATH, 'soffice')

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

        if not desktop:
            raise Exception("Failed to create OpenOffice desktop on port %d" % self.port)

        if did_start:
            _started_desktops[self.port] = desktop

        return desktop


    def startup(self):
        """
        Start a headless instance of OpenOffice.
        """
        args = [OPENOFFICE_BIN,
                '-accept=socket,host=localhost,port=%d;urp;StarOffice.ServiceManager' % self.port,
                '-norestore',
                '-nofirststartwizard',
                '-nologo',
                '-headless',
                ]

        try:
            pid = os.spawnve(os.P_NOWAIT, args[0], args, env)
        except Exception as e:
            raise Exception("Failed to start OpenOffice on port %d: %s" % (self.port, e.message))

        if pid <= 0:
            raise Exception("Failed to start OpenOffice on port %d" % self.port)


    def shutdown(self):
        """
        Shutdown OpenOffice.
        """
        try:
            if _started_desktops.get(self.port):
                _started_desktops[self.port].terminate()
                del _started_desktops[self.port]
        except Exception as e:
            pass

# Keep track of started desktops and shut them down on exit.
_started_desktops = {}

def _shutdown_desktops():
    """ Shutdown all OpenOffice desktops that were started by the program. """
    for port, desktop in _started_desktops.items():
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


def run(source, update_and_save=True, pdf=True):
    fileUrl = uno.systemPathToFileUrl(os.path.realpath(source))

    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    document = desktop.loadComponentFromURL(fileUrl, "_default", 0, ())

    model = desktop.getCurrentComponent()
    dispatcher = smgr.createInstanceWithContext( "com.sun.star.frame.DispatchHelper", ctx)
    doc = model.getCurrentController()

    if update_and_save:
        dispatcher.executeDispatch(doc, ".uno:UpdateAllIndexes", "", 0, ())
        struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        struct.Name = 'URL'
        struct.Value = 'file:///{}'.format(source.replace("\\","/").replace("report","report2"))
        dispatcher.executeDispatch(doc, ".uno:SaveAs", "", 0, tuple([struct]))

    if pdf:
        pass


if __name__ == "__main__":
    run(sys.argv[1], sys.argv[2], sys.argv[3])
