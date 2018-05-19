import subprocess
import os
import time


def open_libreoffice():
    print("Starting LibreOffice...")
    libreoffice_dir = os.environ["LIBREOFFICE_PROGRAM"]
    p = subprocess.Popen(["soffice.exe",
                          '--writer --accept="socket,host=localhost,port=2002;urp;"',
                          libreoffice_dir])
    print("Waiting 10 seconds...")
    time.sleep(10)
    return p

def close_libreoffice(p):
    print("Terminating LibreOffice...")
    p.terminate()
    return_code = p.wait()
