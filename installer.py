"""
pyRoK dependencies installer
"""

import ctypes
import os
import subprocess
import sys
import tkinter as tk
import zipfile
import shutil
from tkinter import messagebox
from urllib.request import urlopen

ASADMIN = 'asadmin'


def download(url: str, path_to_file: str):
    """
    downoad a file and save it
    """
    # Download from URL
    with urlopen(url) as file:
        content = file.read()
    # Save to file
    with open(path_to_file, 'wb') as downloaded_file:
        downloaded_file.write(content)


def unzip(zipped_file, dest_directory):
    """
    extract .zip file
    """
    with zipfile.ZipFile(zipped_file, 'r') as zip_ref:
        zip_ref.extractall(dest_directory)


def install():
    """
    installer
    """
    install_command = [sys.executable, "-m", "pip",
                       "install", "--no-input", "-r", "requirements.txt"]
    subprocess.run(install_command, check=False)

    TESSERACT_URL = "https://github.com/Damocles078/tesseract-setup/archive/refs/heads/main.zip"
    TESSDATA_URL = "https://github.com/Damocles078/tessdata/archive/refs/heads/main.zip"
    if os.path.isfile(r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'):
        tessdata = r'C:\\Program Files\\Tesseract-OCR\\tessdata\\'
    elif os.path.isfile(os.getenv("LOCALAPPDATA") + r'\\Programs\\Tesseract-OCR\\tesseract.exe'):
        tessdata = os.getenv("LOCALAPPDATA") + \
            r'\\Programs\\Tesseract-OCR\\tessdata\\'
    else:
        download(TESSERACT_URL, "./tesseract.zip")
        unzip("./tesseract.zip", "./")
        os.remove("./tesseract.zip")
        subprocess.run(
            "./tesseract-setup-main/tesseract-ocr-w64-setup-v5.1.0.20220510.exe", check=False)
        shutil.rmtree("./tesseract-setup-main")
        if os.path.isfile(r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'):
            tessdata = r'C:\\Program Files\\Tesseract-OCR\\tessdata\\'
        elif os.path.isfile(os.getenv("LOCALAPPDATA") + r'\\Programs\\Tesseract-OCR\\tesseract.exe'):
            tessdata = os.getenv("LOCALAPPDATA") + \
                r'\\Programs\\Tesseract-OCR\\tessdata\\'
        else:
            messagebox.showerror(
                "pyRoK", "Tesseract not installed properly, install it without changing default installation folder. Please uninstall your tesseract and restart this script")
            sys.exit(1)
    download(TESSDATA_URL, "./tessdata.zip")
    unzip("./tessdata.zip", "./")
    os.remove("./tessdata.zip")
    present_files = os.listdir(tessdata)
    for file in os.listdir("./tessdata-main"):
        if file in present_files:
            os.remove(tessdata+file)
        shutil.copy("./tessdata-main/" + file, tessdata + file)
    shutil.rmtree("./tessdata-main")


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    if ctypes.windll.shell32.IsUserAnAdmin() != 1:
        messagebox.showerror("pyRoK", "This script will restart as admin")
        # restart as admin
        ctypes.windll.shell32.ShellExecuteW(
            None, u"runas", sys.executable, __file__, None, 1)
    else:
        install()
        messagebox.showinfo("pyRoK", "installation process done")
    sys.exit(0)
