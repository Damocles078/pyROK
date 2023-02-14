"""
pyRoK dependencies installer
"""

import ctypes
import io
import os
import pathlib
import shutil
import subprocess
import sys
import tkinter as tk
import urllib.request
import zipfile
from tkinter import messagebox

ASADMIN = 'asadmin'
TESSERACT_URL = "https://github.com/Damocles078/tesseract-setup/archive/refs/heads/main.zip"
TESSDATA_URL = "https://github.com/Damocles078/tessdata/archive/refs/heads/main.zip"

os.chdir(os.path.dirname(__file__))

def download_zip_and_extract(url, target_dir):
    print(f"Starting download from {url}")
    buffer_all = io.BytesIO()
    buffer_all_size = 0
    with urllib.request.urlopen(url) as res:
        length = res.headers.get('Content-length')
        if length:
            length = int(length)
            fragment_size = max(4096, length // 20)
            print(f"Detected size : {length}")
            while True:
                temp_buffer = res.read(fragment_size)
                if not temp_buffer:
                    break
                buffer_all.write(temp_buffer)
                buffer_all_size += len(temp_buffer)
                if length:
                    percent = int((buffer_all_size / length)*100)
                    sys.stdout.write(
                        f"\r[{'=' * (percent//2)}{' ' * (50-percent//2)}] {percent}% done")
                    sys.stdout.flush()
        else:
            print("Size not detected, you will not have progress information")
            buffer_all.write(res.read())
        print("\nDownload done, unzipping")
        with zipfile.ZipFile(buffer_all, 'r') as zip_ref:
            zip_ref.extractall(target_dir)
        print("Unzipping done")

def create_shortcut():
    
    def escape_path(path):
        return str(path).replace('\\', '/')

    shortcut_path = escape_path(pathlib.Path(os.path.expanduser("~/Desktop/pyRoK.lnk")))
    script_powershell = escape_path(os.path.realpath("./shortcut.ps1"))
    pyRoK_path = escape_path(os.path.dirname(__file__) + "/pyROK.py")
    shortcut_working_directory = escape_path(os.path.dirname(__file__))
    icon = escape_path(os.path.realpath("./lib/pyROK.ico"))
    temp = sys.executable.split("\\")
    temp[-1] = temp[-1].replace("python", "pythonw")
    executable = "/".join(temp)
    subprocess.run(["powershell", script_powershell, shortcut_path, executable, pyRoK_path, shortcut_working_directory, icon], check=False)


def install():
    """
    installer
    """
    print("\nInstalling python packages\n")
    install_command = [sys.executable, "-m", "pip",
                       "install", "--no-input", "-r", "requirements.txt"]
    subprocess.run(install_command, check=False)
    print("\nPython packages installation done\n")
    print("Searching for Tesseract-OCR")
    # try to locate tesseract
    if os.path.isfile(r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'):
        tessdata = r'C:\\Program Files\\Tesseract-OCR\\tessdata\\'
        print("Tesseract-OCR found at : 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'")
    elif os.path.isfile(os.getenv("LOCALAPPDATA") + r'\\Programs\\Tesseract-OCR\\tesseract.exe'):
        tessdata = os.getenv("LOCALAPPDATA") + \
            r'\\Programs\\Tesseract-OCR\\tessdata\\'
        print("Tesseract-OCR found at : %s", os.getenv("LOCALAPPDATA") + '\\Programs\\Tesseract-OCR\\tesseract.exe')
    else:
        # tesseract not found, downloading and installing
        print("Tesseract-OCR not found, downloading installer")
        download_zip_and_extract(TESSERACT_URL, "./")
        print("Running installer, do not change default Tesseract-OCR installation path")
        subprocess.run(
            "./tesseract-setup-main/tesseract-ocr-w64-setup-v5.1.0.20220510.exe", check=False)
        shutil.rmtree("./tesseract-setup-main")
        print("Searching for Tesseract-OCR")
        if os.path.isfile(r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'):
            tessdata = r'C:\\Program Files\\Tesseract-OCR\\tessdata\\'
            print("Tesseract-OCR found at : 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'")
        elif os.path.isfile(os.getenv("LOCALAPPDATA") + r'\\Programs\\Tesseract-OCR\\tesseract.exe'):
            tessdata = os.getenv("LOCALAPPDATA") + \
                r'\\Programs\\Tesseract-OCR\\tessdata\\'
            print("Tesseract-OCR found at : %s", os.getenv("LOCALAPPDATA") + '\\Programs\\Tesseract-OCR\\tesseract.exe')
        else:
            messagebox.showerror(
                "pyRoK", "Tesseract not installed properly, install it without changing default installation folder. Please uninstall your tesseract and restart this script")
            sys.exit(1)
    print("Downloading Tesseract model data")
    download_zip_and_extract(TESSDATA_URL, "./")
    print("Installing Tesseract model data")
    present_files = os.listdir(tessdata)
    for file in os.listdir("./tessdata-main"):
        if file in present_files:
            os.remove(tessdata+file)
        shutil.copy("./tessdata-main/" + file, tessdata + file)
    shutil.rmtree("./tessdata-main")
    print("creating pyROK shortcut on the Desktop")
    create_shortcut()
    print("done creating pyROK shortcut")


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    if ctypes.windll.shell32.IsUserAnAdmin() != 1:
        messagebox.showerror("pyRoK", "This script will restart as admin")
        # restart as admin
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, __file__, None, 3)
    else:
        try:
            install()
            messagebox.showinfo("pyRoK", "installation process done")
        except:
            messagebox.showerror("pyRoK", "Something failed during installation")
    sys.exit(0)
