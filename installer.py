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

TESSERACT_URL = "https://github.com/Damocles078/tesseract-setup/archive/refs/heads/main.zip"
TESSDATA_URL = "https://github.com/Damocles078/tessdata/archive/refs/heads/main.zip"
EXPECTED_SIZE = {
    "ara.traineddata": 2494806,
    "chi_sim.traineddata": 44366093,
    "eng.traineddata": 23466654,
    "jpn.traineddata": 35659159,
    "kor.traineddata": 15317715,
    "rus.traineddata": 19920885,
}

os.chdir(os.path.dirname(__file__))

def sizeof_fmt(num, suffix="B"):
    for unit in ["", "Ki", "Mi", "Gi"]:
        if abs(num) < 1024.0:
            return f"{num:.1f} {unit}{suffix}"
        num /= 1024.0
    return f"{num:.1f} Ti{suffix}"

def download_zip_and_extract(url, target_dir):
    print(f"Starting download from {url}")
    buffer_all = io.BytesIO()
    buffer_all_size = 0
    request = urllib.request.Request(url=url)
    request.add_header('Accept-Encoding', '')
    request.add_header('Content-Encoding','gzip')
    with urllib.request.urlopen(request) as res:
        length = res.headers.get('Content-length')
        if length:
            length = int(length)
            fragment_size = max(4096, length // 20)
            print(f"Detected size : {sizeof_fmt(length)}")
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

    shortcut_path = escape_path(pathlib.Path(os.path.expanduser("~/Desktop/pyROK.lnk")))
    script_powershell = escape_path(os.path.realpath("./shortcut.ps1"))
    pyRoK_path = escape_path(os.path.dirname(__file__) + "/pyROK.py")
    shortcut_working_directory = escape_path(os.path.dirname(__file__))
    icon = escape_path(os.path.realpath("./lib/pyROK.ico"))
    temp = sys.executable.split("\\")
    temp[-1] = temp[-1].replace("python", "pythonw")
    executable = "/".join(temp)
    subprocess.run(["powershell", script_powershell, shortcut_path, executable, pyRoK_path, shortcut_working_directory, icon], check=False)


def find_tessdata_folder():
    tessdata = None
    if os.path.isfile(r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'):
        tessdata = r'C:\\Program Files\\Tesseract-OCR\\tessdata\\'
    elif os.path.isfile(os.getenv("LOCALAPPDATA") + r'\\Programs\\Tesseract-OCR\\tesseract.exe'):
        tessdata = os.getenv("LOCALAPPDATA") + \
            r'\\Programs\\Tesseract-OCR\\tessdata\\'
    return tessdata

def need_download(tessdata):
    need_dl = False
    for file, size in EXPECTED_SIZE.items():
        if not os.path.isfile(tessdata + file):
            print(file, " not found, will download all files because of it")
            need_dl = True
            continue
        if os.stat(tessdata + file).st_size != size:
            print(file, " incorrect size, will download all files because of it")
            need_dl = True
    return need_dl

def install_tesseract():
    download_zip_and_extract(TESSERACT_URL, "./")
    print("Running installer, do not change default Tesseract-OCR installation path")
    subprocess.run(
        "./tesseract-setup-main/tesseract-ocr-w64-setup-v5.1.0.20220510.exe", check=False)
    shutil.rmtree("./tesseract-setup-main")

def install_tessdata(tessdata):
    download_zip_and_extract(TESSDATA_URL, "./")
    print("Installing Tesseract model data")
    present_files = os.listdir(tessdata)
    for file in os.listdir("./tessdata-main"):
        if file in present_files:
            os.remove(tessdata+file)
        shutil.copy("./tessdata-main/" + file, tessdata + file)
    shutil.rmtree("./tessdata-main")

def install():
    """
    installer
    """
    print("\nVerifying python packages\n")
    install_command = [sys.executable, "-m", "pip",
                       "install", "--no-input", "-r", "requirements.txt"]
    subprocess.run(install_command, check=False)
    print("\nPython packages verification done\n")
    print("Searching for Tesseract-OCR")
    # try to locate tesseract
    tessdata = find_tessdata_folder()
    if tessdata is None:
        # tesseract not found, downloading and installing
        print("Tesseract-OCR not found, downloading installer")
        install_tesseract()
        print("Searching for Tesseract-OCR")
        tessdata = find_tessdata_folder()
        if tessdata is None:
            messagebox.showerror(
                "pyRoK", "Tesseract not installed properly, install it without changing default installation folder. Please uninstall your tesseract and restart this script")
            sys.exit(1)
    else:
        print("Tesseract-OCR found, continuing")
    if need_download(tessdata):
        print("Downloading Tesseract model data")
        install_tessdata(tessdata)
        print("Done installing tessdata")
    else:
        print("Tessdata size matches expected, assuming files are correct, continuing")
    if os.path.isfile(os.path.expanduser("~/Desktop/pyROK.lnk")):
        print("Updating pyROK shortcut on the desktop")
    else:
        print("creating pyROK shortcut on the desktop")
    create_shortcut()
    print("Shortcut setup done")


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
