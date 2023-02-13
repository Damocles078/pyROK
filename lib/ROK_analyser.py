import datetime
import queue
import os
import sqlite3 as sq
import threading
import time
import unicodedata  # import unicode converter
from typing import ClassVar  # import time to name the .csv file

import cv2  # import openCV
import numpy as np
import pytesseract

from lib.ROK_GUI import pyROK_GUI  # import tesseract tools for Python

# set path to tesseract.exe
if os.path.isfile(r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'):
    pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
elif os.path.isfile(os.getenv("LOCALAPPDATA") + r'\\Programs\\Tesseract-OCR\\tesseract.exe'):
    pytesseract.pytesseract.tesseract_cmd = os.getenv("LOCALAPPDATA") + r'\\Programs\\Tesseract-OCR\\tesseract.exe'
else:
    raise RuntimeError("Tesseract-OCR is not installed in expected folders")

class ROKanalyzer():

    # setting considered languages for nicknames
    custom_config_names: ClassVar[str] = r'-l chi_sim+eng+rus+jpn+kor --oem 3 --psm 7'
    custom_config_numbers: ClassVar[str] = r'-l eng -c tessedit_char_whitelist=0123456789, --oem 3'
    custom_config_IDs: ClassVar[str] = r'-l eng -c tessedit_char_whitelist=0123456789), --oem 2'
    custom_config_alliances: ClassVar[str] = r'-l eng --oem 3 --psm 7'

    def __init__(self,
                 GUI: pyROK_GUI = None,
                 database: str = "rok.db",
                 request: str = "",
                 acquire_name: bool = True,
                 acquire_alliance: bool = True,
                 acquire_power: bool = True,
                 acquire_T4: bool = True,
                 acquire_T5: bool = True,
                 acquire_dead: bool = True,
                 acquire_rssAssist: bool = True) -> None:

        self.gui = GUI
        self.database: str = database
        self.acquire_name: bool = acquire_name
        self.acquire_alliance: bool = acquire_alliance
        self.acquire_power: bool = acquire_power
        self.acquire_T4: bool = acquire_T4
        self.acquire_T5: bool = acquire_T5
        self.acquire_dead: bool = acquire_dead
        self.acquire_rssAssist: bool = acquire_rssAssist
        self.queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._stop_event: threading.Event = threading.Event()
        self.processed: int = 0
        self.time: str = str(
            datetime.datetime.now().strftime("%Y_%m_%d__%H_%M_%S"))
        self.request: str = request
        self.worker_thread: threading.Thread = threading.Thread(
            target=self.worker).start()

    def worker(self) -> None:
        connection = sq.connect(self.database)
        while(not self._stop_event.is_set() and not self.gui.close.is_set()):
            if(not self.queue.qsize() == 0):
                current_player = self.queue.get()
                image1: cv2.Mat = cv2.imread(current_player[0])
                image2: cv2.Mat = cv2.imread(current_player[1])

                id = self.acquire_ID_from_image(image1[360:385, 940:1095])
                try:
                    if self.acquire_name:
                        name = self.acquire_name_from_img(
                            image2[275:315, 620:860])
                    else:
                        name = "Not Acquired"
                except:
                    name = "failed"

                try:
                    if self.acquire_power:
                        power = self.acquire_power_from_img(
                            image2[270:315, 980:1150])
                    else:
                        power = -1
                except:
                    power = -2

                try:
                    if self.acquire_alliance:
                        alliance = self.acquire_alliance_from_img(
                            image1[465:490, 820:920])
                    else:
                        alliance = "Not Acquired"
                except:
                    alliance = "failed"

                try:
                    if self.acquire_T4:
                        t4 = self.acquire_T4_from_img(
                            image1[660:685, 1010:1220])
                    else:
                        t4 = -1
                except:
                    t4 = -2

                try:
                    if self.acquire_T5:
                        t5 = self.acquire_T5_from_img(
                            image1[700:725, 1010:1220])
                    else:
                        t5 = -1
                except:
                    t5 = -2

                try:
                    if self.acquire_dead:
                        dead = self.acquire_dead_from_img(
                            image2[535:565, 1200:1400])
                    else:
                        dead = -1
                except:
                    dead = -2

                try:
                    if self.acquire_rssAssist:
                        rssAssist = self.acquire_rssAssist_from_img(
                            image2[725:760, 1200:1400])
                    else:
                        rssAssist = -1
                except:
                    rssAssist = -2

                with connection:
                    mycursor = connection.cursor()
                    mycursor.execute(
                        self.request, (self.time, id, name, alliance, power, t4, t5, dead, rssAssist))
                    connection.commit()

                self.processed += 1
                self.gui.updateBar()

            else:
                time.sleep(2)

    def stop_worker_thread(self):
        self._stop_event.set()

    def add_player(self, image1: str, image2: str) -> None:
        self.queue.put((image1, image2))

    def acquire_ID_from_image(self, image: cv2.Mat) -> int:
        cleaned_image = traitement(image)
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_IDs)
        text = text.replace(')', '').replace('}', '').replace(']', '')
        text = text.split()
        res = ''
        for t in text:
            res += t
        return int(res)

    def acquire_name_from_img(self, image: cv2.Mat) -> str:
        cleaned_image = traitement(image)
        # identify text from image using more lannguages than default
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_names)
        res = text.split()
        fin = ''
        try:
            fin = res[0]
        except:
            1 == 1
        for k in range(len(res)-1):
            fin = fin + ' ' + res[k+1]
        return unicodedata.normalize("NFKC", fin)

    def acquire_alliance_from_img(self, image: cv2.Mat) -> str:
        cleaned_image = traitement(image)
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_alliances)
        text = text.split("]")[0].split("[")[-1]
        if "\n\x0c" not in text and len(text) < 5 and len(text) > 2:
            result = str(text)
        else:
            result = ""
        return result

    def acquire_power_from_img(self, image: cv2.Mat) -> int:
        cleaned_image = traitement(image)
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_numbers)
        text = text.replace(',', '').replace('.', '').replace(':', '')
        text = text.split()
        res = ''
        for t in text:
            res += t
        try:
            result = int(res)
        except:
            result = res[0]
        return result

    def acquire_T4_from_img(self, image: cv2.Mat) -> int:
        cleaned_image = traitement(image, use_bitwise_not=True)
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_numbers)
        text = text.replace(',', '').replace('.', '')
        text = text.split()
        res = ''
        for t in text:
            res += t
        if res == '':
            res = 0
        return int(res)

    def acquire_T5_from_img(self, image: cv2.Mat) -> int:
        cleaned_image = traitement(image, use_bitwise_not=True)
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_numbers)
        text = text.replace(',', '').replace('.', '')
        text = text.split()
        res = ''
        for t in text:
            res += t
        if res == '':
            res = 0
        return int(res)

    def acquire_dead_from_img(self, image: cv2.Mat) -> int:
        cleaned_image = traitement(image)
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_numbers)
        text = text.replace(',', '').replace('.', '')
        text = text.split()
        res = ''
        for t in text:
            res += t
        if res == '':
            res = 0
        return int(res)

    def acquire_rssAssist_from_img(self, image: cv2.Mat) -> int:
        cleaned_image = traitement(image)
        text: str = pytesseract.image_to_string(
            cleaned_image, config=self.custom_config_numbers)
        text = text.replace(',', '').replace('.', '')
        text = text.split()
        res = ''
        for t in text:
            res += t
        if res == '':
            res = 0
        return int(res)


def traitement(img: cv2.Mat, use_bitwise_not: bool = False) -> cv2.Mat:
    temp = cv2.bitwise_not(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
    result = cv2.threshold(temp, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)[1]
    if use_bitwise_not:
        result = cv2.bitwise_not(result)
    return(result)


def main():
    ROKanalyzer()


if __name__ == '__main__':
    main()
