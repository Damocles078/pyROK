import numpy as np #import basic math package
import time #import time to name the .png file
from PIL import ImageGrab #import screenshot utilities
from lib.ROK_analyser import ROKanalyzer
from lib.ROK_GUI import pyROK_GUI
import os
from scipy.linalg import norm #import some numpy functions
from numpy import sum, average
import win32api, win32con #import input controller
import keyboard #check if key p is pressed to stop program
from screeninfo import Monitor, get_monitors

class ROKpuller():

    def __init__(self, GUI: pyROK_GUI, analyzer: ROKanalyzer) -> None:
        self.GUI: pyROK_GUI = GUI
        self.analyzer: ROKanalyzer = analyzer
        self.to_capture = int(self.GUI.players_to_acquire)
        self.dir = str(self.GUI.ProjectScreenshotsDir)
        monitors: list[Monitor] = get_monitors()
        for monitor in monitors:
            print(monitor)
            if monitor.x == 0 and monitor.y == 0:
                self.x_coef = monitor.width/1920
                self.y_coef = monitor.height/1080
                break
        self.count = 0

    def mousePos(self, x=(0,0)): #setup command to place the cursor on pixel x
        win32api.SetCursorPos(x)
        time.sleep(0.01)

    def leftClick(self): # setup command to left-click
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
        time.sleep(.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
        time.sleep(.01)

    def dragClick(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
        time.sleep(.01)
        for k in range(20):
            self.mousePos((int(860*self.x_coef),int((700-k*4)*self.y_coef)))
            time.sleep(.05)
        time.sleep(0.3)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)


    def screenGrab(self) -> str: #setup command to screenshot
        time.sleep(.1)
        im = ImageGrab.grab() #capture only game 
        im.resize((1920, 1040-29))
        name = self.dir + '\\snap_' + str(int(time.time())) + '.png'
        im.save(name, 'PNG') #save screenshot
        return(name)

    def to_grayscale(self,arr):
        #"If arr is a color image (3D array), convert it to grayscale (2D array)."
        if len(arr.shape) == 3:
            return average(arr, -1)  # average over the last axis (color channels)
        else:
            return arr

    def normalize(self, arr):
        rng = arr.max()-arr.min()
        amin = arr.min()
        return (arr-amin)*255/rng

    def compare_images(self, img1, img2):
        # normalize to compensate for exposure difference, this may be unnecessary
        # consider disabling it
        img1 = self.normalize(img1)
        img2 = self.normalize(img2)
        # calculate the difference and its norms
        diff = img1 - img2  # elementwise for scipy arrays
        m_norm = sum(abs(diff))  # Manhattan norm
        z_norm = norm(diff.ravel(), 0)  # Zero norm
        return (m_norm + z_norm > 500000.0)

    def Opening(self, a): #tries to open a player stats and return if actually opened
        img1 = np.asarray(ImageGrab.grab(bbox=(int(200*self.x_coef), int(600*self.y_coef), int(1600*self.x_coef), int(1000*self.y_coef))))  # capture before click screen
        self.leftClick()
        time.sleep(a)
        img2 = np.asarray(ImageGrab.grab(bbox=(int(200*self.x_coef), int(600*self.y_coef), int(1600*self.x_coef), int(1000*self.y_coef))))  # capture after click screen
        opened = self.compare_images(img1, img2)  # compares before/after click image to determine if stats are open
        time.sleep(0.1)
        return opened



## open kills

    def openKills(self): #open details of kills
        self.mousePos((int(1223*self.x_coef),int(458*self.y_coef)))
        while not(self.Opening(0.3)) and not self.GUI.close.is_set():
            time.sleep(0.01)

## open dead

    def openDead(self): # open dead data
        self.mousePos((int(620*self.x_coef),int(715*self.y_coef)))
        self.Opening(0.3)
        time.sleep(0.1)

## return to leaderboard

    def returnLeaderBoard(self): #from dead data, return to leaderboard page
        self.mousePos((int(1451*self.x_coef),int(206*self.y_coef)))
        self.leftClick()
        time.sleep(1)
        self.mousePos((int(1431*self.x_coef),int(255*self.y_coef)))
        self.leftClick()
        time.sleep(1)

## Acquire Player stats

    def AcquirePlayer(self): #define the sequence to acquire one's data
        self.openKills()
        time.sleep(0.1)
        image1 = self.screenGrab()
        self.openDead()
        image2 = self.screenGrab()
        self.analyzer.add_player(image1, image2)
        self.returnLeaderBoard()

## Acquire first 4

    def first4(self): # record the 4 first players
        time.sleep(0.1)
        self.mousePos((int(860*self.x_coef),int(404*self.y_coef)))
        if self.Opening(1.5):
            self.AcquirePlayer()
            self.count += 1
            self.GUI.updateBar()
        else:
            self.GUI.updateBar(inc=2)
        self.mousePos((int(860*self.x_coef),int(490*self.y_coef)))
        if self.Opening(1.5):
            self.AcquirePlayer()
            self.count += 1
            self.GUI.updateBar()
        else:
            self.GUI.updateBar(inc=2)
        self.mousePos((int(860*self.x_coef),int(570*self.y_coef)))
        if self.Opening(1.5):
            self.AcquirePlayer()
            self.count += 1
            self.GUI.updateBar()
        else:
            self.GUI.updateBar(inc=2)
        self.mousePos((int(860*self.x_coef),int(650*self.y_coef)))
        if self.Opening(1.5):
            self.AcquirePlayer()
            self.count += 1
            self.GUI.updateBar()
        else:
            self.GUI.updateBar(inc=2)

## Acquire players

    def Acquire(self): # record top x
        time.sleep(1)
        self.first4()# record first players
        stop = False
        for k in range(self.to_capture-4): # record others
            if(k<994):
                self.mousePos((int(860*self.x_coef),int(670*self.y_coef)))
                if keyboard.is_pressed("k") or self.GUI.close.is_set():
                    stop = True
                    break
                if self.Opening(1):
                    self.AcquirePlayer()
                    self.count += 1
                    self.GUI.updateBar()
                else:
                    self.dragClick()
                    self.GUI.updateBar(inc=2)
        if not stop and self.to_capture == 1000 and not self.GUI.close.is_set():
            self.mousePos((int(860*self.x_coef),int(760*self.y_coef)))
            if self.Opening(1.5):
                self.AcquirePlayer()
                self.count += 1
                self.GUI.updateBar()
            else:
                self.GUI.updateBar(inc=2)
            self.mousePos((int(860*self.x_coef),int(840*self.y_coef)))
            if self.Opening(1.5):
                self.AcquirePlayer()
                self.count += 1
                self.GUI.updateBar()
            else:
                self.GUI.updateBar(inc=2)
            
        if stop :
            self.analyzer.stop_worker_thread()
            self.GUI.window.bring_to_front()
            self.GUI.done_pulling()
        else:
            while self.count != self.analyzer.processed:
                time.sleep(2)
            self.analyzer.stop_worker_thread()
            self.GUI.window.bring_to_front()
            self.GUI.done_pulling()

