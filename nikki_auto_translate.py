# -*- coding: utf-8 -*-

from PyQt6 import QtGui, QtCore 
from PyQt6.QtWidgets import *

import time, pyuac, sys
import pyperclip
import win32gui, win32process, win32con, win32api
import psutil
import ctypes
from googletrans import Translator
from ctypes import wintypes
from PIL import Image, ImageGrab
import pyautogui

ctypes.windll.shcore.SetProcessDpiAwareness(1)

class Trans(QtCore.QThread):
    log = QtCore.pyqtSignal(str)

    def __init__(self, word, orginal, trans, parent=None):
        super(Trans, self).__init__(parent)
        self.word = word
        self.orginal = orginal
        self.translate = trans

    def run(self):
        translator = Translator()
        translation = translator.translate(self.word, dest=self.translate)
        self.log.emit("Translate: "+translation.text)
        output = translation.text
        translator2 = Translator()
        translation2 = translator.translate(output, dest=self.orginal)
        self.log.emit("Recheck: "+translation2.text)

class Sender(QtCore.QThread):
    log = QtCore.pyqtSignal(str)

    def __init__(self, word, parent=None):
        super(Sender, self).__init__(parent)
        self.xpid = 0
        self.hwnds = []
        self.word = word

    def callback(self, hwnd, hwnds):
        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
        if found_pid == self.xpid and win32gui.IsWindowVisible(hwnd):
            self.hwnds.append(hwnd)
        return True

    def run(self):

        for pid in psutil.pids():
            process = psutil.Process(pid)
            if process.name() == 'X6Game-Win64-Shipping.exe':
                self.xpid = pid
       
        win32gui.EnumWindows(self.callback, self.hwnds)
        #global hwnd
        hwnd = self.hwnds[0]
        win32gui.SetForegroundWindow(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        #print (rect)
        x = rect[0]
        y = rect[1]
        w = rect[2] - x  #1102
        h = rect[3] - y  #1976
        #box1 = (x, y, x+w, y+h)
        #im = ImageGrab.grab(box1)
        #im.save('temp_cd.png')
        #x+725 y+1400
        pyautogui.moveTo(x+725, y+1400)
        time.sleep(0.5)
        pyautogui.click()
        pyautogui.keyDown('ctrl')
        time.sleep(0.01)
        pyautogui.keyDown('v')
        time.sleep(0.01)
        pyautogui.keyUp('v')
        time.sleep(0.01)
        pyautogui.keyUp('ctrl')
        
        """
        pyautogui.write(self.word)
        time.sleep(0.5)
        """
        pyautogui.press('enter')
        #self.log.emit("Range: 36 - 71")
        

#walker = Walker()
#walker.start()

class Form(QWidget):

    def __init__(self):
        super(Form, self).__init__()
        #self.resize(300,200)
        self.setFixedSize(700,300)
        self.setWindowTitle(u"Translator: Nikki")
        self.setWindowFlags(QtCore.Qt.WindowType.Window   |
                                   QtCore.Qt.WindowType.CustomizeWindowHint |
                                   QtCore.Qt.WindowType.WindowTitleHint |                                   
                                   QtCore.Qt.WindowType.WindowCloseButtonHint)
        ################

        self.output_text = ""
        self.long = ['English', 'Japanese','Chinese', 'Thai', 'Vietnamese', 'Indonesian', 'Romanian', 'Malayalam',
                     'Korean']
        self.short = ['en', 'ja', 'zh-Tw', 'th', 'vi', 'id', 'ro', 'ml', 'ko']

        ################
        
        vbox = QVBoxLayout()
        self.text_area = QTextBrowser()
        self.text_area.setFixedSize(330,240)
        habox = QHBoxLayout()
        habox.addWidget(self.text_area)

        self.e1 = QTextEdit()
        self.e1.setFixedSize(350,240)

        #hbbox = QHBoxLayout()
        #hbbox.addWidget(self.e1)
        habox.addWidget(self.e1)

        hcbox = QHBoxLayout()
        self.text = QLabel(u'Orginal:')
        self.text.setFixedSize(60,35)
        self.cb = QComboBox()
        self.cb.addItems(self.long)
        self.cb.currentIndexChanged.connect(self.selectionchange)  
        self.text2 = QLabel(u'Translate:')
        self.text2.setFixedSize(100,35)
        self.cb2 = QComboBox()
        self.cb2.addItems(self.long)
        self.cb2.currentIndexChanged.connect(self.selectionchange2)
        
        self.translate_btn = QPushButton(u'Translate')
        self.translate_btn.setFixedSize(110,35)
        self.translate_btn.clicked.connect(self.translate)
        self.send_btn = QPushButton(u'Send Result')
        self.send_btn.setFixedSize(150,35)
        self.send_btn.clicked.connect(self.send)

        hcbox.addWidget(self.text)
        hcbox.addWidget(self.cb)
        hcbox.addWidget(self.text2)
        hcbox.addWidget(self.cb2)
        hcbox.addWidget(self.translate_btn)
        hcbox.addWidget(self.send_btn)

        vbox.addLayout(habox)
        #vbox.addLayout(hbbox)
        vbox.addLayout(hcbox)
        self.setLayout(vbox)

        self.input = 'en'


    def log_status(self, log_list):
        self.text_area.append(log_list)
        if log_list.count("Translate: ") >0:
            #print (log_list.replace("Recheck: ",""))
            #self.output_text = log_list.replace("Translate: ","")
            pyperclip.copy(log_list.replace("Translate: ",""))

    def translate(self):
        self.walker_trans = Trans(self.e1.toPlainText(), self.input, self.output)
        self.walker_trans.log.connect(self.log_status)
        self.walker_trans.start()        
        
    def send(self):
        #self.manga_name = self.e1.toPlainText()
        #print (self.output_text)
        self.walker_send = Sender(self.output_text)
        self.walker_send.start()

    def selectionchange(self,i):
        self.input = self.short[i]
        #print (self.input)

    def selectionchange2(self,i):
        self.output = self.short[i]
        print (self.output)
        
def my_excepthook(type, value, tback):
      sys.__excepthook__(type, value, tback)

if __name__ == '__main__':
    if not pyuac.isUserAdmin():        
        pyuac.runAsAdmin()
    else:
        sys.excepthook = my_excepthook
        app = QApplication(sys.argv)
        widget1= Form()
        widget1.show()
        sys.exit(app.exec())
