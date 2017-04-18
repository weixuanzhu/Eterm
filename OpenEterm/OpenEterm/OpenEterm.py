#-*- coding: UTF-8 -*-

import os
import sys
from time import sleep
import win32gui
import win32api
import win32con
import ConfigParser
import win32clipboard as w
#参考页面 http://blog.csdn.net/fallinlovelj/article/details/54343520;http://blog.csdn.net/suzyu12345/article/details/52934328

class OpenEterm():
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'
    ROBOT_LIBRARY_VERSION = '0.1'

    def __init__(self):
        self.eterm = None

    def OpenEterm(self):
        #打开Eterm
        os.startfile("C:\Program Files (x86)\Travelsky\eTerm3\eTerm3.exe")
        sleep(1)
        # 获取窗口句柄
        windowName = "eTerm"
        self.eterm = win32gui.FindWindow(None,windowName)
        print self.eterm

        sleep(1)
        #模拟按下Enter键，（键表在1指令.xlsx里）
        win32api.keybd_event(13,0,win32con.WM_KEYUP,0)

    def esc(self):
        sleep(1)
        win32api.keybd_event(27,0,win32con.WM_KEYUP,0)

    def F12(self):
       sleep(1)
       win32api.keybd_event(123,0,win32con.WM_KEYUP,0)

    def CtrlA(self):
        sleep(0.5)
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        sleep(0.5)
        win32api.keybd_event(65, 0, 0, 0)
        sleep(0.5)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP , 0)

    def F1(self):
        sleep(1)
        win32api.keybd_event(112, 0,win32con.WM_KEYUP,0)

    def F2(self):
        sleep(1)
        win32api.keybd_event(113 ,0,win32con.WM_KEYUP,0)

    def F3(self):
        sleep(1)
        win32api.keybd_event(114,0,win32con.WM_KEYUP,0)

    def F4(self):
        sleep(1)
        win32api.keybd_event(115,0,win32con.WM_KEYUP,0)