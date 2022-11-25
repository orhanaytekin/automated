'''
Not working properly,
'''

# RGB: (211, 211, 211) rgb value of the ball and the obstacles
# X: 1068 Y:  833 RGB: (  0,   0,   0) max y value of the map
# X: 1068 Y:  276 RGB: (  0,   0,   0) min y value of the map

from pyautogui import *
import pyautogui
import keyboard
#import random
import win32api , win32con
import time

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(0.015)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)


def check_squares():
    while keyboard.is_pressed('q') == False:
        for i in range(276,833,50):
            print("Checking  : 833 , y : " + str(i) )
            if pyautogui.pixel(1068,i)[0] == 211:
                click(1100,i)
                print("Clicked")
                
    

if __name__ == '__main__':
    print("sleep for")
    time.sleep(2)
    check_squares()