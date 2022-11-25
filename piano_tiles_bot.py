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
    
# X: 1624 Y:  607 RGB: (  0,   0,   0)
# X: 1510 Y:  609 RGB: (255, 255, 255)
# X: 1406 Y:  594 RGB: ( 22,  22,  22)
# X: 1300 Y:  575 RGB: (177, 181, 234)

if __name__ == "__main__":
    y = 500

    while keyboard.is_pressed('q') == False:
        if pyautogui.pixel(1624,y)[0] == 0:
            click(1624,y+15)
        if pyautogui.pixel(1510,y)[0] == 0:
            click(1510,y+15)
        if pyautogui.pixel(1406,y)[0] == 0:
            click(1406,y+15)
        if pyautogui.pixel(1300,y)[0] == 0:
            click(1300,y+15)    
        
