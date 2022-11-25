from piano_tiles_bot import click
from pyautogui import *
import keyboard
import time
import string
import random
print("seleeping for you")
time.sleep(5)
print("sleep time is over babyyy")

# X: 1118 Y:  966 RGB: (255, 255, 255)

#a = 1


def send_message(x):
    for i in range(x):
        if i == 0:
            S = int(input("message length: "))
            print("PRESS SPACE TO INITIATE babyyy")
            keyboard.wait('space')
            time.sleep(1)

        while keyboard.is_pressed('q') == False:

            ran = ''.join(random.choices(
                string.ascii_uppercase + string.digits, k=S))
            keyboard.write(ran, 0.0000001)
            keyboard.send('enter')
            time.sleep(0.00000001)


if __name__ == "__main__":
    send_message(500)

    # while keyboard.is_pressed('q') == False:
    #     if a == 1:
    #         print("PRESS SPACE TO INITIATE babyyy")
    #         keyboard.wait('space')
    #         a += 1
    #     keyboard.write('Automated Message by An An')
    #     keyboard.send('enter')
    #     time.sleep(1)
