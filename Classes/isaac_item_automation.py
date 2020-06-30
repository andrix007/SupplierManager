from mss import mss
import matplotlib.pyplot as plt
import matplotlib.image as mpimg

with mss() as sct:

    file = sct.shot()
    img = mpimg.imread(file)
    imgplot = plt.imshow(img)
    plt.show()

"""
import pyautogui as pag
import time

INTERVAL = 1
screenshot_path = 'C:\\Users\\Andrei Bancila\\Desktop\\Isaac\\scr.png'


def restart_run():

    pag.keyDown("r")
    time.sleep(2)
    pag.keyUp("r")


def save_screenshot():

    scr = pag.screenshot()
    scr.save(screenshot_path)

time.sleep(5)
print("The execution has begun...")

for _ in range(3):

    save_screenshot()

    restart_run()
    time.sleep(INTERVAL)

print("The execution has ended...")
"""
