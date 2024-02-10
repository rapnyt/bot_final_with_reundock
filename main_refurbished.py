import PIL.ImageGrab
import PIL.Image
import colorsys
import sys
import datetime
import pyttsx3
import time
import random
import win32con
import win32api
from playsound import playsound

engine = pyttsx3.init()
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Define icon coordinates
undock_icon = (2698, 193, 2780, 212)
overview_open_icon = (2775, 309, 2791, 322)
overview_type_icon = (2664, 43, 2741, 58)
mining_belt_icon = (2670, 409, 2783, 425)
warp_belt = (2490, 133, 2608, 169)
zoom_out = (2356, 455, 2376, 471)
second_ore_rock = (2660, 138, 2728, 150)
approach_ore_rock = (2480, 191, 2556, 211)
first_strip_miner_module = (2713, 485, 2739, 511)
second_strip_miner_module = (2768, 486, 2790, 508)
overview_station_type_icon = (2670, 213, 2777, 235)
ten_forward_station_icon = (2665, 80, 2771, 99)
dock_icon = (2489, 86, 2596, 111)
dock_icon_2 = (2480, 134, 2593, 164)
cargohold = (1937, 105, 1973, 118)
ore_hold = (1937, 411, 2066, 429)
select_all = (2583, 474, 2623, 509)
move_to = (1969, 119, 2071, 152)
item_hangar = (2200, 133, 2338, 154)
exit_inventory = (2780, 52, 2790, 64)
player_list_field = (2060, 290, 2170, 320)
red_field_check = (2189, 229, 2202, 386)
last_player_icon = (2190, 372, 2200, 382)

def delay(x, y):
    time.sleep(random.uniform(x, y))

def mouse_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    delay(0.05, 0.09)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

def mouse_scroll(x):
    win32api.SetCursorPos((random.randint(x[0], x[2]), random.randint(x[1], x[3])))
    delay(1, 2)
    for i in range(1, 5):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, random.randint(x[0], x[2]), random.randint(x[1], x[3]), -1, 0)
        delay(0.4, 0.5)
    delay(1, 2)

def alarm_bot(x, y):
    v, counter = 0, 0
    start_time = time.time()

    while time.time() - start_time < x:
        icon = PIL.ImageGrab.grab(y, include_layered_windows=False, all_screens=True)
        icon2 = icon
        icon.save('icon.jpg')
        icon = PIL.Image.open('icon.jpg').convert('RGB')
        icon = list(icon.getdata())

        for i in icon:
            h = (colorsys.rgb_to_hsv(i[0], i[1], i[2]))[0]
            if h > 0.92:
                enemy_notification_hostile()
                return 1
            v += (colorsys.rgb_to_hsv(i[0], i[1], i[2]))[2]

        if v / len(icon) < 80:
            counter += 1

        if counter > 40:
            enemy_notification_neutral()
            return 2
        v = 0

def checking_if_reds_still_in_system(x, y):
    start_time = time.time()
    
    while time.time() - start_time < x:
        icon = PIL.ImageGrab.grab(y, include_layered_windows=False, all_screens=True)
        icon.save('icon.jpg')
        icon = PIL.Image.open('icon.jpg').convert('RGB')
        icon = list(icon.getdata())

        for i in icon:
            h = (colorsys.rgb_to_hsv(i[0], i[1], i[2]))[0]
            s = (colorsys.rgb_to_hsv(i[0], i[1], i[2]))[1]
            if h > 0.96 and s > 0.3:
                return 1

def enemy_notification_neutral():
    engine.say("Attention. Neutral player in the system")
    engine.runAndWait()
    playsound("Red Alert.mp3")
    print("Neutral ", datetime.datetime.now())

def enemy_notification_hostile():
    engine.say("Attention. Hostile player in the system")
    engine.runAndWait()
    playsound("Red Alert.mp3")
    print("Hostile ", datetime.datetime.now())

def initial_setup():
    delay(2, 3)
    mouse_scroll(player_list_field)
    delay(0.5, 0.9)

def undocking_sequence():
    print("Undocking at ", datetime.datetime.now())
    mouse_sequence(undock_icon)
    delay(25, 35)
    mouse_scroll(player_list_field)
    delay(0.5, 0.9)

# Define other functions...

if __name__ == "__main__":
    while True:
        while True:
            start_time = time.time()
            initial_setup()
            status = alarm_bot(10, last_player_icon)
            if status == 1:
                break
            elif status == 2:
                sys.exit()
            undocking_sequence()
            status = alarm_bot(5, last_player_icon)
            if status:
                flee_sequence_1()
                if status == 1:
                    break
                elif status == 2:
                    sys.exit()
            # Perform other sequences...

            time_check_if_finished()
        delay(120, 180)
        # Additional logic...
