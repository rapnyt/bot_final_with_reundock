# grid layout 2, default, width 940 height 528, remember
# import ctypes

import PIL.ImageGrab
import PIL.Image
import PIL.ImageChops
import pytesseract
import random
import win32con
import win32api
import time
import datetime
import pyttsx3
import colorsys
import sys
from playsound import playsound
engine = pyttsx3.init()
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

undock_icon = (2698, 193, 2780, 212)
overview_open_icon = (2775, 309, 2791, 322)
overview_type_icon = (2664, 43, 2741, 58)
overview_mining_type_icon = (2670, 409, 2783, 425)
mining_belt_icon = (2660, 78, 2777, 99)
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


def alarm_bot(x, y):

    # import PIL.ImageGrab
    # import pyttsx3
    # import pytesseract
    # import time
    # import datetime
    # import PIL.Image
    # import PIL.ImageChops
    # import colorsys
    # from playsound import playsound
    #
    # engine = pyttsx3.init()
    # pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    # last_player_icon = (2186, 372, 2194, 378)
    # last_player_icon = (2190, 372, 2200, 382)
    # wait_interval = 1
    v = 0
    start_time = time.time()
    counter = 0

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
            # print("counter raised by v")
        if counter > 40:
            enemy_notification_neutral()
            return 2
        v = 0
        # print("Everything's alright ", datetime.datetime.now())
        # time.sleep(wait_interval)
        # print(time.time() - start_time)


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
                print((colorsys.rgb_to_hsv(i[0], i[1], i[2]))[0], (colorsys.rgb_to_hsv(i[0], i[1], i[2]))[1], (colorsys.rgb_to_hsv(i[0], i[1], i[2]))[2])
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


def mouse_click(x, y):
    win32api.SetCursorPos((x, y))
    # ctypes.windll.user32.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(random.uniform(0.05, 0.09))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def mouse_scroll(x):
    win32api.SetCursorPos((random.randint(x[0], x[2]), random.randint(x[1], x[3])))
    # ctypes.windll.user32.SetCursorPos((random.randint(x[0], x[2]), random.randint(x[1], x[3])))
    delay(1, 2)
    for i in range(1, 5):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, random.randint(x[0], x[2]), random.randint(x[1], x[3]), -1, 0)
        time.sleep(random.uniform(0.4, 0.5))
    delay(1, 2)


def mouse_scroll_for_reds(x, y):
    win32api.SetCursorPos((random.randint(x[0], x[2]), random.randint(x[1], x[3])))
    # ctypes.windll.user32.SetCursorPos((random.randint(x[0], x[2]), random.randint(x[1], x[3])))
    delay(1, 2)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, random.randint(x[0], x[2]), random.randint(x[1], x[3]), y, 0)
    # time.sleep(random.uniform(0.4, 0.5))
    # delay(1, 2)


def mouse_sequence(x):
    mouse_click(random.randint(x[0], x[2]), random.randint(x[1], x[3]))
    time.sleep(random.uniform(0.8, 2.2))
    mouse_click(random.randint(x[0] + 940, x[2] + 940), random.randint(x[1], x[3]))
    time.sleep(random.uniform(0.8, 2.2))
    mouse_click(random.randint(x[0], x[2]), random.randint(x[1] + 528, x[3] + 528))
    time.sleep(random.uniform(0.8, 2.2))
    mouse_click(random.randint(x[0] + 940, x[2] + 940), random.randint(x[1] + 528, x[3] + 528))
    time.sleep(random.uniform(0.8, 2.2))


def mouse_mining_sequence(x, y):
    mouse_click(random.randint(x[0], x[2]), random.randint(x[1], x[3]))
    time.sleep(random.uniform(0.4, 0.8))
    mouse_click(random.randint(y[0], y[2]), random.randint(y[1], y[3]))
    time.sleep(random.uniform(0.8, 2.2))
    mouse_click(random.randint(x[0] + 940, x[2] + 940), random.randint(x[1], x[3]))
    time.sleep(random.uniform(0.4, 0.8))
    mouse_click(random.randint(y[0] + 940, y[2] + 940), random.randint(y[1], y[3]))
    time.sleep(random.uniform(0.8, 2.2))
    mouse_click(random.randint(x[0], x[2]), random.randint(x[1] + 528, x[3] + 528))
    time.sleep(random.uniform(0.4, 0.8))
    mouse_click(random.randint(y[0], y[2]), random.randint(y[1] + 528, y[3] + 528))
    time.sleep(random.uniform(0.8, 2.2))
    mouse_click(random.randint(x[0] + 940, x[2] + 940), random.randint(x[1] + 528, x[3] + 528))
    time.sleep(random.uniform(0.4, 0.8))
    mouse_click(random.randint(y[0] + 940, y[2] + 940), random.randint(y[1] + 528, y[3] + 528))
    time.sleep(random.uniform(0.8, 2.2))


def delay(x, y):
    time.sleep(random.uniform(x, y))


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


def overview_sequence():
    mouse_sequence(overview_open_icon)
    delay(0.5, 0.9)
    mouse_sequence(overview_type_icon)
    delay(0.5, 0.9)
    mouse_sequence(overview_mining_type_icon)
    delay(0.5, 0.9)
    mouse_sequence(mining_belt_icon)


def warping():
    print("Warping at ", datetime.datetime.now())
    mouse_sequence(warp_belt)
    delay(0.5, 0.9)
    mouse_sequence(zoom_out)
    delay(55, 60)


def initial_mining_sequence():
    print("Mining at ", datetime.datetime.now())
    mouse_mining_sequence(first_strip_miner_module, second_strip_miner_module)


def approach_and_flee_sequence():
    mouse_sequence(second_ore_rock)
    delay(0.5, 0.9)
    mouse_sequence(approach_ore_rock)
    delay(0.5, 0.9)
    mouse_sequence(overview_type_icon)
    delay(0.5, 0.9)
    mouse_sequence(overview_station_type_icon)
    delay(0.5, 0.9)
    mouse_sequence(ten_forward_station_icon)
    print("Commencing alarm bot at ", datetime.datetime.now())


def docking_sequence():
    print("Docking at ", datetime.datetime.now())
    mouse_sequence(dock_icon)
    delay(70, 90)


def unload_sequence():
    mouse_sequence(cargohold)
    delay(10, 20)
    mouse_sequence(ore_hold)
    delay(2, 5)
    mouse_sequence(select_all)
    delay(2, 5)
    mouse_sequence(move_to)
    delay(2, 5)
    mouse_sequence(item_hangar)
    delay(10, 20)
    mouse_sequence(exit_inventory)
    delay(4, 9)
    print("Cycle completed at ", datetime.datetime.now())


def time_check_if_finished():
    now = datetime.datetime.now()
    endgame = now.replace(day=16, hour=18, minute=0, second=0, microsecond=0)
    if now > endgame:
        print("Mining bot completed at ", datetime.datetime.now())
        sys.exit()


def flee_sequence_1():
    mouse_sequence(overview_open_icon)
    delay(0.5, 0.9)
    mouse_sequence(ten_forward_station_icon)
    delay(0.5, 0.9)
    mouse_sequence(dock_icon_2)


def flee_sequence_2():
    mouse_sequence(overview_type_icon)
    delay(0.5, 0.9)
    mouse_sequence(overview_station_type_icon)
    delay(1, 3)
    mouse_sequence(ten_forward_station_icon)
    delay(0.5, 0.9)
    mouse_sequence(dock_icon_2)


def flee_sequence_3():
    delay(1, 3)
    mouse_sequence(overview_type_icon)
    delay(0.5, 0.9)
    mouse_sequence(overview_station_type_icon)
    delay(0.5, 0.9)
    mouse_sequence(ten_forward_station_icon)
    delay(0.5, 0.9)
    mouse_sequence(dock_icon)


def flee_sequence_4():
    delay(1, 5)
    mouse_sequence(dock_icon)


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
        overview_sequence()
        status = alarm_bot(5, last_player_icon)
        if status:
            flee_sequence_2()
            if status == 1:
                break
            elif status == 2:
                sys.exit()
        warping()
        status = alarm_bot(5, last_player_icon)
        if status:
            flee_sequence_3()
            if status == 1:
                break
            elif status == 2:
                sys.exit()
        initial_mining_sequence()
        status = alarm_bot(5, last_player_icon)
        if status:
            flee_sequence_3()
            if status == 1:
                break
            elif status == 2:
                sys.exit()
        approach_and_flee_sequence()
        status = alarm_bot(random.uniform(740, 780), last_player_icon)
        if status:
            flee_sequence_4()
            if status == 1:
                break
            elif status == 2:
                sys.exit()
        docking_sequence()
        unload_sequence()
        time_check_if_finished()
    delay(120, 180)
    if time.time() - start_time > 600:
        delay(60, 120)
        unload_sequence()
        print("Unloading procedure after emergency docking sequence done at ", datetime.datetime.now())
    while status == 1:
        j = 0
        mouse_scroll_for_reds(player_list_field, +100)
        delay(2, 5)
        for i in range(1, 3):
            mouse_scroll_for_reds(player_list_field, -100)
            if checking_if_reds_still_in_system(5, red_field_check) == 1:
                j += 1
        time_check_if_finished()
        if j == 0:
            print("No hostile detected at ", datetime.datetime.now())
            break
        print("Hostiles still present at ", datetime.datetime.now())
        delay(120, 300)


# last_player_icon = (2190, 372, 2200, 382)
# icon = PIL.ImageGrab.grab(last_player_icon, include_layered_windows=False, all_screens=True)
# icon.show()
