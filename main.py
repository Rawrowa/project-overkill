# Windows only
# The whole thing is pretty slow, it takes around 0.5s to save 1080p screenshot, and additional 1s to get a reply from the LLM
# If multiple screenshots were taken in a span of a second and the process is fresh (PID updates every launch), the LLM will be prompted every time
# The screen should be captured right away, before all the slow silly stuff

from win32process import GetWindowThreadProcessId
from pygetwindow import getActiveWindow
from psutil import Process
from win32api import GetFileVersionInfo

import cohere
import creds

import os
import datetime as dt
import pyautogui as ag
import keyboard as kb

from threading import Thread
import time


EXE_PROPS = ["InternalName", "ProductName", "CompanyName", "FileDescription"]
BAD_CHARS = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '™']

memory = {
    "pid": int(),
    "app": str(),
}

# TODO: GUI using Tkinter
BIND = "F10"
DIRECTORY = r"C:\TEST"
FORMAT = ".png"


def identify_app(pid):
    """Collects data and identifies the application"""
    # win_title = getActiveWindow().title
    exe_path = Process(pid).exe()

    # Getting the exe metadata, the number of fields is limited it in the global variable
    try:
        exe_data = set()
        language, codepage = GetFileVersionInfo(exe_path, "\\VarFileInfo\\Translation")[0]
        for prop in EXE_PROPS:
            string_file_info = u"\\StringFileInfo\\%04X%04X\\%s" % (language, codepage, prop)
            data = str(GetFileVersionInfo(exe_path, string_file_info))
            if data != "None" and data != "" and data != " ":
                exe_data.add(data)
    except:
        exe_data = "None"

    # Asking AI-senpai for help UwU
    co = cohere.Client(creds.api_key)
    response = co.generate(
        model="command",
        temperature=1,
        prompt="Identify the app and print ONLY its title (or 'Other' if it can't be identified)"
                f"\nexe path: {exe_path}"
                f"\nexe data: {exe_data}")
    # print(f"exe path: {exe_path}\nexe data: {exe_data}")

    # Formating the name since Windows doesn't allow for certain characters in paths
    app_name = "".join(char for char in response.generations[0].text.strip() if char not in BAD_CHARS)
    return app_name


def capture_screenshot():
    t_start = time.perf_counter()

    # Main event - capturing the whole screen
    screenshot = ag.screenshot()
    scr_time = dt.datetime.now().strftime('%y-%m-%d %H-%M-%S-%f')[:-4]

    # TODO: Better "memory system"
    # Checking if the running app was already identified to save on tokens
    pid_active = GetWindowThreadProcessId(getActiveWindow()._hWnd)[1]
    if pid_active == memory["pid"]:
        app_active = memory["app"]

        print(f"{app_active} was already indentified")
    else:
        app_active = identify_app(pid_active)
        memory["pid"] = pid_active
        memory["app"] = app_active

        print(f"Identified app: {app_active}")

    # Creating the folder if it doesn't already exists
    dir_active = os.path.join(DIRECTORY, app_active)
    if not os.path.exists(dir_active):
        os.mkdir(dir_active)
    
    # Saving already taken screenshot to a file
    scr_name = f"{app_active} {scr_time}{FORMAT}"
    screenshot.save(os.path.join(dir_active, scr_name))

    t_end = time.perf_counter()
    print(f"Screenshot was saved to {dir_active}\nElapsed time: {t_end - t_start:0.3f} s\n")


def new_screenshot_thread():
    """Calls capture_screenshot in its own thread"""
    # Every screenshot is taken separately, so spaming the key should be ok
    thrd1 = Thread(target=capture_screenshot)
    thrd1.start()


kb.add_hotkey(BIND, new_screenshot_thread)
kb.wait()
