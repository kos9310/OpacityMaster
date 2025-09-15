import sys
import os

from tkinter import *
import threading
import keyboard

import win32gui
import win32con
import win32process
import psutil

import pystray
from PIL import Image

if getattr(sys, 'frozen', False):
    program_directory = os.path.dirname(os.path.abspath(sys.executable))
else:
    program_directory = os.path.dirname(os.path.abspath(__file__))
os.chdir(program_directory)

# 투명도 초기값
scaleVal = 50

hotkey_thread = None


# 리소스 경로 설정
def resource_path(relative_path):
    try:
        # PyInstaller에서 실행되는 경우
        base_path = sys._MEIPASS
    except Exception:
        # 개발 중일 경우
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
icon_path = resource_path("OpacityMaster.ico")
image_path = resource_path("image.png")

def get_chrome_pip_hwnd():
    global hwnd
    pip_hwnd = None

    def enum_window_callback(hwnd, _):
        nonlocal pip_hwnd
        if not win32gui.IsWindowVisible(hwnd):
            return
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            proc = psutil.Process(pid)
            if "chrome" in proc.name().lower():
                class_name = win32gui.GetClassName(hwnd)
                title = win32gui.GetWindowText(hwnd)
                # 크롬의 기본 PIP 기능을 사용하지 않는 사람은 "PIP 모드" 대신 다른 텍스트 입력이 필요
                if class_name == "Chrome_WidgetWin_1" and "PIP 모드" in title:
                    pip_hwnd = hwnd
        except Exception:
            pass

    win32gui.EnumWindows(enum_window_callback, None)
    hwnd = pip_hwnd

def set_window_transparency(hwnd, alpha=180):
    if hwnd is None:
        print("PIP 창을 찾지 못했습니다.")
        return

    print("alpha", alpha);
    styles = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, styles | win32con.WS_EX_LAYERED)
    win32gui.SetLayeredWindowAttributes(hwnd, 0, alpha, win32con.LWA_ALPHA)
    print("투명도: ", int(alpha*100/255))
def close_target_window(hwnd):
    if hwnd and win32gui.IsWindow(hwnd):
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            print("창 종료 요청 전송")
        except Exception as e:
            print(f"PostMessage 실패: {e}")
    else:
        print("유효하지 않은 창 핸들입니다. 종료 요청 생략됨.")

def on_scale_change(value):
    alpha = int(round(float(value) / 100 * 255))
    set_window_transparency(hwnd, alpha)

def watch_hotkey():
    # ctrl+alt+shift 누르면 종료
    keyboard.add_hotkey('ctrl+alt+shift', lambda: close_target_window(hwnd))
    # z + x + c + v 누르면 창열기
    keyboard.add_hotkey('z+x+c+v', lambda: toggle_window())

def start_hotkey_listener():
    global hotkey_thread
    print(start_hotkey_listener, hotkey_thread)
    if hotkey_thread is None or not hotkey_thread.is_alive():
        hotkey_thread = threading.Thread(target=watch_hotkey, daemon=True)
        hotkey_thread.start()

def quit_window(icon, item):
    """ 프로그램 종료 """
    icon.visible = False
    icon.stop()
    root.quit()
def show_window():
    """ 프로그램 창 열기 """
    root.after(0, lambda: (
        root.deiconify(),  # 창 다시 보이게
        root.lift(),  # 다른 창들 위로 올리기
        root.focus_force()  # 창에 포커스 주기
    ))
def withdraw_window():
    """ 프로그램 창 숨기기 """
    root.withdraw()
def toggle_window():
    """ 프로그램 창 상태를 토글 """
    if root.state() == 'normal':  # 창이 열려 있으면
        withdraw_window()  # 창 숨기기
    else:
        show_window()  # 창 보이기

# 실행
get_chrome_pip_hwnd()
set_window_transparency(hwnd, alpha=int(round(float(scaleVal)/100*255)))

root = Tk()
root.title("OpacityMaster")
root.iconbitmap(icon_path)

scale = Scale(root, from_=0, to=100, orient="horizontal", length=250, command=on_scale_change)
scale.set(scaleVal)
scale.pack(pady=10)
btn = Button(root, text="동기화", command=get_chrome_pip_hwnd)
btn.pack(anchor="center", pady=10)

start_hotkey_listener()

root.withdraw()

image = Image.open(image_path)
icon = pystray.Icon("name", image, "OpacityMaster", pystray.Menu(
    pystray.MenuItem("Show", show_window, default=True),
    pystray.MenuItem("Quit", quit_window)))
icon.run_detached()

root.protocol('WM_DELETE_WINDOW', withdraw_window)
root.mainloop()