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

# 설정 파일 경로 및 기본 핫키
CONFIG_FILE = os.path.join(program_directory, "hotkey.txt")
DEFAULT_TOGGLE_HOTKEY = 'z+x+c+v'
toggle_hotkey = None
close_hotkey = 'ctrl+alt+shift'
toggle_hotkey_handle = None
close_hotkey_handle = None


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

def load_hotkey():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            val = f.read().strip()
            if val:
                return val
    return DEFAULT_TOGGLE_HOTKEY

def save_hotkey(hotkey):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        f.write(hotkey)

def bind_hotkeys():
    global close_hotkey_handle, toggle_hotkey_handle
    if close_hotkey_handle is not None:
        keyboard.remove_hotkey(close_hotkey_handle)
    if toggle_hotkey_handle is not None:
        keyboard.remove_hotkey(toggle_hotkey_handle)
    close_hotkey_handle = keyboard.add_hotkey(close_hotkey, lambda: close_target_window(hwnd))
    toggle_hotkey_handle = keyboard.add_hotkey(toggle_hotkey, lambda: toggle_window())

def open_hotkey_dialog():
    dialog = Toplevel(root)
    dialog.title("핫키 설정")
    Label(dialog, text="창 토글 핫키를 입력하세요").pack(padx=20, pady=(20, 10))

    hotkey_var = StringVar(value=toggle_hotkey)
    Entry(dialog, textvariable=hotkey_var, state="readonly", justify="center").pack(padx=20, pady=(0, 10))

    def capture():
        global toggle_hotkey_handle, close_hotkey_handle
        if toggle_hotkey_handle is not None:
            keyboard.remove_hotkey(toggle_hotkey_handle)
            toggle_hotkey_handle = None
        if close_hotkey_handle is not None:
            keyboard.remove_hotkey(close_hotkey_handle)
            close_hotkey_handle = None
        hotkey_var.set(keyboard.read_hotkey(suppress=False))

    def save():
        global toggle_hotkey
        new_hotkey = hotkey_var.get()
        if new_hotkey:
            toggle_hotkey = new_hotkey
            save_hotkey(new_hotkey)
        bind_hotkeys()
        dialog.destroy()

    def cancel():
        bind_hotkeys()
        dialog.destroy()

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    threading.Thread(target=capture, daemon=True).start()

    btn_frame = Frame(dialog)
    btn_frame.pack(pady=(0, 20))
    Button(btn_frame, text="저장", command=save).pack(side=LEFT, padx=10)
    Button(btn_frame, text="취소", command=cancel).pack(side=LEFT, padx=10)

toggle_hotkey = load_hotkey()

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

btn_hotkey = Button(root, text="핫키 설정", command=open_hotkey_dialog)
btn_hotkey.pack(anchor="center", pady=5)

bind_hotkeys()

root.withdraw()

image = Image.open(image_path)
icon = pystray.Icon("name", image, "OpacityMaster", pystray.Menu(
    pystray.MenuItem("Show", show_window, default=True),
    pystray.MenuItem("Quit", quit_window)))
icon.run_detached()

root.protocol('WM_DELETE_WINDOW', withdraw_window)
root.mainloop()
