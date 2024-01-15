import sys
import json
import os
import pygetwindow as gw
import time
import threading
import win32con
import win32gui
import win32api
from colormap import color_mapping
from pynput import keyboard
from rich.console import Console
from rich.style import Style
import random

# Globals
overlay_window_handle = None
overlay_thread_running = False
GLOBAL_DELAY = None
GLOBAGL_DEBUG = False
keyboard_listener = None
shift_pressed = False
config_file_path = 'colorconfig.json'
color_keys = list(color_mapping.keys())

# Config
def load_config():
    global GLOBAL_DELAY, GLOBAGL_DEBUG
    try:
        with open('colorconfig.json', 'r') as config_file:
            config = json.load(config_file)
            GLOBAL_DELAY = config.get('delay', 1)
            GLOBAGL_DEBUG = config.get('debug', False)
            return config
    except Exception as e:
        print(f"Error loading config file: {e}")
        sys.exit(1)
config = load_config()
def save_config(config):
    with open(config_file_path, 'w') as file:
        json.dump(config, file, indent=4)

# Keyboard
def start_keyboard_listener():
    with keyboard.Listener(
        on_release=on_key_release,
        on_press=on_key_press) as listener:
        listener.join()
def on_key_release(key):
    global shift_pressed, overlay_window_handle
    if overlay_window_handle:
        try:
            if key == keyboard.Key.left:
                if shift_pressed:
                    cycle_color('previous')
                else:
                    adjust_transparency(-0.10)
                    if GLOBAGL_DEBUG:
                        print("Left arrow key pressed, decreasing transparency")
            elif key == keyboard.Key.right:
                if shift_pressed:
                    cycle_color('next')
                else:
                    adjust_transparency(0.10)
                    if GLOBAGL_DEBUG:
                        print("Right arrow key pressed, increasing transparency")
            elif key == keyboard.Key.shift:
                shift_pressed = False
        except AttributeError:
            pass
def on_key_press(key):
    global shift_pressed, overlay_window_handle
    if overlay_window_handle:
        try:
            if key == keyboard.Key.shift:
                shift_pressed = True
        except AttributeError:
            pass
def adjust_transparency(adjustment):
    global config
    config = load_config()
    new_transparency = round(max(0, min(1, config['transparency'] + adjustment)), 2)
    config['transparency'] = new_transparency
    save_config(config)
    recreate_overlay_window()
def cycle_color(direction):
    global config
    config = load_config()
    current_color = config.get('overlay_color', 'black')
    current_index = color_keys.index(current_color)
    if direction == 'next':
        new_index = (current_index + 1) % len(color_keys)
    else:
        new_index = (current_index - 1) % len(color_keys)
    config['overlay_color'] = color_keys[new_index]
    save_config(config)
    recreate_overlay_window()

# Windows Class and proc
def window_proc(hwnd, msg, wParam, lParam):
    config = load_config()
    overlay_color_name = config.get('overlay_color', 'red')
    transparency = config.get('transparency', 0.5)
    overlay_color_rgb = color_mapping.get(overlay_color_name.lower(), (255, 0, 0))
    if msg == win32con.WM_PAINT:
        hdc, paintStruct = win32gui.BeginPaint(hwnd)
        rect = win32gui.GetClientRect(hwnd)
        win32gui.FillRect(hdc, rect, win32gui.CreateSolidBrush(win32api.RGB(*overlay_color_rgb)))
        win32gui.EndPaint(hwnd, paintStruct)
        return 0
    elif msg == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    else:
        return win32gui.DefWindowProc(hwnd, msg, wParam, lParam)
wc = win32gui.WNDCLASS()
wc.hInstance = win32gui.GetModuleHandle(None)
wc.lpszClassName = 'TransparentWindowClass'
wc.lpfnWndProc = window_proc
class_atom = win32gui.RegisterClass(wc)
if not class_atom:
    raise Exception("Failed to register window class")

# Overlay
def create_click_through_window(x, y, width, height, overlay_color_name, transparency):
    global overlay_window_handle, keyboard_listener
    config = load_config()
    overlay_color_name = config.get('overlay_color', 'red')
    transparency = config.get('transparency', 0.5)
    if GLOBAGL_DEBUG:   
        print(f"Creating click-through window at {x}, {y}, size {width}x{height}")
    overlay_color_rgb = color_mapping.get(overlay_color_name.lower(), (0, 0, 0))
    hwnd = win32gui.CreateWindowEx(
        win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED,
        'TransparentWindowClass',
        '',
        0,  # Window style
        x, y, width, height,
        0, 0, wc.hInstance, None
    )
    if hwnd:
        if GLOBAGL_DEBUG:
            print("Window handle created:", hwnd)
        alpha = int(transparency * 255)
        colorref = win32api.RGB(*overlay_color_rgb) # I could not figure this out only with windows proc.
        win32gui.SetLayeredWindowAttributes(hwnd, colorref, alpha, win32con.LWA_ALPHA)
        keyboard_listener = keyboard.Listener(on_release=on_key_release, on_press=on_key_press)
        keyboard_listener.start()
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        overlay_window_handle = hwnd
        bring_to_front(hwnd)
        while True:
            ret, msg = win32gui.GetMessage(None, 0, 0)
            if ret:
                win32gui.TranslateMessage(msg)
                win32gui.DispatchMessage(msg)
            else:
                break
    else:
        print("Failed to create window handle")
def keep_overlay_on_top():
    global overlay_window_handle, overlay_thread_running
    overlay_thread_running = True
    while overlay_thread_running:
        if overlay_window_handle:
            bring_to_front(overlay_window_handle)
            if GLOBAGL_DEBUG:
                log_window_info(overlay_window_handle, "Overlay Window")
        time.sleep(GLOBAL_DELAY)
    if GLOBAGL_DEBUG:    
        print("Overlay keeping thread stopped.")
def recreate_overlay_window(): # Hacky way to redraw overlay if config changes just kill it...
    global overlay_window_handle
    if overlay_window_handle:
        close_overlay_window() # It will reload.
def close_overlay_window():
    global overlay_window_handle, overlay_thread_running, keyboard_listener
    if overlay_window_handle:
        if GLOBAGL_DEBUG:
            print("Closing overlay window")
        win32gui.PostMessage(overlay_window_handle, win32con.WM_CLOSE, 0, 0)
        overlay_window_handle = None
        overlay_thread_running = False
    if keyboard_listener:
        keyboard_listener.stop()
        keyboard_listener = None
def bring_to_front(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
# Debug
def log_window_info(hwnd, window_title):
    if hwnd:
        rect = win32gui.GetWindowRect(hwnd)
        styles = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        ex_styles = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        print(f"Window '{window_title}' Info: Handle: {hwnd}, Rect: {rect}, Styles: {styles}, ExStyles: {ex_styles}")
def print_colorful(text, color="random"):
    console = Console()
    colors = ['red', 'green', 'yellow', 'blue', 'magenta', 'cyan', 'white']

    for char in text:
        if color == "random":
            chosen_color = random.choice(colors)
        else:
            chosen_color = color
        style = Style(color=chosen_color)
        console.print(char, style=style, end='')
    console.print()

def check_active_window():
    global overlay_window_handle
    last_mod_time = os.path.getmtime(config_file_path)
    config = load_config()
    window_title = config.get('window_title', '')
    other_title_array = config.get('other_title_array', [])
    overlay_color = config.get('overlay_color', 'red')
    transparency = config.get('transparency', 0.5)
    print_colorful("D2ColorBot Started","random")
    while True:
        current_mod_time = os.path.getmtime(config_file_path)
        if current_mod_time != last_mod_time:
            print_colorful("Config file has changed. Reloading...","green")
            last_mod_time = current_mod_time
            config = load_config()
            window_title = config.get('window_title', '')
            other_title_array = config.get('other_title_array', [])
            overlay_color = config.get('overlay_color', 'red')
            transparency = config.get('transparency', 0.5)
        active_window = gw.getActiveWindow()
        if active_window:
            current_window_title = active_window.title
            if GLOBAGL_DEBUG:
                print(f"Current active window: {current_window_title}")
            if window_title in current_window_title:
                if GLOBAGL_DEBUG:
                    print(f"Window title '{window_title}' is active!")
                if not overlay_window_handle:
                    if GLOBAGL_DEBUG:
                        print(f"Creating overlay for '{current_window_title}'")
                    threading.Thread(target=create_click_through_window, args=(active_window.left, active_window.top, active_window.width, active_window.height, overlay_color, transparency)).start()
            elif any(ui_title in current_window_title for ui_title in other_title_array):
                if GLOBAGL_DEBUG:
                    print(f"UI element '{current_window_title}' is active! Keeping the overlay open.")
                bring_to_front(overlay_window_handle)
            elif overlay_window_handle:
                if GLOBAGL_DEBUG:
                    print_colorful("No specified title is active. Closing the overlay window.","red")
                close_overlay_window()
        time.sleep(GLOBAL_DELAY)

# Main
def main():
    try:
        listener_thread = threading.Thread(target=start_keyboard_listener)
        listener_thread.start()
        check_active_window()
        listener_thread.join()
    except Exception as e:
        print_colorful("Error in script: {e}","red")

if __name__ == "__main__":
    threading.Thread(target=keep_overlay_on_top).start()
    main()