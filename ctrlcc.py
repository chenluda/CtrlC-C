import os
import sys
import pyperclip
import keyboard
from threading import Timer
from pystray import MenuItem as item
import pystray
from PIL import Image, ImageDraw
from ctypes import windll
import winreg as reg
import win32event
import win32api
from winerror import ERROR_ALREADY_EXISTS
import logging
import subprocess


def get_log_file_path():
    """Get the path for the log file, differentiating between script and executable."""
    if getattr(sys, 'frozen', False):
        home_dir = os.path.expanduser('~')
        log_dir = os.path.join(home_dir, 'ctrlcc')
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        return os.path.join(log_dir, 'CtrlC_C_log.txt')
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CtrlC_C_log.txt')


def view_logs(icon, item):
    """Open the log file with the default text editor."""
    try:
        if os.name == 'nt':  # Windows
            os.startfile(log_filepath)
        elif os.name == 'posix':  # Unix-like
            subprocess.run(['open', log_filepath], check=True)
        else:
            logging.error("Unsupported OS for viewing logs")
    except Exception as e:
        logging.error(f"Error opening log file: {e}")


def add_to_startup():
    """Add the application to the Windows startup registry."""
    if getattr(sys, 'frozen', False):
        app_path = sys.executable
    else:
        app_path = os.path.abspath(__file__)

    try:
        key = reg.HKEY_CURRENT_USER
        key_value = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        open = reg.OpenKey(key, key_value, 0, reg.KEY_ALL_ACCESS)
        reg.SetValueEx(open, "CtrlC+C", 0, reg.REG_SZ, app_path)
        reg.CloseKey(open)
        return True
    except WindowsError:
        return False


def remove_from_startup():
    """Remove the application from the Windows startup registry."""
    try:
        key = reg.HKEY_CURRENT_USER
        key_value = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        open = reg.OpenKey(key, key_value, 0, reg.KEY_ALL_ACCESS)
        reg.DeleteValue(open, "CtrlC+C")
        reg.CloseKey(open)
        return True
    except WindowsError:
        return False


def toggle_startup():
    """Toggle whether the application is in the Windows startup registry."""
    if is_in_startup():
        return remove_from_startup()
    else:
        return add_to_startup()


def is_in_startup():
    """Check whether the application is in the Windows startup registry."""
    try:
        key = reg.HKEY_CURRENT_USER
        key_value = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        open = reg.OpenKey(key, key_value, 0, reg.KEY_READ)
        reg.QueryValueEx(open, "CtrlC+C")
        reg.CloseKey(open)
        return True
    except WindowsError:
        return False


def show_message_box(title, message):
    """Show a message box with the given title and message."""
    return windll.user32.MessageBoxW(0, message, title, 0)


def strip_newlines(text):
    """Remove all types of newline characters from a string."""
    try:
        text = text.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        logging.info(f"Copied text: {text}")  # Log the copied text
        return text
    except Exception as e:
        logging.error(f"Error while accessing clipboard: {e}")
        return ""


def get_clipboard_text():
    """Get the current text on the clipboard and then clear it."""
    try:
        text = pyperclip.paste()
        pyperclip.copy("")  # Clear the clipboard
        return text
    except Exception as e:
        logging.error(f"Error while accessing clipboard: {e}")
        return ""


def set_clipboard_text(text):
    """Set the given text on the clipboard."""
    pyperclip.copy(text)


def on_c_press(event):
    global strip_newlines_attempted
    """Start a timer after first 'C' press, if 'Ctrl' is held down"""
    if keyboard.is_pressed('ctrl'):
        if hasattr(on_c_press, 'first_press_timer') and on_c_press.first_press_timer is not None:
            # If timer exists, we are within 0.5 seconds of first press, so we execute the action
            on_c_press.first_press_timer.cancel()
            strip_newlines_attempted = True
            perform_clipboard_action()
            Timer(0.1, check_conflict).start()
        else:
            # Start a timer that waits for another 'C' press within 0.5 seconds
            on_c_press.first_press_timer = Timer(0.5, reset_first_press_timer)
            on_c_press.first_press_timer.start()
    else:
        reset_first_press_timer()


def check_conflict():
    """Check if strip_newlines was attempted but not executed, indicating a hotkey conflict."""
    global strip_newlines_attempted, strip_newlines_executed
    if strip_newlines_attempted and not strip_newlines_executed:
        # If strip_newlines was attempted but not executed, then we might have a conflict
        show_message_box("快捷键冲突", "Ctrl+C+C 快捷键可能被其他程序占用。")
    # Reset flags for the next attempt
    strip_newlines_attempted = False
    strip_newlines_executed = False


def perform_clipboard_action():
    """Perform the action of copying the clipboard text and removing newlines."""
    global strip_newlines_executed
    reset_first_press_timer()
    try:
        current_data = get_clipboard_text()
        stripped_text = strip_newlines(current_data)
        if stripped_text != current_data:
            set_clipboard_text(stripped_text)
            logging.info("Newlines removed from clipboard text.")
            strip_newlines_executed = True
    except Exception as e:
        logging.error(f"Error in perform_clipboard_action: {e}")


def reset_first_press_timer():
    """Reset the timer for double 'C' press"""
    on_c_press.first_press_timer = None


def create_icon(image_name):
    """The application is frozen"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    # The application is not frozen
    else:
        base_path = os.path.abspath(".")

    # Correct the path to the icon file
    icon_path = os.path.join(base_path, image_name)
    return Image.open(icon_path)


def setup_tray_icon():
    """Load your own image as icon and setup tray icon with menu"""
    icon_image = create_icon("ctrlcc.ico")

    # The menu that will appear when the user right-clicks the icon
    menu = (item('Toggle Start on Boot', toggle_startup, checked=lambda item: is_in_startup()),
            item('View Logs', view_logs),
            item('Exit', exit_program),)
    icon = pystray.Icon("test_icon", icon_image, "CtrlC+C", menu)
    icon.run()


def exit_program(icon, item):
    """Exit the program and stop the system tray icon"""
    icon.stop()  # This will stop the system tray icon and the associated message loop.
    keyboard.unhook_all()
    # print("Exiting program...")


if __name__ == "__main__":

    log_filepath = get_log_file_path()
    logging.basicConfig(filename=log_filepath, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')

    strip_newlines_attempted = False
    strip_newlines_executed = False

    mutex_name = "CtrlC_C_Application_Mutex"

    mutex = win32event.CreateMutex(None, False, mutex_name)
    last_error = win32api.GetLastError()

    if last_error == ERROR_ALREADY_EXISTS:
        show_message_box("应用程序已运行", "CtrlC+C 应用程序已经在运行了。")
        sys.exit(0)

    # Show instruction message box
    instruction_message = (
        "欢迎使用 CtrlC+C ！\n"
        "\n"
        "【按住 Ctrl 键并按 C 键两次】以复制文本并删除换行。\n"
        "\n"
        "如有问题，请联系：chenluda01@gmail.com"
    )
    show_message_box("© 2023 Glenn.", instruction_message)

    # Initialization and event hooks
    # print("Hold Ctrl and press C twice to copy text and remove newlines...")
    # print("Press Esc to quit.")
    on_c_press.first_press_timer = None
    keyboard.on_release_key('c', on_c_press)
    # keyboard.add_hotkey('esc', lambda: exit_program(None, None))  # Adapted for direct call
    setup_tray_icon()  # Start the system tray icon
