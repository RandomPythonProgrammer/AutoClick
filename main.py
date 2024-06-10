import time

import pygetwindow
import win32api
import win32con
import win32gui
import yaml
import keyboard


class Application:
    def __init__(self):
        self._current_application = None
        self._running = True
        self._start = time.time()
        self._timer = time.time()

    @property
    def _config(self):
        with open('config.yaml', 'r') as file:
            return yaml.load(file, yaml.Loader)

    def update(self):
        if keyboard.is_pressed('ctrl+alt+c'):
            time.sleep(5)
            self.select_window()
        if self._current_application is not None:
            if time.time() - self._start > self._config['duration']:
                self._running = False
            elif time.time() - self._timer > self._config['interval']:
                self.click()
                self._timer = time.time()

    def select_window(self):
        window = pygetwindow.getActiveWindow()
        self._current_application = window._hWnd

        self._timer = time.time()
        self._start = time.time()

    def start(self):
        while self._running:
            self.update()
            time.sleep(0.1)

    def click(self):
        position = win32api.MAKELONG(0, 0)
        win32gui.SendMessage(self._current_application, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, position)
        win32gui.SendMessage(self._current_application, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, position)


if __name__ == '__main__':
    app = Application()
    app.start()
