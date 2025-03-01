import sys
import platform
import pygetwindow as gw
import time
from datetime import datetime
import win32con
import win32gui
import ctypes
from ctypes import wintypes
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu, QAction, 
                          QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                          QLabel, QSpinBox, QPushButton, QDesktopWidget, QCheckBox, QMessageBox)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))

    return os.path.join(base_path, relative_path)

icon_path = resource_path("icon.ico")

class SettingsWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('PiP Pin')
        self.setFixedSize(300, 200)
        self.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        refresh_layout = QHBoxLayout()
        refresh_layout.addWidget(QLabel('Delay (ms):'))
        self.refresh_spinbox = QSpinBox()
        self.refresh_spinbox.setRange(100, 5000)
        self.refresh_spinbox.setValue(1000)
        self.refresh_spinbox.setSingleStep(100)
        refresh_layout.addWidget(self.refresh_spinbox)
        
        self.pin_checkbox = QCheckBox('PiP Pin')
        self.pin_checkbox.setChecked(True)
        self.pin_checkbox.stateChanged.connect(self.toggle_pin)
        
        apply_button = QPushButton('Apply')
        apply_button.clicked.connect(self.apply_settings)
        
        layout.addLayout(refresh_layout)
        layout.addWidget(self.pin_checkbox)
        layout.addWidget(apply_button)

    def apply_settings(self):
        if hasattr(self, 'callback'):
            self.callback(self.refresh_spinbox.value())
        
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle('PiP Pin')
        msg_box.setText('Settings saved.')
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    def toggle_pin(self, state):
        if hasattr(self, 'toggle_callback'):
            self.toggle_callback(state == Qt.Checked)

class PiPPinner:
    def __init__(self):
        self.app = QApplication(sys.argv)
        
        if os.path.exists(icon_path):
            self.app.setWindowIcon(QIcon(icon_path))
        
        self.windows_version = self.get_windows_version()
        self.pip_window = None
        self.is_pinned = False
        self.setup_ui()
        self.setup_tray()
        self.setup_timer()
        
    def setup_ui(self):
        self.settings_window = SettingsWindow()
        self.settings_window.callback = self.update_refresh_rate
        self.settings_window.toggle_callback = self.set_pin_enabled
        self.settings_window.show()
        
    def setup_tray(self):
        self.tray = QSystemTrayIcon(QIcon(icon_path), self.app)
        
        self.tray.setToolTip('PiP Pin')
        self.menu = QMenu()
        
        settings_action = QAction('Settings', self.menu)
        settings_action.triggered.connect(self.show_settings)
        
        exit_action = QAction('Quit', self.menu)
        exit_action.triggered.connect(self.app.quit)
        
        self.menu.addAction(settings_action)
        self.menu.addSeparator()
        self.menu.addAction(exit_action)
        
        self.tray.setContextMenu(self.menu)
        self.tray.show()
    
    def setup_timer(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_pip)
        self.timer.start(1000)
    
    def show_settings(self):
        self.settings_window.show()
        self.settings_window.activateWindow()
    
    def update_refresh_rate(self, value):
        self.timer.setInterval(value)
        
    def get_windows_version(self):
        ver = sys.getwindowsversion()
        win_ver = platform.win32_ver()
        return {'version': "11" if ver.build >= 22000 else win_ver[0]}

    def is_pip_window(self, window):
        title = window.title.strip().lower()
        pip_keywords = [
            "picture-in-picture", 
            "화면 속 화면", 
            "pip mode",
            "pip 모드"
        ]
        
        return title and any(keyword in title for keyword in pip_keywords)

    def pin_window(self, hwnd):
        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
            ctypes.windll.user32.SetWindowDisplayAffinity(hwnd, 0x0000000B)
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style | win32con.WS_EX_TOOLWINDOW)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.is_pinned = True
        except: pass

    def unpin_window(self, hwnd):
        try:
            ctypes.windll.user32.SetWindowDisplayAffinity(hwnd, 0)
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style & ~win32con.WS_EX_TOOLWINDOW)
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.is_pinned = False
        except: pass

    def check_pip(self):
        try:
            found_pip = False
            for window in gw.getAllWindows():
                if self.is_pip_window(window):
                    found_pip = True
                    if not self.is_pinned:
                        self.pip_window = window
                        self.pin_window(window._hWnd)
                    break
            
            if not found_pip and self.is_pinned and self.pip_window:
                self.unpin_window(self.pip_window._hWnd)
                self.pip_window = None
        except: pass

    def set_pin_enabled(self, enabled):
        self.is_pinned = enabled
        if not enabled and self.pip_window:
            self.unpin_window(self.pip_window._hWnd)
            self.pip_window = None

    def run(self):
        return self.app.exec_()

if __name__ == "__main__":
    pinner = PiPPinner()
    sys.exit(pinner.run()) 