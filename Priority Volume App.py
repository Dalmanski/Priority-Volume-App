import sys
import os
import json
import ctypes
from ctypes import wintypes
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QSlider, QHBoxLayout, QScrollArea, QPushButton, QFrame, QSizePolicy, QSpinBox
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPixmap, QIcon
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

class SHFILEINFO(ctypes.Structure):
    _fields_ = [("hIcon", wintypes.HICON), ("iIcon", wintypes.INT), ("dwAttributes", wintypes.DWORD), ("szDisplayName", wintypes.WCHAR * 260), ("szTypeName", wintypes.WCHAR * 80)]

class AppRow(QFrame):
    def __init__(self, pid, name, vol_iface, icon_pixmap, controller):
        super().__init__()
        self.pid = pid
        self.name = name
        self.vol_iface = vol_iface
        self.controller = controller
        self.setFixedHeight(56)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFrameShape(QFrame.NoFrame)
        self.layout = QHBoxLayout()
        self.layout.setContentsMargins(8, 6, 8, 6)
        self.layout.setSpacing(8)
        self.icon_label = QLabel()
        self.icon_label.setFixedSize(40, 40)
        if icon_pixmap is not None:
            self.icon_label.setPixmap(icon_pixmap.scaled(40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.name_label = QLabel(f"{self.name}")
        self.name_label.setMinimumWidth(160)
        self.name_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.slider = QSlider(Qt.Horizontal)
        self.slider.setRange(0, 100)
        self.slider.setFixedHeight(22)
        self.percent_label = QLabel("")
        self.percent_label.setFixedWidth(44)
        self.layout.addWidget(self.icon_label, 0)
        self.layout.addWidget(self.name_label, 2)
        self.layout.addWidget(self.slider, 6)
        self.layout.addWidget(self.percent_label, 0)
        self.setLayout(self.layout)
        try:
            v = int(round(self.vol_iface.GetMasterVolume() * 100))
        except Exception:
            v = 0
        self.slider.setValue(v)
        self.percent_label.setText(f"{v}%")
        self.name_label.mousePressEvent = self.on_click
        self.icon_label.mousePressEvent = self.on_click
        self.slider.sliderPressed.connect(self.on_slider_pressed)
        self.slider.valueChanged.connect(self.on_slider_changed)
        self.set_normal_style()

    def set_normal_style(self):
        self.setStyleSheet("QFrame{background-color:transparent;border-bottom:1px solid #e6e6e6;margin:0;padding:6px 8px;} QLabel{font-size:12px;}")

    def set_selected_style(self):
        self.setStyleSheet("QFrame{background-color:#dff1ff;border:1px solid #9fc5ff;border-radius:6px;margin:0;padding:6px 8px;} QLabel{font-size:12px;}")

    def on_click(self, ev):
        self.controller.set_priority_by_pid(self.pid)

    def on_slider_pressed(self):
        self.controller.set_priority_by_pid(self.pid)

    def on_slider_changed(self, val):
        self.percent_label.setText(f"{val}%")
        try:
            self.vol_iface.SetMasterVolume(val / 100.0, None)
        except Exception:
            pass
        if self.controller.priority_pid == self.pid:
            self.controller.priority_locked_to_target = False

    def update_volume_display(self, vol):
        block = self.slider.blockSignals(True)
        self.slider.setValue(int(round(vol * 100)))
        self.percent_label.setText(f"{int(round(vol * 100))}%")
        self.slider.blockSignals(block)

class VolumeController(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")
        self.load_settings()
        self.setWindowTitle("Per-App Priority Volume")
        self.resize(820, 640)
        try:
            self.setWindowIcon(QIcon("favicon.ico"))
        except Exception:
            pass
        self.container_widget = QWidget()
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(6,6,6,6)
        self.main_layout.setSpacing(8)
        self.container_widget.setLayout(self.main_layout)
        self.top_bar = QWidget()
        self.top_layout = QHBoxLayout()
        self.top_layout.setContentsMargins(0,0,0,0)
        self.top_layout.setSpacing(8)
        self.top_bar.setLayout(self.top_layout)
        self.btn_all_100 = QPushButton("All 100%")
        self.btn_all_0 = QPushButton("All 0%")
        self.btn_auto_priority = QPushButton("Auto Priority Volume")
        self.btn_auto_priority.setCheckable(True)
        self.btn_auto_priority.setChecked(self.settings.get("auto_priority_enabled", False))
        self.btn_auto_priority.setStyleSheet(self.auto_style(self.btn_auto_priority.isChecked()))
        self.top_layout.addWidget(self.btn_all_100)
        self.top_layout.addWidget(self.btn_all_0)
        self.top_layout.addWidget(self.btn_auto_priority)
        self.top_layout.addStretch()
        self.main_layout.addWidget(self.top_bar)
        info_row = QWidget()
        info_layout = QHBoxLayout()
        info_layout.setContentsMargins(2,0,2,0)
        info_layout.setSpacing(8)
        info_row.setLayout(info_layout)
        info_label = QLabel("Click an app to prioritize its volume: 100% on the selected app and 20% on other apps.")
        info_label.setWordWrap(True)
        info_label.setMinimumHeight(36)
        info_layout.addWidget(info_label, 6)
        spin_container = QWidget()
        spin_layout = QHBoxLayout()
        spin_layout.setContentsMargins(0,0,0,0)
        spin_layout.setSpacing(6)
        spin_container.setLayout(spin_layout)
        pri_label = QLabel("Priority (%)")
        self.spin_priority = QSpinBox()
        self.spin_priority.setRange(0,100)
        self.spin_priority.setValue(self.settings.get("priority_percent", 100))
        other_label = QLabel("Other apps (%)")
        self.spin_other = QSpinBox()
        self.spin_other.setRange(0,100)
        self.spin_other.setValue(self.settings.get("background_percent", 20))
        spin_layout.addWidget(pri_label)
        spin_layout.addWidget(self.spin_priority)
        spin_layout.addWidget(other_label)
        spin_layout.addWidget(self.spin_other)
        info_layout.addWidget(spin_container, 2)
        self.main_layout.addWidget(info_row)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.list_container = QWidget()
        self.list_layout = QVBoxLayout()
        self.list_layout.setContentsMargins(0,0,0,0)
        self.list_layout.setSpacing(0)
        self.list_container.setLayout(self.list_layout)
        self.scroll.setWidget(self.list_container)
        self.main_layout.addWidget(self.scroll)
        self.setCentralWidget(self.container_widget)
        self.sessions = {}
        self.rows = {}
        self.icons_cache = {}
        self.priority_pid = None
        self.priority_locked_to_target = True
        self.priority_percent = int(self.spin_priority.value())
        self.background_percent = int(self.spin_other.value())
        self.btn_all_100.clicked.connect(self.set_all_100)
        self.btn_all_0.clicked.connect(self.set_all_0)
        self.btn_auto_priority.toggled.connect(self.on_auto_toggled)
        self.spin_priority.valueChanged.connect(self.on_priority_spin_changed)
        self.spin_other.valueChanged.connect(self.on_other_spin_changed)
        self.refresh_sessions()
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.poll)
        self.timer.start()

    def load_settings(self):
        default = {"priority_percent": 100, "background_percent": 20, "auto_priority_enabled": False}
        try:
            if os.path.exists(self.settings_path):
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    d = json.load(f)
                default.update(d)
        except Exception:
            pass
        self.settings = default

    def save_settings(self):
        try:
            self.settings["priority_percent"] = int(self.priority_percent)
            self.settings["background_percent"] = int(self.background_percent)
            self.settings["auto_priority_enabled"] = bool(self.btn_auto_priority.isChecked()) if hasattr(self, "btn_auto_priority") else self.settings.get("auto_priority_enabled", False)
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except Exception:
            pass

    def auto_style(self, checked):
        if checked:
            return "QPushButton{background-color:#dff1ff;border:1px solid #9fc5ff;border-radius:6px;padding:6px 10px;}QPushButton:checked{background-color:#dff1ff;}"
        else:
            return "QPushButton{background-color:transparent;border:1px solid #cccccc;border-radius:6px;padding:6px 10px;}"

    def on_auto_toggled(self, checked):
        self.btn_auto_priority.setStyleSheet(self.auto_style(checked))
        self.save_settings()
        if checked:
            fg = self.get_foreground_pid()
            if fg and fg in self.sessions:
                self.set_priority_by_pid(fg)

    def on_priority_spin_changed(self, val):
        self.priority_percent = int(val)
        self.save_settings()
        self.enforce_priority()

    def on_other_spin_changed(self, val):
        self.background_percent = int(val)
        self.save_settings()
        self.enforce_priority()

    def poll(self):
        self.refresh_sessions()
        if self.btn_auto_priority.isChecked():
            fg = self.get_foreground_pid()
            if fg and fg in self.sessions and fg != self.priority_pid:
                self.set_priority_by_pid(fg)

    def get_foreground_pid(self):
        try:
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            if not hwnd:
                return None
            pid = wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return int(pid.value)
        except Exception:
            return None

    def get_icon_for_pid(self, pid, proc):
        try:
            pexe = None
            try:
                pexe = proc.exe()
            except Exception:
                pexe = None
            if pexe and pexe in self.icons_cache:
                return self.icons_cache[pexe]
            if pexe:
                try:
                    shfi = SHFILEINFO()
                    SHGFI_ICON = 0x000000100
                    SHGFI_LARGEICON = 0x000000000
                    res = ctypes.windll.shell32.SHGetFileInfoW(pexe, 0, ctypes.byref(shfi), ctypes.sizeof(shfi), SHGFI_ICON | SHGFI_LARGEICON)
                    if res and shfi.hIcon:
                        hicon = shfi.hIcon
                        pix = QPixmap.fromWinHICON(int(hicon))
                        ctypes.windll.user32.DestroyIcon(hicon)
                        self.icons_cache[pexe] = pix
                        return pix
                except Exception:
                    pass
                try:
                    large = (wintypes.HICON * 1)()
                    small = (wintypes.HICON * 1)()
                    cnt = ctypes.windll.shell32.ExtractIconExW(pexe, 0, large, small, 1)
                    if cnt > 0 and large[0]:
                        hicon = large[0]
                        pix = QPixmap.fromWinHICON(int(hicon))
                        ctypes.windll.user32.DestroyIcon(hicon)
                        self.icons_cache[pexe] = pix
                        return pix
                except Exception:
                    pass
        except Exception:
            pass
        return None

    def refresh_sessions(self):
        current = {}
        sessions = AudioUtilities.GetAllSessions()
        for s in sessions:
            proc = s.Process
            if proc is None:
                continue
            pid = proc.pid
            try:
                name = proc.name()
            except Exception:
                name = f"pid:{pid}"
            try:
                vol_iface = s._ctl.QueryInterface(ISimpleAudioVolume)
            except Exception:
                continue
            current[pid] = (name, vol_iface, proc)
        removed = set(self.sessions.keys()) - set(current.keys())
        for pid in removed:
            self.remove_row(pid)
            del self.sessions[pid]
        added = set(current.keys()) - set(self.sessions.keys())
        for pid in added:
            name, vol_iface, proc = current[pid]
            self.sessions[pid] = vol_iface
            icon = self.get_icon_for_pid(pid, proc)
            self.add_row(pid, name, vol_iface, icon)
        for pid in current.keys():
            try:
                vol_iface = current[pid][1]
                cur = vol_iface.GetMasterVolume()
                if pid in self.rows:
                    self.rows[pid].update_volume_display(cur)
            except Exception:
                pass

    def add_row(self, pid, name, vol_iface, icon_pixmap):
        row = AppRow(pid, name, vol_iface, icon_pixmap, self)
        self.list_layout.addWidget(row)
        self.rows[pid] = row
        if self.priority_pid is not None and pid != self.priority_pid:
            try:
                vol_iface.SetMasterVolume(max(0.0, min(1.0, self.background_percent / 100.0)), None)
                row.update_volume_display(max(0.0, min(1.0, self.background_percent / 100.0)))
                row.set_normal_style()
            except Exception:
                pass
        elif self.priority_pid == pid:
            row.set_selected_style()

    def remove_row(self, pid):
        if pid in self.rows:
            row = self.rows[pid]
            self.list_layout.removeWidget(row)
            row.deleteLater()
            del self.rows[pid]
            if self.priority_pid == pid:
                self.priority_pid = None
                self.priority_locked_to_target = True
                self.enforce_priority()

    def set_priority_by_pid(self, pid):
        if pid not in self.sessions:
            return
        self.priority_pid = pid
        self.priority_locked_to_target = True
        self.enforce_priority()

    def enforce_priority(self):
        pri_val = max(0.0, min(1.0, self.priority_percent / 100.0))
        bg_val = max(0.0, min(1.0, self.background_percent / 100.0))
        for pid, vol_iface in self.sessions.items():
            try:
                if self.priority_pid is None:
                    if pid in self.rows:
                        self.rows[pid].set_normal_style()
                    continue
                if pid == self.priority_pid:
                    if pid in self.rows:
                        self.rows[pid].set_selected_style()
                        if self.priority_locked_to_target:
                            vol_iface.SetMasterVolume(pri_val, None)
                            self.rows[pid].update_volume_display(pri_val)
                        else:
                            current = self.rows[pid].slider.value() / 100.0
                            vol_iface.SetMasterVolume(current, None)
                            self.rows[pid].update_volume_display(current)
                else:
                    vol_iface.SetMasterVolume(bg_val, None)
                    if pid in self.rows:
                        self.rows[pid].update_volume_display(bg_val)
                        self.rows[pid].set_normal_style()
            except Exception:
                pass

    def set_all_100(self):
        self.priority_pid = None
        self.priority_locked_to_target = True
        for pid, vol_iface in self.sessions.items():
            try:
                vol_iface.SetMasterVolume(1.0, None)
                if pid in self.rows:
                    self.rows[pid].update_volume_display(1.0)
                    self.rows[pid].set_normal_style()
            except Exception:
                pass

    def set_all_0(self):
        self.priority_pid = None
        self.priority_locked_to_target = True
        for pid, vol_iface in self.sessions.items():
            try:
                vol_iface.SetMasterVolume(0.0, None)
                if pid in self.rows:
                    self.rows[pid].update_volume_display(0.0)
                    self.rows[pid].set_normal_style()
            except Exception:
                pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    try:
        app.setWindowIcon(QIcon("favicon.ico"))
    except Exception:
        pass
    win = VolumeController()
    win.show()
    sys.exit(app.exec_())
