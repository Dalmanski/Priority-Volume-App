import sys
import os
import json
import ctypes
from ctypes import wintypes
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QSlider, QHBoxLayout, QScrollArea, QPushButton, QFrame, QSizePolicy, QSpinBox, QStyle
from PyQt5.QtCore import Qt, QTimer, QRect, QSize
from PyQt5.QtGui import QPixmap, QIcon, QPainter, QColor, QBrush, QPen
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

try:
    from pycaw.pycaw import IAudioMeterInformation
    METER_SUPPORTED = True
except Exception:
    IAudioMeterInformation = None
    METER_SUPPORTED = False

try:
    import platform
    if platform.system() == "Windows":
        import winreg
    else:
        winreg = None
except Exception:
    winreg = None


class SHFILEINFO(ctypes.Structure):
    _fields_ = [
        ("hIcon", wintypes.HICON),
        ("iIcon", wintypes.INT),
        ("dwAttributes", wintypes.DWORD),
        ("szDisplayName", wintypes.WCHAR * 260),
        ("szTypeName", wintypes.WCHAR * 80),
    ]


class VolumeMeter(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.level = 0.0
        self.setMinimumWidth(16)
        self.setMaximumWidth(22)
        self.setFixedHeight(40)

    def set_level(self, lvl):
        try:
            lv = max(0.0, min(1.0, float(lvl)))
        except Exception:
            lv = 0.0
        self.level = lv
        self.update()

    def paintEvent(self, event):
        pw = self.width()
        ph = self.height()
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        outer = QRect(1, 1, pw - 2, ph - 2)
        pen = QPen(QColor(140, 140, 140))
        pen.setWidth(1)
        painter.setPen(pen)
        painter.setBrush(QBrush(QColor(30, 30, 30, 10)))
        painter.drawRect(outer)
        inner_h = max(0, int((ph - 6) * self.level))
        fill_rect = QRect(3, ph - 3 - inner_h, pw - 6, inner_h)
        l = self.level
        if l <= 0.33:
            fill_color = QColor(50, 220, 90)
        elif l <= 0.66:
            fill_color = QColor(255, 165, 40)
        else:
            fill_color = QColor(220, 60, 60)
        painter.setBrush(QBrush(fill_color))
        painter.setPen(Qt.NoPen)
        painter.drawRect(fill_rect)
        painter.end()


class AppRow(QFrame):
    def __init__(self, pid, name, vol_iface, icon_pixmap, meter_iface, controller):
        super().__init__()
        self.pid = pid
        self.name = name
        self.vol_iface = vol_iface
        self.meter_iface = meter_iface
        self.controller = controller
        self.setFixedHeight(56)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFrameShape(QFrame.NoFrame)

        self.layout = QHBoxLayout()
        self.layout.setContentsMargins(8, 6, 8, 6)
        self.layout.setSpacing(8)

        self.meter = VolumeMeter()

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(40, 40)
        if icon_pixmap is not None:
            self.icon_label.setPixmap(icon_pixmap.scaled(40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        self.name_label = QLabel(f"{self.name}")
        self.name_label.setMinimumWidth(140)
        self.name_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)

        self.mute_button = QPushButton()
        self.mute_button.setCheckable(True)
        self.mute_button.setFixedSize(30, 28)
        self.mute_button.setIconSize(QSize(18, 18))
        self.mute_button.setCursor(Qt.PointingHandCursor)

        self.slider = QSlider(Qt.Horizontal)
        self.slider.setRange(0, 100)
        self.slider.setFixedHeight(22)

        self.percent_label = QLabel("")
        self.percent_label.setFixedWidth(44)

        self.layout.addWidget(self.meter, 0)
        self.layout.addWidget(self.icon_label, 0)
        self.layout.addWidget(self.name_label, 2)
        self.layout.addWidget(self.mute_button, 0)
        self.layout.addWidget(self.slider, 6)
        self.layout.addWidget(self.percent_label, 0)
        self.setLayout(self.layout)

        try:
            v = int(round(self.vol_iface.GetMasterVolume() * 100))
        except Exception:
            v = 0
        self.slider.setValue(v)
        self.percent_label.setText(f"{v}%")
        self.meter.set_level(0.0)

        if self.meter_iface is not None and METER_SUPPORTED:
            try:
                peak = float(self.meter_iface.GetPeakValue())
                self.meter.set_level(peak)
            except Exception:
                pass

        muted = False
        try:
            muted = bool(self.vol_iface.GetMute())
        except Exception:
            muted = False
        self.update_mute_display(muted)

        self.name_label.mousePressEvent = self.on_click
        self.icon_label.mousePressEvent = self.on_click
        self.slider.valueChanged.connect(self.on_slider_changed)
        self.mute_button.toggled.connect(self.on_mute_toggled)
        self.set_normal_style()

    def set_normal_style(self):
        self.setStyleSheet("QFrame{background-color:transparent;border-bottom:1px solid #e6e6e6;margin:0;padding:6px 8px;} QLabel{font-size:12px;}")

    def set_selected_style(self):
        self.setStyleSheet("QFrame{background-color:#dff1ff;border:1px solid #9fc5ff;border-radius:6px;margin:0;padding:6px 8px;} QLabel{font-size:12px;}")

    def set_mute_style(self, muted):
        if muted:
            self.mute_button.setStyleSheet("QPushButton{background-color:#ffe0e0;border:1px solid #ff9f9f;border-radius:6px;padding:0px;}QPushButton:checked{background-color:#ffe0e0;}")
        else:
            self.mute_button.setStyleSheet("QPushButton{background-color:transparent;border:1px solid #cccccc;border-radius:6px;padding:0px;}")

        if muted:
            icon = self.style().standardIcon(QStyle.SP_MediaVolumeMuted)
        else:
            icon = self.style().standardIcon(QStyle.SP_MediaVolume)

        self.mute_button.setIcon(icon)

    def update_mute_display(self, muted):
        blocked = self.mute_button.blockSignals(True)
        self.mute_button.setChecked(bool(muted))
        self.set_mute_style(bool(muted))
        self.mute_button.blockSignals(blocked)

    def on_click(self, ev):
        self.controller.set_priority_by_pid(self.pid)

    def mousePressEvent(self, event):
        try:
            pos = event.pos()
            if self.slider.geometry().contains(pos) or self.mute_button.geometry().contains(pos):
                super().mousePressEvent(event)
                return
        except Exception:
            pass
        self.controller.set_priority_by_pid(self.pid)
        super().mousePressEvent(event)

    def on_slider_changed(self, val):
        self.percent_label.setText(f"{val}%")
        try:
            self.vol_iface.SetMasterVolume(val / 100.0, None)
        except Exception:
            pass
        if self.controller.priority_pid == self.pid:
            self.controller.priority_locked_to_target = False

    def on_mute_toggled(self, checked):
        self.set_mute_style(bool(checked))
        try:
            self.vol_iface.SetMute(bool(checked), None)
        except Exception:
            pass

    def update_volume_display(self, vol):
        block = self.slider.blockSignals(True)
        self.slider.setValue(int(round(vol * 100)))
        self.percent_label.setText(f"{int(round(vol * 100))}%")
        self.slider.blockSignals(block)

    def update_meter_from_peak(self):
        if self.meter_iface is None or not METER_SUPPORTED:
            return
        try:
            peak = float(self.meter_iface.GetPeakValue())
            self.meter.set_level(peak)
        except Exception:
            pass


class VolumeController(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")
        self.load_settings()
        self.setWindowTitle("Per-App Priority Volume")
        self.resize(900, 640)
        try:
            self.setWindowIcon(QIcon("favicon.ico"))
        except Exception:
            pass

        self.container_widget = QWidget()
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(6, 6, 6, 6)
        self.main_layout.setSpacing(8)
        self.container_widget.setLayout(self.main_layout)

        self.top_bar = QWidget()
        self.top_layout = QHBoxLayout()
        self.top_layout.setContentsMargins(0, 0, 0, 0)
        self.top_layout.setSpacing(8)
        self.top_bar.setLayout(self.top_layout)

        self.btn_all_100 = QPushButton("All 100%")
        self.btn_all_0 = QPushButton("All 0%")
        self.btn_auto_priority = QPushButton("Auto Priority Volume")
        self.btn_auto_priority.setCheckable(True)
        self.btn_auto_priority.setChecked(self.settings.get("auto_priority_enabled", False))
        self.btn_auto_priority.setStyleSheet(self.auto_style(self.btn_auto_priority.isChecked()))

        self.btn_auto_100 = QPushButton("Auto 100% open app")
        self.btn_auto_100.setCheckable(True)
        self.btn_auto_100.setChecked(self.settings.get("auto_100_enabled", False))
        self.btn_auto_100.setStyleSheet(self.auto_style(self.btn_auto_100.isChecked()))

        self.btn_startup = QPushButton("Start with Windows")
        self.btn_startup.setCheckable(True)
        startup_state = self.get_startup_enabled()
        self.btn_startup.setChecked(bool(startup_state))
        self.btn_startup.setStyleSheet(self.auto_style(self.btn_startup.isChecked()))
        self.btn_startup.toggled.connect(self.on_startup_toggled)

        if winreg is None:
            self.btn_startup.setEnabled(False)

        self.top_layout.addWidget(self.btn_all_100)
        self.top_layout.addWidget(self.btn_all_0)
        self.top_layout.addWidget(self.btn_auto_priority)
        self.top_layout.addWidget(self.btn_auto_100)
        self.top_layout.addWidget(self.btn_startup)
        self.top_layout.addStretch()
        self.main_layout.addWidget(self.top_bar)

        info_row = QWidget()
        info_layout = QHBoxLayout()
        info_layout.setContentsMargins(2, 0, 2, 0)
        info_layout.setSpacing(8)
        info_row.setLayout(info_layout)

        info_label = QLabel("Click an app to prioritize its volume: 100% on the selected app and lower on other apps.")
        info_label.setWordWrap(True)
        info_label.setMinimumHeight(36)
        info_layout.addWidget(info_label, 6)

        spin_container = QWidget()
        spin_layout = QHBoxLayout()
        spin_layout.setContentsMargins(0, 0, 0, 0)
        spin_layout.setSpacing(6)
        spin_container.setLayout(spin_layout)

        pri_label = QLabel("Priority (%)")
        self.spin_priority = QSpinBox()
        self.spin_priority.setRange(0, 100)
        self.spin_priority.setValue(self.settings.get("priority_percent", 100))

        other_label = QLabel("Other apps (%)")
        self.spin_other = QSpinBox()
        self.spin_other.setRange(0, 100)
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
        self.list_layout.setContentsMargins(0, 0, 0, 0)
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
        self.btn_auto_100.toggled.connect(self.on_auto_100_toggled)
        self.spin_priority.valueChanged.connect(self.on_priority_spin_changed)
        self.spin_other.valueChanged.connect(self.on_other_spin_changed)

        self.refresh_sessions()

        self.poll_timer = QTimer()
        self.poll_timer.setInterval(1000)
        self.poll_timer.timeout.connect(self.poll)
        self.poll_timer.start()

        self.meter_timer = QTimer()
        self.meter_timer.setInterval(100)
        self.meter_timer.timeout.connect(self.update_meters)
        if METER_SUPPORTED:
            self.meter_timer.start()

    def load_settings(self):
        default = {
            "priority_percent": 100,
            "background_percent": 20,
            "auto_priority_enabled": False,
            "auto_100_enabled": False,
            "run_at_startup": False,
        }
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
            self.settings["auto_100_enabled"] = bool(self.btn_auto_100.isChecked()) if hasattr(self, "btn_auto_100") else self.settings.get("auto_100_enabled", False)
            self.settings["run_at_startup"] = bool(self.btn_startup.isChecked()) if hasattr(self, "btn_startup") else self.settings.get("run_at_startup", False)
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

    def on_auto_100_toggled(self, checked):
        self.btn_auto_100.setStyleSheet(self.auto_style(checked))
        self.save_settings()

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

    def update_meters(self):
        for pid, info in list(self.sessions.items()):
            try:
                row = self.rows.get(pid)
                if row:
                    row.update_meter_from_peak()
            except Exception:
                pass

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
            meter_iface = None
            if METER_SUPPORTED:
                try:
                    meter_iface = s._ctl.QueryInterface(IAudioMeterInformation)
                except Exception:
                    meter_iface = None
            current[pid] = (name, vol_iface, proc, meter_iface)

        removed = set(self.sessions.keys()) - set(current.keys())
        for pid in removed:
            self.remove_row(pid)
            try:
                del self.sessions[pid]
            except Exception:
                pass

        added = set(current.keys()) - set(self.sessions.keys())
        for pid in added:
            name, vol_iface, proc, meter_iface = current[pid]
            self.sessions[pid] = {"vol": vol_iface, "proc": proc, "meter": meter_iface}
            icon = self.get_icon_for_pid(pid, proc)
            self.add_row(pid, name, vol_iface, icon, meter_iface)
            try:
                if hasattr(self, "btn_auto_100") and self.btn_auto_100.isChecked():
                    vol_iface.SetMasterVolume(1.0, None)
                    if pid in self.rows:
                        self.rows[pid].update_volume_display(1.0)
            except Exception:
                pass

        for pid in current.keys():
            try:
                vol_iface = current[pid][1]
                cur = vol_iface.GetMasterVolume()
                muted = False
                try:
                    muted = bool(vol_iface.GetMute())
                except Exception:
                    muted = False
                if pid in self.rows:
                    self.rows[pid].update_volume_display(cur)
                    self.rows[pid].update_mute_display(muted)
            except Exception:
                pass

    def add_row(self, pid, name, vol_iface, icon_pixmap, meter_iface):
        row = AppRow(pid, name, vol_iface, icon_pixmap, meter_iface, self)
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
        for pid, info in self.sessions.items():
            try:
                vol_iface = info.get("vol")
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
        for pid, info in self.sessions.items():
            try:
                vol_iface = info.get("vol")
                vol_iface.SetMasterVolume(1.0, None)
                if pid in self.rows:
                    self.rows[pid].update_volume_display(1.0)
                    self.rows[pid].set_normal_style()
            except Exception:
                pass

    def set_all_0(self):
        self.priority_pid = None
        self.priority_locked_to_target = True
        for pid, info in self.sessions.items():
            try:
                vol_iface = info.get("vol")
                vol_iface.SetMasterVolume(0.0, None)
                if pid in self.rows:
                    self.rows[pid].update_volume_display(0.0)
                    self.rows[pid].set_normal_style()
            except Exception:
                pass

    def _expected_run_command(self):
        if getattr(sys, "frozen", False):
            return f'"{sys.executable}"'
        else:
            exe_path = sys.executable
            script = os.path.abspath(__file__)
            return f'"{exe_path}" "{script}"'

    def get_startup_enabled(self):
        try:
            if winreg is None:
                return bool(self.settings.get("run_at_startup", False))
            key = r"Software\Microsoft\Windows\CurrentVersion\Run"
            name = "PerAppPriorityVolume"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ) as rk:
                    val, typ = winreg.QueryValueEx(rk, name)
                    expected = self._expected_run_command()
                    if isinstance(val, str) and expected in val:
                        return True
                    else:
                        return True if isinstance(val, str) else False
            except FileNotFoundError:
                return False
            except OSError:
                return False
            except Exception:
                return bool(self.settings.get("run_at_startup", False))
        except Exception:
            return bool(self.settings.get("run_at_startup", False))

    def on_startup_toggled(self, checked):
        success = self.set_startup_enabled(checked)
        try:
            self.btn_startup.blockSignals(True)
            if success:
                self.btn_startup.setChecked(bool(checked))
                self.btn_startup.setStyleSheet(self.auto_style(bool(checked)))
            else:
                self.btn_startup.setChecked(not bool(checked))
                self.btn_startup.setStyleSheet(self.auto_style(not bool(checked)))
            self.btn_startup.blockSignals(False)
        except Exception:
            try:
                self.btn_startup.blockSignals(False)
            except Exception:
                pass
        self.save_settings()

    def set_startup_enabled(self, enable):
        try:
            self.save_settings()
            if winreg is None:
                return False
            key = r"Software\Microsoft\Windows\CurrentVersion\Run"
            name = "PerAppPriorityVolume"
            cmd = self._expected_run_command()
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ):
                    pass
            except FileNotFoundError:
                try:
                    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key):
                        pass
                except Exception:
                    pass
            if enable:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE) as rk:
                        winreg.SetValueEx(rk, name, 0, winreg.REG_SZ, cmd)
                    return True
                except PermissionError:
                    return False
                except Exception:
                    try:
                        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key) as rk:
                            winreg.SetValueEx(rk, name, 0, winreg.REG_SZ, cmd)
                        return True
                    except Exception:
                        return False
            else:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE) as rk:
                        try:
                            winreg.DeleteValue(rk, name)
                        except FileNotFoundError:
                            pass
                    return True
                except FileNotFoundError:
                    return True
                except PermissionError:
                    return False
                except Exception:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ) as rk:
                            try:
                                winreg.DeleteValue(rk, name)
                            except Exception:
                                pass
                        return True
                    except Exception:
                        return False
        except Exception:
            return False


if __name__ == "__main__":
    app = QApplication(sys.argv)
    try:
        app.setWindowIcon(QIcon("favicon.ico"))
    except Exception:
        pass
    win = VolumeController()
    win.show()
    sys.exit(app.exec_())