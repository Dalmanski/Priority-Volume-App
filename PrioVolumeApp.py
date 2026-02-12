import sys
import ctypes
from ctypes import wintypes
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QSlider, QHBoxLayout, QScrollArea, QPushButton, QFrame, QSizePolicy
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPixmap
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import psutil

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
            self.controller.priority_locked_to_100 = False

    def update_volume_display(self, vol):
        block = self.slider.blockSignals(True)
        self.slider.setValue(int(round(vol * 100)))
        self.percent_label.setText(f"{int(round(vol * 100))}%")
        self.slider.blockSignals(block)

class VolumeController(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Per-App Priority Volume")
        self.resize(820, 640)
        self.container_widget = QWidget()
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(6,6,6,6)
        self.main_layout.setSpacing(6)
        self.container_widget.setLayout(self.main_layout)
        self.top_bar = QWidget()
        self.top_layout = QHBoxLayout()
        self.top_layout.setContentsMargins(0,0,0,0)
        self.top_layout.setSpacing(8)
        self.top_bar.setLayout(self.top_layout)
        self.btn_all_100 = QPushButton("All 100%")
        self.btn_all_0 = QPushButton("All 0%")
        self.top_layout.addWidget(self.btn_all_100)
        self.top_layout.addWidget(self.btn_all_0)
        self.top_layout.addStretch()
        self.main_layout.addWidget(self.top_bar)
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
        self.priority_locked_to_100 = True
        self.btn_all_100.clicked.connect(self.set_all_100)
        self.btn_all_0.clicked.connect(self.set_all_0)
        self.refresh_sessions()
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.poll)
        self.timer.start()

    def poll(self):
        self.refresh_sessions()

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
                vol_iface.SetMasterVolume(0.2, None)
                row.update_volume_display(0.2)
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
                self.priority_locked_to_100 = True
                self.enforce_priority()

    def set_priority_by_pid(self, pid):
        if pid not in self.sessions:
            return
        self.priority_pid = pid
        self.priority_locked_to_100 = True
        self.enforce_priority()

    def enforce_priority(self):
        for pid, vol_iface in self.sessions.items():
            try:
                if self.priority_pid is None:
                    if pid in self.rows:
                        self.rows[pid].set_normal_style()
                    continue
                if pid == self.priority_pid:
                    if pid in self.rows:
                        self.rows[pid].set_selected_style()
                        if self.priority_locked_to_100:
                            vol_iface.SetMasterVolume(1.0, None)
                            self.rows[pid].update_volume_display(1.0)
                        else:
                            current = self.rows[pid].slider.value() / 100.0
                            vol_iface.SetMasterVolume(current, None)
                            self.rows[pid].update_volume_display(current)
                else:
                    vol_iface.SetMasterVolume(0.2, None)
                    if pid in self.rows:
                        self.rows[pid].update_volume_display(0.2)
                        self.rows[pid].set_normal_style()
            except Exception:
                pass

    def set_all_100(self):
        self.priority_pid = None
        self.priority_locked_to_100 = True
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
        self.priority_locked_to_100 = True
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
    win = VolumeController()
    win.show()
    sys.exit(app.exec_())
