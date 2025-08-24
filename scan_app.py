import sys, os, datetime, pathlib, yaml, shutil, traceback, subprocess, threading, re
from typing import Dict, Any, Optional, Tuple

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QComboBox, QPushButton,
    QFileDialog, QMessageBox, QHBoxLayout, QLineEdit, QDialog, QFormLayout,
    QDialogButtonBox, QSpinBox, QCheckBox, QMenu, QSystemTrayIcon, QStyle
)
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtCore import Qt, QTimer

import requests  # for HTTP reachability probe

from escl_client import ESCLScanner


# ---------- Config helpers (App Support) ----------

def app_support_dir() -> pathlib.Path:
    base = pathlib.Path.home() / "Library" / "Application Support" / "Class Scanner"
    base.mkdir(parents=True, exist_ok=True)
    return base

def config_path() -> pathlib.Path:
    return app_support_dir() / "config.yaml"

def load_config() -> Dict[str, Any]:
    cfg_path = config_path()
    if not cfg_path.exists():
        here = pathlib.Path(getattr(sys, "_MEIPASS", pathlib.Path(__file__).parent))
        default_cfg = here / "default_config.yaml"
        shutil.copy(default_cfg, cfg_path)
    with open(cfg_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def save_config(cfg: Dict[str, Any]) -> None:
    with open(config_path(), "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, allow_unicode=True)


# ---------- Filename pattern ----------

def make_filename(pattern: str, cls: str, topic: str = "", ext: str = "pdf") -> str:
    """
    Tokens: {class} {date} {time} {topic} (and optional {ext})
    Avoids dangling underscore when topic is empty.
    """
    now = datetime.datetime.now()
    safe_class = cls.replace(" ", "_")
    date = now.strftime("%Y-%m-%d")
    time = now.strftime("%H-%M-%S")
    topic_clean = topic.strip().replace(" ", "_") if topic else ""

    values = {"class": safe_class, "date": date, "time": time, "topic": topic_clean, "ext": ext}
    name = pattern.format(**values)

    if not topic_clean and "{topic}" in pattern:
        stem, suffix = os.path.splitext(name)
        stem = stem.rstrip("_-")
        name = stem + (suffix or "." + ext)

    if not name.lower().endswith("." + ext.lower()):
        name = f"{name}.{ext}"
    return name


# ---------- About dialog ----------

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Class Scanner")
        layout = QVBoxLayout(self)

        title = QLabel("<h2>Class Scanner</h2>", self)
        title.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(title)

        body = QLabel(
            "Scan class documents directly to per-class folders using your network scanner (AirScan/eSCL).<br>"
            "Version 1.2 • © 2025 You", self
        )
        body.setTextFormat(Qt.TextFormat.RichText)
        body.setWordWrap(True)
        layout.addWidget(body)

        features = QLabel(
            "Features:<ul>"
            "<li>Class picker & custom filename patterns</li>"
            "<li>Menu bar quick actions</li>"
            "<li>Open file & open location</li>"
            "<li>Preferences: IP/MAC, DPI, color, duplex, page size</li>"
            "<li>Network status (green/amber/red with MAC verify)</li>"
            "</ul>", self
        )
        features.setTextFormat(Qt.TextFormat.RichText)
        features.setWordWrap(True)
        layout.addWidget(features)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok, self)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


# ---------- Preferences dialog ----------

class PreferencesDialog(QDialog):
    def __init__(self, cfg: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preferences")
        self.cfg = cfg

        frm = QFormLayout(self)

        self.ed_host = QLineEdit(self)
        self.ed_host.setText(str(cfg.get("scanner", {}).get("host", "")))
        frm.addRow("Scanner IP / Host:", self.ed_host)

        self.ed_mac = QLineEdit(self)
        self.ed_mac.setPlaceholderText("e.g. 84:2A:FD:A6:F2:B0 (optional, for verification)")
        self.ed_mac.setText(cfg.get("scanner", {}).get("mac", ""))
        frm.addRow("Scanner MAC:", self.ed_mac)

        self.spin_dpi = QSpinBox(self)
        self.spin_dpi.setRange(75, 1200)
        self.spin_dpi.setSingleStep(25)
        self.spin_dpi.setValue(int(cfg.get("scanner", {}).get("dpi", 300)))
        frm.addRow("DPI:", self.spin_dpi)

        self.ed_color = QComboBox(self)
        self.ed_color.addItems(["Color", "Grayscale"])
        self.ed_color.setCurrentText(cfg.get("scanner", {}).get("color_mode", "Color"))
        frm.addRow("Color mode:", self.ed_color)

        self.chk_duplex = QCheckBox("Use duplex (ADF double-sided)")
        self.chk_duplex.setChecked(bool(cfg.get("scanner", {}).get("duplex", False)))
        frm.addRow(self.chk_duplex)

        self.ed_page = QComboBox(self)
        self.ed_page.addItems(["A4", "Letter", "Legal", "A5", "A3"])
        self.ed_page.setEditable(True)
        self.ed_page.setCurrentText(cfg.get("scanner", {}).get("page_size", "A4"))
        frm.addRow("Page size:", self.ed_page)

        self.ed_pattern = QLineEdit(self)
        self.ed_pattern.setText(cfg.get("ui", {}).get("filename_pattern", "{class}_{date}_{time}.pdf"))
        frm.addRow("Filename pattern:", self.ed_pattern)

        self.chk_remember = QCheckBox("Remember last class for one-click scanning")
        self.chk_remember.setChecked(bool(cfg.get("ui", {}).get("remember_last_class", True)))
        frm.addRow(self.chk_remember)

        hint = QLabel("Tokens: {class} {date} {time} {topic} (and {ext})", self)
        hint.setStyleSheet("color: #666;")
        frm.addRow("", hint)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        frm.addRow(buttons)

    def apply(self):
        self.cfg.setdefault("scanner", {})
        self.cfg.setdefault("ui", {})
        self.cfg["scanner"]["host"] = self.ed_host.text().strip()
        self.cfg["scanner"]["mac"] = self.ed_mac.text().strip()
        self.cfg["scanner"]["dpi"] = int(self.spin_dpi.value())
        self.cfg["scanner"]["color_mode"] = self.ed_color.currentText()
        self.cfg["scanner"]["duplex"] = bool(self.chk_duplex.isChecked())
        self.cfg["scanner"]["page_size"] = self.ed_page.currentText().strip()
        self.cfg["ui"]["filename_pattern"] = self.ed_pattern.text().strip() or "{class}_{date}_{time}.pdf"
        self.cfg["ui"]["remember_last_class"] = bool(self.chk_remember.isChecked())


# ---------- Main app ----------

class ScanApp(QWidget):
    def __init__(self, config: Dict[str, Any]):
        super().__init__()
        self.config = config
        self.setWindowTitle("Class Scanner (macOS)")
        self.setMinimumWidth(620)

        self._scanner: Optional[ESCLScanner] = None
        self.last_saved_path: Optional[str] = None

        layout = QVBoxLayout(self)

        header = QLabel("Choose class and scan:", self)
        layout.addWidget(header)

        # Row 1: class + topic
        row1 = QHBoxLayout()
        self.class_combo = QComboBox(self)
        self.class_combo.addItems(list(self.config["classes"].keys()))
        last_cls = self.config.get("ui", {}).get("last_class", "")
        if last_cls and last_cls in self.config["classes"]:
            self.class_combo.setCurrentText(last_cls)
        row1.addWidget(self.class_combo, 2)

        self.ed_topic = QLineEdit(self)
        self.ed_topic.setPlaceholderText("Optional topic (e.g., lecture3, homework2)")
        row1.addWidget(self.ed_topic, 3)
        layout.addLayout(row1)

        # Row 2: buttons (Scan / Open folder / Prefs / About)
        row2 = QHBoxLayout()
        self.btn_scan = QPushButton("Scan")
        self.btn_scan.clicked.connect(lambda: self.on_scan())  # lambda avoids stray 'checked' bool
        self.btn_scan.setShortcut("Return")
        row2.addWidget(self.btn_scan)

        self.btn_open = QPushButton("Open folder…")
        self.btn_open.clicked.connect(self.on_open_folder)
        row2.addWidget(self.btn_open)

        self.btn_prefs = QPushButton("Preferences…")
        self.btn_prefs.clicked.connect(self.on_prefs)
        row2.addWidget(self.btn_prefs)

        self.btn_about = QPushButton("About…")
        self.btn_about.clicked.connect(self.on_about)
        row2.addWidget(self.btn_about)

        layout.addLayout(row2)

        # Status label
        self.status = QLabel("", self)
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        # Row 3: network reachability (dot + label)
        netrow = QHBoxLayout()
        self.net_flag = QLabel("●")
        self.net_flag.setStyleSheet("font-size: 18px; color: #999;")  # gray initially
        self.net_label = QLabel("Checking printer…")
        netrow.addWidget(self.net_flag)
        netrow.addWidget(self.net_label)
        netrow.addStretch(1)
        layout.addLayout(netrow)

        # Row 4: Open file / Open location
        btnrow = QHBoxLayout()
        self.btn_open_file = QPushButton("Open file")
        self.btn_open_file.setEnabled(False)
        self.btn_open_file.clicked.connect(self.on_open_file)
        btnrow.addWidget(self.btn_open_file)

        self.btn_open_loc = QPushButton("Open location")
        self.btn_open_loc.setEnabled(False)
        self.btn_open_loc.clicked.connect(self.on_open_location)
        btnrow.addWidget(self.btn_open_loc)
        layout.addLayout(btnrow)

        # Tray (menu bar extra)
        self.tray: Optional[QSystemTrayIcon] = None
        self.setup_tray()

        # Events
        self.class_combo.currentTextChanged.connect(self.update_last_class)

        # Start the network monitor
        self._net_timer: Optional[QTimer] = None
        self.start_net_monitor()

    # ---- Helpers ----

    def scanner(self) -> ESCLScanner:
        host = self.config["scanner"]["host"]
        if self._scanner is None or host not in getattr(self._scanner, "base", ""):
            self._scanner = ESCLScanner(host)
        return self._scanner

    def current_target_dir(self) -> str:
        cls = self.class_combo.currentText()
        return self.config["classes"][cls]

    def ensure_dir(self, p: str) -> None:
        pathlib.Path(p).mkdir(parents=True, exist_ok=True)

    def update_last_class(self, cls: str):
        if self.config.get("ui", {}).get("remember_last_class", True):
            self.config.setdefault("ui", {})["last_class"] = cls
            save_config(self.config)
            self.rebuild_tray_menu()

    # ---- Network reachability ----

    @staticmethod
    def _normalize_mac(mac: str) -> str:
        return mac.lower().replace("-", ":").strip()

    @staticmethod
    def _parse_arp_mac(text: str) -> Optional[str]:
        m = re.search(r"\b([0-9a-fA-F]{2}(?::[0-9a-fA-F]{2}){5})\b", text)
        return ScanApp._normalize_mac(m.group(1)) if m else None

    def _extract_ip(self) -> str:
        host = self.config["scanner"]["host"].strip()
        if host.startswith("http://"):
            host = host[len("http://"):]
        elif host.startswith("https://"):
            host = host[len("https://"):]
        host = host.strip("/")
        ip = host.split("/")[0]
        return ip

    def _check_printer_once(self) -> tuple[bool, bool, Optional[str]]:
        """
        Returns (reachable, mac_matches, seen_mac).
    
        Strategy:
          1) HTTP GET to /eSCL/ScannerStatus with a hard 2s deadline and NO proxies.
          2) Best-effort ARP MAC lookup (also hard timeout).
          3) Never throws; always returns quickly.
        """
        ip = self._extract_ip()
        seen_mac: Optional[str] = None
    
        # 1) HTTP probe (proxy-free)
        reachable = False
        try:
            url = f"http://{ip}/eSCL/ScannerStatus"
            r = requests.get(
                url,
                timeout=2.0,
                headers={"Connection": "close"},
                allow_redirects=False,
                proxies={"http": None, "https": None},  # <- ignore system/corp proxies
            )
            # Any 2xx/3xx/4xx means *something* responded at that IP
            reachable = (200 <= r.status_code < 500)
        except Exception:
            reachable = False
    
        # 2) ARP lookup (best-effort)
        try:
            arp_bin = "/usr/sbin/arp" if os.path.exists("/usr/sbin/arp") else "arp"
            out = subprocess.run(
                [arp_bin, "-n", ip],
                capture_output=True,
                text=True,
                check=False,
                timeout=2.0
            )
            seen_mac = self._parse_arp_mac((out.stdout or "") + (out.stderr or ""))
        except Exception:
            seen_mac = None
    
        expected = self._normalize_mac(self.config["scanner"].get("mac", ""))
        mac_matches = bool(seen_mac and expected and seen_mac == expected)
        return reachable, mac_matches, seen_mac

    def _update_net_ui(self, reachable: bool, mac_matches: bool, seen_mac: Optional[str]):
        if not reachable:
            self.net_flag.setStyleSheet("font-size: 18px; color: #d22;")  # red
            self.net_label.setText("Printer unreachable")
            self.net_label.setToolTip("No HTTP/ARP response from the device.")
        elif reachable and mac_matches:
            self.net_flag.setStyleSheet("font-size: 18px; color: #2a2;")  # green
            self.net_label.setText("Printer OK (IP & MAC match)")
            self.net_label.setToolTip(f"Matched MAC: {seen_mac}")
        else:
            self.net_flag.setStyleSheet("font-size: 18px; color: #e6a100;")  # amber
            self.net_label.setText("Warning: IP reachable, MAC mismatch")
            exp = self.config["scanner"].get("mac", "")
            tip = f"ARP reports {seen_mac or 'unknown'}, expected {exp or 'not set'}"
            self.net_label.setToolTip(tip)

    def start_net_monitor(self):
        # stop previous timer if any
        if hasattr(self, "_net_timer") and self._net_timer is not None:
            try:
                self._net_timer.stop()
            except Exception:
                pass
    
        def worker():
            try:
                reachable, mac_matches, seen_mac = self._check_printer_once()
            except Exception:
                reachable, mac_matches, seen_mac = (False, False, None)
            # Always post a UI update back to the main thread
            QTimer.singleShot(0, lambda: self._update_net_ui(reachable, mac_matches, seen_mac))
    
        # Run once immediately so it doesn't linger on "Checking…"
        threading.Thread(target=worker, daemon=True).start()
    
        # Periodic checks every 10s
        self._net_timer = QTimer(self)
        self._net_timer.setInterval(10_000)
        self._net_timer.timeout.connect(lambda: threading.Thread(target=worker, daemon=True).start())
        self._net_timer.start()

    # ---- Actions ----

    def on_open_folder(self):
        cls = self.class_combo.currentText()
        target = self.current_target_dir()
        start = target if os.path.isdir(target) else str(pathlib.Path.home())
        chosen = QFileDialog.getExistingDirectory(self, f"Choose folder for {cls}", start)
        if chosen:
            self.config["classes"][cls] = chosen
            save_config(self.config)
            self.status.setText(f"Updated folder for “{cls}” → {chosen}")
            self.rebuild_tray_menu()

    def on_prefs(self):
        dlg = PreferencesDialog(self.config, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            dlg.apply()
            save_config(self.config)
            self._scanner = None  # pick up new host next time
            self.status.setText("Preferences saved.")
            self.rebuild_tray_menu()
            self.start_net_monitor()  # restart with new IP/MAC

    def on_about(self):
        AboutDialog(self).exec()

    def on_scan(self, cls: Optional[str] = None):
        try:
            if cls is None:
                cls = self.class_combo.currentText()
            if cls not in self.config["classes"]:
                raise RuntimeError(f"Unknown class: {cls}")

            target_dir = self.config["classes"][cls]
            self.ensure_dir(target_dir)

            s = self.config["scanner"]
            ui = self.config.get("ui", {})
            pattern = ui.get("filename_pattern", "{class}_{date}_{time}.pdf")
            topic = self.ed_topic.text().strip()

            filename = make_filename(pattern, cls=cls, topic=topic)
            out_path = os.path.join(target_dir, filename)

            self.status.setText(f"Scanning to {out_path} …")
            QApplication.processEvents()

            self.scanner().scan_to_pdf(
                out_path,
                dpi=int(s.get("dpi", 300)),
                color_mode=s.get("color_mode", "Color"),
                duplex=bool(s.get("duplex", False)),
                page_size=s.get("page_size", "A4"),
            )

            self.last_saved_path = out_path
            self.btn_open_file.setEnabled(True)
            self.btn_open_loc.setEnabled(True)

            self.status.setText(f"Saved: {out_path}")
            QMessageBox.information(self, "Scan complete", f"Saved to:\n{out_path}")

            if self.config.get("ui", {}).get("remember_last_class", True):
                self.config.setdefault("ui", {})["last_class"] = cls
                save_config(self.config)
                self.rebuild_tray_menu()

        except Exception as e:
            tb = traceback.format_exc()
            self.status.setText("Scan failed.")
            QMessageBox.critical(self, "Error", f"{e}\n\nDetails:\n{tb}")

    def on_open_file(self):
        if not self.last_saved_path or not os.path.exists(self.last_saved_path):
            QMessageBox.warning(self, "No file", "There is no scanned file to open yet.")
            return
        try:
            subprocess.run(["open", self.last_saved_path], check=False)
        except Exception as e:
            QMessageBox.critical(self, "Open failed", str(e))

    def on_open_location(self):
        if self.last_saved_path and os.path.exists(self.last_saved_path):
            try:
                subprocess.run(["open", "-R", self.last_saved_path], check=False)  # reveal in Finder
                return
            except Exception:
                pass
        try:
            folder = self.current_target_dir()
            if folder and os.path.isdir(folder):
                subprocess.run(["open", folder], check=False)
                return
        except Exception:
            pass
        QMessageBox.warning(self, "No location", "There is no scanned file or target folder to open yet.")

    # ---- Tray (menu bar extra) ----

    def setup_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return
        self.tray = QSystemTrayIcon(self)
        app_icon = QApplication.windowIcon()
        if not app_icon.isNull():
            self.tray.setIcon(app_icon)
        else:
            self.tray.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
        self.rebuild_tray_menu()
        self.tray.show()

    def rebuild_tray_menu(self):
        if not self.tray:
            return
        menu = QMenu()

        last_cls = self.config.get("ui", {}).get("last_class", "")
        act_scan_last = QAction(f"Scan to Last Class ({last_cls})" if last_cls else "Scan to Last Class", self)
        act_scan_last.triggered.connect(lambda checked=False, c=last_cls: self.on_scan(cls=c or None))
        menu.addAction(act_scan_last)

        sub = QMenu("Scan to…", menu)
        for cls in self.config["classes"].keys():
            a = QAction(cls, sub)
            a.triggered.connect(lambda checked=False, c=cls: self.on_scan(cls=c))
            sub.addAction(a)
        menu.addMenu(sub)

        menu.addSeparator()

        a_about = QAction("About…", menu)
        a_about.triggered.connect(self.on_about)
        menu.addAction(a_about)

        a_open = QAction("Open Main Window", menu)
        a_open.triggered.connect(self.showNormal)
        menu.addAction(a_open)

        a_prefs = QAction("Preferences…", menu)
        a_prefs.triggered.connect(self.on_prefs)
        menu.addAction(a_prefs)

        a_quit = QAction("Quit", menu)
        a_quit.triggered.connect(QApplication.instance().quit)
        menu.addAction(a_quit)

        self.tray.setContextMenu(menu)


# ---------- main ----------

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Class Scanner")
    app.setWindowIcon(QIcon())  # bundle icon used when packaged
    cfg = load_config()
    cfg.setdefault("scanner", {}).setdefault("mac", "")  # ensure key exists
    w = ScanApp(cfg)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

