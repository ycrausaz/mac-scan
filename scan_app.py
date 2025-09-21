import sys, os, datetime, pathlib, yaml, shutil, traceback, subprocess, threading, re
from typing import Dict, Any, Optional, Tuple

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QComboBox, QPushButton,
    QFileDialog, QMessageBox, QHBoxLayout, QLineEdit, QDialog, QFormLayout,
    QDialogButtonBox, QSpinBox, QCheckBox, QMenu, QSystemTrayIcon, QStyle
)
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QObject

import requests  # for HTTP reachability probe

import PyPDF2
from PyQt6.QtWidgets import QProgressDialog

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
            "Version 1.4 • © 2025 You", self
        )
        body.setTextFormat(Qt.TextFormat.RichText)
        body.setWordWrap(True)
        layout.addWidget(body)

        features = QLabel(
            "Features:<ul>"
            "<li>Class picker & custom filename patterns</li>"
            "<li>Menu bar quick actions</li>"
            "<li>Open file & open location</li>"
            "<li>Preferences: IP/MAC, DPI, color, page size</li>"
            "<li>Manual input source selection (ADF/Flatbed)</li>"
            "<li>Network status (green/amber/red with MAC verify)</li>"
            "<li>Manual duplex scanning support</li>"
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

        self.ed_page = QComboBox(self)
        self.ed_page.addItems(["A4", "Letter", "Legal", "A5", "A3"])
        self.ed_page.setEditable(True)
        self.ed_page.setCurrentText(cfg.get("scanner", {}).get("page_size", "A4"))
        frm.addRow("Page size:", self.ed_page)

        # Input source selection
        self.ed_source = QComboBox(self)
        self.ed_source.addItems(["Auto", "Feeder (ADF)", "Platen (Flatbed)"])
        current_source = cfg.get("scanner", {}).get("input_source", "Auto")
        self.ed_source.setCurrentText(current_source)
        frm.addRow("Input source:", self.ed_source)

        # Add help text for input source
        source_hint = QLabel("Auto: detect based on document presence. Feeder: force ADF use. Platen: force flatbed use.", self)
        source_hint.setStyleSheet("color: #666; font-size: 11px;")
        source_hint.setWordWrap(True)
        frm.addRow("", source_hint)

        self.ed_pattern = QLineEdit(self)
        self.ed_pattern.setText(cfg.get("ui", {}).get("filename_pattern", "{class}_{date}_{time}.pdf"))
        frm.addRow("Filename pattern:", self.ed_pattern)

        self.chk_remember = QCheckBox("Remember last class for one-click scanning")
        self.chk_remember.setChecked(bool(cfg.get("ui", {}).get("remember_last_class", True)))
        frm.addRow(self.chk_remember)

        # Debug mode checkbox
        self.chk_debug = QCheckBox("Enable debug output (check console for scanner status)")
        self.chk_debug.setChecked(bool(cfg.get("ui", {}).get("debug_mode", False)))
        frm.addRow(self.chk_debug)

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
        self.cfg["scanner"]["page_size"] = self.ed_page.currentText().strip()
        self.cfg["scanner"]["input_source"] = self.ed_source.currentText()
        self.cfg["ui"]["filename_pattern"] = self.ed_pattern.text().strip() or "{class}_{date}_{time}.pdf"
        self.cfg["ui"]["remember_last_class"] = bool(self.chk_remember.isChecked())
        self.cfg["ui"]["debug_mode"] = bool(self.chk_debug.isChecked())

class _NetProbe(QObject):
    done = pyqtSignal(bool, bool, object)  # reachable, mac_matches, seen_mac

    def __init__(self, owner: "ScanApp"):
        super().__init__()
        self._owner = owner

    def run_once(self):
        """
        Run the owner's check and emit a Qt signal with the result.
        This function is safe to call from any thread.
        """
        try:
            reachable, mac_matches, seen_mac = self._owner._check_printer_once()
        except Exception:
            reachable, mac_matches, seen_mac = (False, False, None)
        # Emit the result; Qt will deliver to slots on the main thread
        self.done.emit(reachable, mac_matches, seen_mac)


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

        # Row 2: buttons
        row2 = QHBoxLayout()
        self.btn_scan = QPushButton("Scan (Single-sided)")
        self.btn_scan.clicked.connect(lambda: self.on_scan())
        self.btn_scan.setShortcut("Return")
        self.btn_scan.setToolTip("Scan single-sided documents")
        row2.addWidget(self.btn_scan)
        
        self.btn_manual_duplex = QPushButton("Manual Duplex")
        self.btn_manual_duplex.clicked.connect(lambda: self.on_scan_manual_duplex())
        self.btn_manual_duplex.setToolTip("Scan front sides, then back sides, then combine")
        row2.addWidget(self.btn_manual_duplex)
        
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

    def debug_print(self, msg: str):
        """Print debug message if debug mode is enabled"""
        if self.config.get("ui", {}).get("debug_mode", False):
            print(f"[DEBUG] {msg}")

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
        1) GET http://<ip>/eSCL/ScannerStatus with timeout, no proxies
        2) best-effort ARP for MAC
        """
        ip = self._extract_ip()
        seen_mac: Optional[str] = None
    
        # HTTP probe (proxy-free)
        reachable = False
        try:
            url = f"http://{ip}/eSCL/ScannerStatus"
            r = requests.get(
                url,
                timeout=2.0,
                headers={"Connection": "close"},
                allow_redirects=False,
                proxies={"http": None, "https": None},
            )
            reachable = (200 <= r.status_code < 500)
            self.debug_print(f"HTTP probe to {url}: status={r.status_code}, reachable={reachable}")
        except Exception as e:
            reachable = False
            self.debug_print(f"HTTP probe failed: {e}")
    
        # ARP (best-effort)
        try:
            arp_bin = "/usr/sbin/arp" if os.path.exists("/usr/sbin/arp") else "arp"
            out = subprocess.run([arp_bin, "-n", ip], capture_output=True, text=True, check=False, timeout=2.0)
            seen_mac = self._parse_arp_mac((out.stdout or "") + (out.stderr or ""))
            self.debug_print(f"ARP lookup for {ip}: seen_mac={seen_mac}")
        except Exception as e:
            seen_mac = None
            self.debug_print(f"ARP lookup failed: {e}")
    
        expected = self._normalize_mac(self.config["scanner"].get("mac", ""))
        mac_matches = bool(seen_mac and expected and seen_mac == expected)
        self.debug_print(f"MAC comparison: expected={expected}, seen={seen_mac}, matches={mac_matches}")
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
    
        # Create (or recreate) the probe and connect its signal to our UI updater
        self._net_probe = _NetProbe(self)
        self._net_probe.done.connect(self._update_net_ui)
    
        # Kick off the first check immediately, in a background thread
        threading.Thread(target=self._net_probe.run_once, daemon=True).start()
    
        # Periodic checks every 10s
        self._net_timer = QTimer(self)
        self._net_timer.setInterval(10_000)
        self._net_timer.timeout.connect(lambda: threading.Thread(target=self._net_probe.run_once, daemon=True).start())
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
            self.status.setText(f"Updated folder for \"{cls}\" → {chosen}")
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

    def _determine_scan_source(self) -> str:
        """Determine which input source to use based on preferences"""
        input_source = self.config.get("scanner", {}).get("input_source", "Auto")
        
        if input_source == "Feeder (ADF)":
            self.debug_print("Using forced Feeder (ADF) source")
            return "Feeder"
        elif input_source == "Platen (Flatbed)":
            self.debug_print("Using forced Platen (Flatbed) source")
            return "Platen"
        else:  # Auto
            self.debug_print("Auto-detecting input source...")
            try:
                # Try to get scanner status for auto-detection
                status = self.scanner().get_status()
                self.debug_print(f"Scanner status XML: {status}")
                source = self.scanner().choose_input_source()
                self.debug_print(f"Auto-detected source: {source}")
                return source
            except Exception as e:
                self.debug_print(f"Auto-detection failed: {e}, defaulting to Platen")
                return "Platen"

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

            # Determine input source
            source = self._determine_scan_source()

            self.status.setText(f"Scanning to {out_path} …")
            QApplication.processEvents()

            # Scan parameters (single-sided only)
            scan_params = {
                "dpi": int(s.get("dpi", 300)),
                "color_mode": s.get("color_mode", "Color"),
                "page_size": s.get("page_size", "A4"),
                "input_source": source,
            }
            
            self.debug_print(f"Single-sided scan parameters: {scan_params}")

            # Perform single-sided scan
            self.scanner().scan_to_pdf(out_path, **scan_params)

            self.last_saved_path = out_path
            self.btn_open_file.setEnabled(True)
            self.btn_open_loc.setEnabled(True)

            # Show success message with scan details
            scan_details = f"Source: {source}, {scan_params['dpi']} DPI, {scan_params['color_mode']}"

            self.status.setText(f"Saved: {out_path}")
            QMessageBox.information(
                self, 
                "Scan complete", 
                f"Saved to:\n{out_path}\n\n{scan_details}"
            )

            if self.config.get("ui", {}).get("remember_last_class", True):
                self.config.setdefault("ui", {})["last_class"] = cls
                save_config(self.config)
                self.rebuild_tray_menu()

        except Exception as e:
            tb = traceback.format_exc()
            self.status.setText("Scan failed.")
            self.debug_print(f"Scan error: {tb}")
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

    def on_scan_manual_duplex(self, cls: Optional[str] = None):
        """
        Manual duplex scanning: scan front sides, then back sides, then combine.
        """
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
    
            base_filename = make_filename(pattern, cls=cls, topic=topic).replace(".pdf", "")
            front_path = os.path.join(target_dir, f"{base_filename}_front.pdf")
            back_path = os.path.join(target_dir, f"{base_filename}_back.pdf")
            final_path = os.path.join(target_dir, f"{base_filename}.pdf")
    
            # Scan parameters (force Feeder for consistent multi-page scanning)
            scan_params = {
                "dpi": int(s.get("dpi", 300)),
                "color_mode": s.get("color_mode", "Color"),
                "page_size": s.get("page_size", "A4"),
                "input_source": "Feeder",  # Force ADF for consistent multi-page scanning
            }
    
            # Step 1: Scan front sides
            reply = QMessageBox.question(
                self, 
                "Manual Duplex - Step 1", 
                "Load your documents in the ADF with the FRONT sides facing down.\n\n"
                "Click OK when ready to scan the front sides.",
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
            )
            if reply == QMessageBox.StandardButton.Cancel:
                return
    
            self.status.setText("Scanning front sides...")
            QApplication.processEvents()
    
            self.debug_print(f"Scanning front sides with params: {scan_params}")
            self.scanner().scan_to_pdf(front_path, **scan_params)
    
            # Step 2: Scan back sides
            reply = QMessageBox.question(
                self, 
                "Manual Duplex - Step 2", 
                f"Front sides saved to:\n{front_path}\n\n"
                "Now FLIP your documents and load them in the ADF with the BACK sides facing down.\n"
                "Make sure they're in the REVERSE order (last page first).\n\n"
                "Click OK when ready to scan the back sides.",
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
            )
            if reply == QMessageBox.StandardButton.Cancel:
                return
    
            self.status.setText("Scanning back sides...")
            QApplication.processEvents()
    
            self.debug_print(f"Scanning back sides with params: {scan_params}")
            self.scanner().scan_to_pdf(back_path, **scan_params)
    
            # Step 3: Combine PDFs
            self.status.setText("Combining front and back sides...")
            QApplication.processEvents()
    
            self._combine_duplex_pdfs(front_path, back_path, final_path)
    
            # Cleanup temporary files
            try:
                os.remove(front_path)
                os.remove(back_path)
            except Exception:
                pass
    
            self.last_saved_path = final_path
            self.btn_open_file.setEnabled(True)
            self.btn_open_loc.setEnabled(True)
    
            self.status.setText(f"Manual duplex complete: {final_path}")
            QMessageBox.information(
                self, 
                "Manual Duplex Complete", 
                f"Combined duplex document saved to:\n{final_path}\n\n"
                f"Scanned with: {scan_params['dpi']} DPI, {scan_params['color_mode']}"
            )
    
            if self.config.get("ui", {}).get("remember_last_class", True):
                self.config.setdefault("ui", {})["last_class"] = cls
                save_config(self.config)
                self.rebuild_tray_menu()
    
        except Exception as e:
            tb = traceback.format_exc()
            self.status.setText("Manual duplex scan failed.")
            self.debug_print(f"Manual duplex error: {tb}")
            QMessageBox.critical(self, "Error", f"{e}\n\nDetails:\n{tb}")
    
    def _combine_duplex_pdfs(self, front_path: str, back_path: str, output_path: str):
        """
        Combine front and back PDF pages into a single duplex document.
        Assumes back pages are in reverse order.
        Fixed version that keeps files open during processing.
        """
        try:
            # Open both files and keep them open during the entire process
            front_file = open(front_path, 'rb')
            back_file = open(back_path, 'rb')
            
            try:
                # Read both PDFs
                front_pdf = PyPDF2.PdfReader(front_file)
                back_pdf = PyPDF2.PdfReader(back_file)
                
                # Get page counts
                front_count = len(front_pdf.pages)
                back_count = len(back_pdf.pages)
                
                self.debug_print(f"Front PDF: {front_count} pages, Back PDF: {back_count} pages")
                
                # Create output PDF
                writer = PyPDF2.PdfWriter()
                
                # Get back pages in reverse order (they should be scanned in reverse)
                back_pages_reversed = list(reversed(back_pdf.pages))
                
                # Interleave front and back pages
                max_pages = max(front_count, back_count)
                for i in range(max_pages):
                    # Add front page
                    if i < front_count:
                        self.debug_print(f"Adding front page {i+1}")
                        writer.add_page(front_pdf.pages[i])
                    
                    # Add corresponding back page (from reversed list)
                    if i < back_count:
                        self.debug_print(f"Adding back page {i+1} (original page {back_count-i})")
                        writer.add_page(back_pages_reversed[i])
    
                # Write combined PDF
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
    
                self.debug_print(f"Successfully combined {front_count} front + {back_count} back pages into {output_path}")
    
            finally:
                # Always close the input files
                front_file.close()
                back_file.close()
    
        except Exception as e:
            raise RuntimeError(f"Failed to combine PDFs: {e}")

# ---------- main ----------

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Class Scanner")
    app.setWindowIcon(QIcon())  # bundle icon used when packaged
    cfg = load_config()
    cfg.setdefault("scanner", {}).setdefault("mac", "")  # ensure key exists
    cfg.setdefault("scanner", {}).setdefault("input_source", "Auto")  # ensure input_source exists
    cfg.setdefault("ui", {}).setdefault("debug_mode", False)  # ensure debug_mode exists
    w = ScanApp(cfg)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
