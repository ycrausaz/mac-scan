import sys, os, datetime, pathlib, yaml, shutil, traceback, subprocess
from typing import Dict, Any, Optional

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QComboBox, QPushButton,
    QFileDialog, QMessageBox, QHBoxLayout, QLineEdit, QDialog, QFormLayout,
    QDialogButtonBox, QSpinBox, QCheckBox, QMenu, QSystemTrayIcon, QStyle
)
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtCore import Qt

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
        # Seed from bundled default_config.yaml (PyInstaller: sys._MEIPASS)
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
    Build a filename from a pattern with tokens:
      {class} {date} {time} {topic} (and optional {ext})
    Ensures no trailing underscore if topic is empty and {topic} is in the pattern.
    """
    now = datetime.datetime.now()
    safe_class = cls.replace(" ", "_")
    date = now.strftime("%Y-%m-%d")
    time = now.strftime("%H-%M-%S")
    topic_clean = topic.strip().replace(" ", "_") if topic else ""

    values = {
        "class": safe_class,
        "date": date,
        "time": time,
        "topic": topic_clean,
        "ext": ext,
    }

    name = pattern.format(**values)

    # If topic is empty but pattern had {topic}, remove trailing _ or - before extension
    if not topic_clean and "{topic}" in pattern:
        stem, suffix = os.path.splitext(name)
        stem = stem.rstrip("_-")
        name = stem + (suffix or "." + ext)

    # Ensure extension
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
            "A tiny macOS app to scan to per-class folders using your network scanner (AirScan/eSCL).<br>"
            "Version 0.99 • © 2025, ton Papounet",
            self
        )
        body.setWordWrap(True)
        body.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(body)

        details = QLabel(
            "Features:<ul>"
            "<li>Class picker, custom filename patterns</li>"
            "<li>Menu bar quick actions</li>"
            "<li>Open file & open location</li>"
            "<li>Preferences: IP/DPI/color/duplex/page size</li>"
            "<li>Combine ADF pages into one PDF (device-supported)</li>"
            "</ul>",
            self
        )
        details.setTextFormat(Qt.TextFormat.RichText)
        details.setWordWrap(True)
        layout.addWidget(details)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok, self)
        btns.accepted.connect(self.accept)
        layout.addWidget(btns)


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
        self.ed_page.setEditable(True)  # allow custom text
        self.ed_page.setCurrentText(cfg.get("scanner", {}).get("page_size", "A4"))
        frm.addRow("Page size:", self.ed_page)

        self.ed_pattern = QLineEdit(self)
        self.ed_pattern.setText(cfg.get("ui", {}).get("filename_pattern", "{class}_{date}_{time}.pdf"))
        frm.addRow("Filename pattern:", self.ed_pattern)

        self.chk_remember = QCheckBox("Remember last class for one-click scanning")
        self.chk_remember.setChecked(bool(cfg.get("ui", {}).get("remember_last_class", True)))
        frm.addRow(self.chk_remember)

        self.chk_combine = QCheckBox("Combine ADF pages into one PDF")
        self.chk_combine.setChecked(bool(cfg.get("ui", {}).get("combine_adf", True)))
        frm.addRow(self.chk_combine)

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
        self.cfg["scanner"]["dpi"] = int(self.spin_dpi.value())
        self.cfg["scanner"]["color_mode"] = self.ed_color.currentText()
        self.cfg["scanner"]["duplex"] = bool(self.chk_duplex.isChecked())
        self.cfg["scanner"]["page_size"] = self.ed_page.currentText().strip()
        self.cfg["ui"]["filename_pattern"] = self.ed_pattern.text().strip() or "{class}_{date}_{time}.pdf"
        self.cfg["ui"]["remember_last_class"] = bool(self.chk_remember.isChecked())
        self.cfg["ui"]["combine_adf"] = bool(self.chk_combine.isChecked())


# ---------- Main app ----------

class ScanApp(QWidget):
    def __init__(self, config: Dict[str, Any]):
        super().__init__()
        self.config = config
        self.setWindowTitle("Class Scanner (macOS)")
        self.setMinimumWidth(580)

        self._scanner: Optional[ESCLScanner] = None
        self.last_saved_path: Optional[str] = None

        layout = QVBoxLayout(self)

        header = QLabel("Choose class and scan:", self)
        layout.addWidget(header)

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

        row2 = QHBoxLayout()
        self.btn_scan = QPushButton("Scan")
        # Robust slot signature; can receive a bool (checked) safely
        self.btn_scan.clicked.connect(self.on_scan)
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

        self.status = QLabel("", self)
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        # Buttons next to status: Open file / Open location
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

        self.class_combo.currentTextChanged.connect(self.update_last_class)

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

    def on_about(self):
        AboutDialog(self).exec()

    def on_scan(self, *_, cls: Optional[str] = None):
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

            # Start scan (most eSCL devices produce a single multi-page PDF from ADF)
            self.scanner().scan_to_pdf(
                out_path,
                dpi=int(s.get("dpi", 300)),
                color_mode=s.get("color_mode", "Color"),
                duplex=bool(s.get("duplex", False)),
                page_size=s.get("page_size", "A4"),
            )

            # If “combine ADF” is on, you could add post-processing here in the future.
            # For HP eSCL, the PDF already includes all pages from the ADF in one file.

            self.last_saved_path = out_path
            self.btn_open_file.setEnabled(True)
            self.btn_open_loc.setEnabled(True)

            self.status.setText(f"Saved: {out_path}")
            QMessageBox.information(self, "Scan complete", f"Saved to:\n{out_path}")

            self.update_last_class(cls)

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
                # Reveal the file in Finder
                subprocess.run(["open", "-R", self.last_saved_path], check=False)
                return
            except Exception:
                pass
        # Fallback: open current target directory
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
    app.setWindowIcon(QIcon())  # use bundle icon when packaged
    cfg = load_config()
    # Ensure 'combine_adf' has a default in config if missing
    cfg.setdefault("ui", {}).setdefault("combine_adf", True)
    w = ScanApp(cfg)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

