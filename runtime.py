# runtime.py - v2.8 - Force-add debugpy data folder to fix packaging
import sys, os, subprocess, urllib.request, hashlib, json, time, shutil
from pathlib import Path

# ================= CONFIG =================
APP_URL = "https://world.oshonet.in/runclap/combine/runtime/app.py"
ICON_URL = "https://world.oshonet.in/runclap/combine/runtime/logo/logo.png"
FINAL_APP_NAME = "Converter-Ultimate"

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QProgressBar,
    QPushButton, QTextEdit, QGroupBox, QCheckBox, QStackedWidget
)
from PyQt6.QtCore import QThread, pyqtSignal

# Application directory and paths
APPDATA = os.environ.get("APPDATA") or os.path.expanduser("~")
APP_DIR = Path(APPDATA) / ".runclap"
RUNTIME_DIR = APP_DIR / ".compile" / "runtime"
APP_PATH = RUNTIME_DIR / "app.py"
ICON_PNG_PATH = RUNTIME_DIR / "logo.png"
ICON_ICO_PATH = RUNTIME_DIR / "logo.ico"
FINAL_EXE_PATH = RUNTIME_DIR / "dist" / f"{FINAL_APP_NAME}.exe"

FEATURES = {
    "core": {
        "name": "Core Functionality (Required)",
        "deps": [
            "PyQt6", "pandas", "python-docx", "PyPDF2", "reportlab",
            "openpyxl", "camelot-py[cv]", "pyarrow", "qtawesome",
            "pyinstaller", "winshell", "Pillow"
        ],
        "checked": True, "disabled": True
    },
    "ai": { "name": "AI Features", "deps": ["google-generativeai", "openai", "requests", "anthropic", "groq"], "checked": True, "disabled": False },
    "gpu": { "name": "GPU Features (PyTorch)", "deps": ["torch", "torchvision", "torchaudio"], "checked": False, "disabled": False },
    "dev": { "name": "Developer Features", "deps": ["PyQt6-sip", "qtconsole", "ipykernel", "psutil", "pyqtgraph", "debugpy"], "checked": False, "disabled": False }
}

STYLESHEET = """
QWidget { font-family: 'Segoe UI', sans-serif; font-size: 11pt; background-color: #2E3440; color: #D8DEE9; }
QGroupBox { border: 1px solid #4C566A; border-radius: 5px; margin-top: 1ex; font-weight: bold; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 3px; }
QTextEdit { background-color: #3B4252; border: 1px solid #4C566A; border-radius: 5px; }
QProgressBar { border: 1px solid #4C566A; border-radius: 5px; text-align: center; color: #ECEFF4; }
QProgressBar::chunk { background-color: #88C0D0; border-radius: 4px; }
QProgressBar[state="success"]::chunk { background-color: #A3BE8C; }
QProgressBar[state="error"]::chunk { background-color: #BF616A; }
QProgressBar[state="cancelled"]::chunk { background-color: #EBCB8B; }
QPushButton { background-color: #5E81AC; color: #ECEFF4; border: none; padding: 8px 16px; border-radius: 5px; font-weight: bold; }
QPushButton:hover { background-color: #81A1C1; }
QPushButton:disabled { background-color: #4C566A; color: #D8DEE9; }
QPushButton#CancelButton { background-color: #BF616A; }
QPushButton#CancelButton:hover { background-color: #D08770; }
"""

# ================= HELPERS =================
class InstallationCancelled(Exception): pass

def check_package_installed(pkg):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "show", pkg], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError: return False

def handle_tensorflow_conflict():
    if sys.prefix != sys.base_prefix: return
    if check_package_installed('tensorflow'):
        print("!! Conflicting TensorFlow installation detected.")
        print("-> Creating a clean virtual environment to ensure a successful build.")
        venv_dir = Path('.venv')
        if not venv_dir.exists():
            subprocess.run([sys.executable, '-m', 'venv', str(venv_dir)], check=True)
        venv_python = venv_dir / 'Scripts' / 'python.exe' if os.name == 'nt' else venv_dir / 'bin' / 'python'
        print("-> Re-launching the installer within the clean environment...")
        subprocess.run([str(venv_python), __file__])
        sys.exit(0)

def get_pyinstaller_hidden_imports(selected_features):
    """Generates extra arguments for PyInstaller for tricky packages."""
    args = []
    # --- *** FIX IS HERE *** ---
    # If dev features are selected, we must manually find and add the debugpy data folder.
    if 'dev' in selected_features:
        try:
            import debugpy
            # Get the path to the _vendored folder inside the installed debugpy library
            source_path = Path(debugpy.__file__).parent / "_vendored"
            if source_path.exists():
                # The separator is ';' on Windows, ':' on Linux/macOS
                separator = os.pathsep
                destination_path = "debugpy/_vendored"
                args.extend(["--add-data", f"{source_path}{separator}{destination_path}"])
        except ImportError:
            print("Warning: Could not import 'debugpy' to find its data path.")
        except Exception as e:
            print(f"Warning: Failed to build debugpy data path: {e}")
    return args

def run_cancellable_command(worker_thread, cmd, cwd=None):
    startupinfo = None
    if os.name == 'nt':
        startupinfo = subprocess.STARTUPINFO(); startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    process = subprocess.Popen(cmd, startupinfo=startupinfo, cwd=cwd)
    while process.poll() is None:
        if worker_thread.isInterruptionRequested():
            process.terminate(); process.wait()
            raise InstallationCancelled("Installation cancelled by user.")
        QThread.msleep(100)
    if process.returncode != 0: raise subprocess.CalledProcessError(process.returncode, cmd)

# ================= WORKER THREADS =================
class InstallWorker(QThread):
    log = pyqtSignal(str); progress = pyqtSignal(int); done = pyqtSignal(bool, str)
    def __init__(self, packages, selected_features, content):
        super().__init__()
        self.packages = packages
        self.selected_features = selected_features
        self.online_content = content

    def run(self):
        try:
            total_deps = len(self.packages)
            for i, pkg in enumerate(self.packages, 1):
                if self.isInterruptionRequested(): raise InstallationCancelled()
                self.log.emit(f"Checking for {pkg}...")
                if not check_package_installed(pkg):
                    self.log.emit(f"-> Installing {pkg}...")
                    run_cancellable_command(self, [sys.executable, "-m", "pip", "install", "--upgrade", pkg])
                self.progress.emit(int(10 + (i / total_deps) * 35))
            
            self.log.emit("Downloading application source..."); RUNTIME_DIR.mkdir(parents=True, exist_ok=True); APP_PATH.write_bytes(self.online_content); self.progress.emit(50)

            icon_arg = []
            try:
                self.log.emit("Downloading application icon...")
                urllib.request.urlretrieve(ICON_URL, ICON_PNG_PATH)
                from PIL import Image
                self.log.emit("Converting icon to .ico format...")
                img = Image.open(ICON_PNG_PATH)
                img.save(ICON_ICO_PATH, format='ICO', sizes=[(256, 256)])
                icon_arg = ["--icon", str(ICON_ICO_PATH)]
                self.progress.emit(60)
            except Exception as e:
                self.log.emit(f"-> Warning: Could not process icon: {e}")

            self.log.emit("Packaging application (this may take a while)...")
            extra_args = get_pyinstaller_hidden_imports(self.selected_features)
            pyinstaller_cmd = [sys.executable, "-m", "PyInstaller", "--onefile", "--windowed", "--name", FINAL_APP_NAME] + icon_arg + extra_args + [str(APP_PATH)]
            run_cancellable_command(self, pyinstaller_cmd, cwd=RUNTIME_DIR); self.progress.emit(90)
            
            self.log.emit("Creating shortcuts...")
            import winshell
            desktop = winshell.desktop()
            start_menu = winshell.programs()
            link_filepath_desktop = os.path.join(desktop, f"{FINAL_APP_NAME}.lnk")
            with winshell.shortcut(link_filepath_desktop) as link:
                link.path = str(FINAL_EXE_PATH); link.description = f"Launch {FINAL_APP_NAME}"; link.working_directory = str(FINAL_EXE_PATH.parent); link.icon_location = (str(FINAL_EXE_PATH), 0)
            
            start_menu_folder = os.path.join(start_menu, FINAL_APP_NAME)
            os.makedirs(start_menu_folder, exist_ok=True)
            link_filepath_start = os.path.join(start_menu_folder, f"{FINAL_APP_NAME}.lnk")
            with winshell.shortcut(link_filepath_start) as link:
                link.path = str(FINAL_EXE_PATH); link.description = f"Launch {FINAL_APP_NAME}"; link.working_directory = str(FINAL_EXE_PATH.parent); link.icon_location = (str(FINAL_EXE_PATH), 0)

            self.progress.emit(100); self.log.emit("✅ Installation complete!")
            self.done.emit(True, "install_ok")
        except InstallationCancelled: self.log.emit("🟡 Installation cancelled."); self.done.emit(False, "cancelled")
        except Exception as e: self.log.emit(f"❌ Error: {e}"); self.done.emit(False, "error")

class UninstallWorker(QThread):
    log = pyqtSignal(str); progress = pyqtSignal(int); done = pyqtSignal(bool, str)
    def run(self):
        try:
            self.log.emit("Removing shortcuts..."); self.progress.emit(25)
            import winshell
            desktop = winshell.desktop(); start_menu = winshell.programs()
            desktop_link = Path(desktop) / f"{FINAL_APP_NAME}.lnk"
            if desktop_link.exists(): desktop_link.unlink()
            start_menu_folder = Path(start_menu) / FINAL_APP_NAME
            if start_menu_folder.exists(): shutil.rmtree(start_menu_folder)
            self.log.emit("Removing application files..."); self.progress.emit(75)
            if APP_DIR.exists(): shutil.rmtree(APP_DIR)
            self.progress.emit(100); self.log.emit("✅ Uninstallation complete.")
            self.done.emit(True, "uninstall_ok")
        except Exception as e: self.log.emit(f"❌ Error: {e}"); self.done.emit(False, "error")

class Installer(QWidget):
    def __init__(self):
        super().__init__(); self.setWindowTitle(f"{FINAL_APP_NAME} Installer"); self.setFixedSize(550, 450)
        self.stack = QStackedWidget(); self.main_layout = QVBoxLayout(self); self.main_layout.addWidget(self.stack)
        try:
            with urllib.request.urlopen(APP_URL) as r: self.online_content = r.read()
        except Exception as e: self.online_content = None; print(f"Could not fetch online app: {e}")
        self.create_setup_page(); self.create_progress_page(); self.create_maintenance_page()
        self.stack.setCurrentWidget(self.maintenance_page if FINAL_EXE_PATH.exists() else self.setup_page)

    def create_setup_page(self):
        self.setup_page = QWidget(); layout = QVBoxLayout(self.setup_page)
        title = QLabel(f"Welcome to {FINAL_APP_NAME} Setup"); title.setStyleSheet("font-size: 14pt; font-weight: bold;")
        group = QGroupBox("Select Optional Features"); group_layout = QVBoxLayout(group)
        self.checkboxes = {}
        for k, f in {k:v for k,v in FEATURES.items() if not v['disabled']}.items():
            cb = QCheckBox(f['name']); cb.setChecked(f["checked"]); group_layout.addWidget(cb); self.checkboxes[k] = cb
        layout.addWidget(title); layout.addWidget(group); layout.addStretch()
        self.install_btn = QPushButton("Install"); self.install_btn.clicked.connect(self.start_installation)
        if not self.online_content: self.install_btn.setDisabled(True); self.install_btn.setText("Cannot connect to update server")
        layout.addWidget(self.install_btn); self.stack.addWidget(self.setup_page)

    def create_maintenance_page(self):
        self.maintenance_page = QWidget(); layout = QVBoxLayout(self.maintenance_page)
        title = QLabel(f"{FINAL_APP_NAME} is already installed."); title.setStyleSheet("font-size: 14pt; font-weight: bold;")
        info = QLabel("You can uninstall the application using the button below.")
        self.uninstall_btn = QPushButton("Uninstall"); self.uninstall_btn.clicked.connect(self.start_uninstallation)
        layout.addWidget(title); layout.addWidget(info); layout.addStretch(); layout.addWidget(self.uninstall_btn); self.stack.addWidget(self.maintenance_page)

    def create_progress_page(self):
        self.progress_page = QWidget(); layout = QVBoxLayout(self.progress_page)
        self.progress_label = QLabel("Initializing…"); self.progress_bar = QProgressBar()
        details_log = QTextEdit(); details_log.setReadOnly(True)
        self.details_group = QGroupBox("Details"); box_layout = QVBoxLayout(); box_layout.addWidget(details_log); self.details_group.setLayout(box_layout)
        self.close_btn = QPushButton("Close"); self.close_btn.setEnabled(False); self.close_btn.clicked.connect(self.close)
        self.cancel_btn = QPushButton("Cancel"); self.cancel_btn.setObjectName("CancelButton"); self.cancel_btn.clicked.connect(self.cancel_operation)
        btn_layout = QHBoxLayout(); btn_layout.addStretch(); btn_layout.addWidget(self.cancel_btn); btn_layout.addWidget(self.close_btn)
        layout.addWidget(self.progress_label); layout.addWidget(self.progress_bar); layout.addWidget(self.details_group); layout.addLayout(btn_layout)
        self.stack.addWidget(self.progress_page); self.worker_log_target = details_log

    def start_installation(self):
        self.stack.setCurrentWidget(self.progress_page); self.close_btn.hide(); self.cancel_btn.show()
        selected = ["core"] + [k for k, cb in self.checkboxes.items() if cb.isChecked()]
        packages = list(dict.fromkeys([d for f in selected for d in FEATURES[f]["deps"]]))
        self.worker = InstallWorker(packages, selected, self.online_content)
        self.worker.log.connect(self.log_to_details); self.worker.progress.connect(self.progress_bar.setValue); self.worker.done.connect(self.on_done); self.worker.start()

    def start_uninstallation(self):
        self.stack.setCurrentWidget(self.progress_page); self.close_btn.hide(); self.cancel_btn.hide()
        self.worker = UninstallWorker(); self.worker.log.connect(self.log_to_details); self.worker.progress.connect(self.progress_bar.setValue); self.worker.done.connect(self.on_done); self.worker.start()

    def cancel_operation(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.cancel_btn.setEnabled(False); self.cancel_btn.setText("Cancelling...")
            self.worker.requestInterruption()

    def log_to_details(self, msg):
        self.progress_label.setText(msg); self.worker_log_target.append(msg)
        self.worker_log_target.verticalScrollBar().setValue(self.worker_log_target.verticalScrollBar().maximum())

    def on_done(self, ok, status):
        self.cancel_btn.hide(); self.close_btn.show(); self.close_btn.setEnabled(True)
        status_map = {
            "install_ok": ("✅ Installation complete! You can close this window.", "success"),
            "uninstall_ok": ("✅ Uninstallation complete.", "success"),
            "cancelled": ("🟡 Operation was cancelled by the user.", "cancelled"),
            "error": ("❌ Operation failed. See details for the error.", "error")
        }
        message, state = status_map.get(status, ("An unknown event occurred.", "error"))
        self.progress_label.setText(message); self.progress_bar.setProperty("state", state)
        self.progress_bar.style().unpolish(self.progress_bar); self.progress_bar.style().polish(self.progress_bar)

if __name__ == "__main__":
    handle_tensorflow_conflict()
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    w = Installer()
    w.show()
    sys.exit(app.exec())