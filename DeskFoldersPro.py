import sys
import os
import json
import winreg
import ctypes
import re
import subprocess
from ctypes import wintypes
from pathlib import Path
from PyQt6.QtCore import Qt, QPoint, QSize, QEvent, QPropertyAnimation, QEasingCurve, QRect, QMimeData, QFileInfo
from PyQt6.QtGui import QIcon, QAction, QCursor, QImage, QPixmap, QDrag
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QGridLayout, QLabel,
    QPushButton, QToolButton, QMenu, QSystemTrayIcon, QFileDialog, QInputDialog,
    QMessageBox, QSizeGrip, QHBoxLayout, QTabWidget, QFormLayout, QCheckBox, QDialog, 
    QListWidget, QListWidgetItem, QStyle, QSlider, QComboBox, QFileIconProvider
)

APP_NAME = "DeskFoldersPro"
# --- كود إصلاح أيقونة شريط المهام (Taskbar) ---
myappid = f'mycompany.myproduct.{APP_NAME}.1'
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass
# ---------------------------------------------
CONFIG_DIR = Path.home() / "AppData" / "Roaming" / APP_NAME
CONFIG_FILE = CONFIG_DIR / "config.json"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)

# ----------------- نظام الترجمة الفوري -----------------
class I18N:
    LANG = "ar"
    STRINGS = {
        "ar": {
            "Dashboard": "لوحة التحكم",
            "Folders": "المجلدات",
            "Preferences": "الإعدادات",
            "Default Folders": "المجلدات الافتراضية",
            "Enable Grid Snapping": "محاذاة المجلدات مع الشبكة (Grid Snapping)",
            "Run App on Windows Startup": "تشغيل التطبيق عند بدء ويندوز",
            "Inner Opacity:": "شفافية المجلد من الداخل:",
            "Outer Opacity:": "شفافية المجلد من الخارج:",
            "Language:": "لغة التطبيق (Language):",
            "Welcome to DeskFoldersPro": "مرحباً بك في DeskFoldersPro",
            "Select default folders to create:": "حدد المجلدات الافتراضية التي ترغب بإنشائها:",
            "Start": "ابدأ",
            "Add": "إضافة",
            "Hide/Delete": "إخفاء/حذف",
            "Factory Reset (Reset All Folders)": "إعادة ضبط المصنع (حذف كل شيء)",
            "+ Create New Folder": "+ إنشاء مجلد جديد",
            "New Folder": "مجلد جديد",
            "Recent": "أحدث البرامج",
            "Games": "الألعاب",
            "Launchers": "منصات الألعاب",
            "Developer Tools": "أدوات المطورين",
            "Design Apps": "تطبيقات التصميم",
            "Office Apps": "تطبيقات أوفيس",
            "Windows Tools": "أدوات ويندوز",
            "File Explorer": "مستكشف الملفات",
            "Paint": "الرسام",
            "System Information": "معلومات النظام",
            "Device Manager": "إدارة الأجهزة",
            "Disk Management": "إدارة الأقراص",
            "Command Prompt": "موجه الأوامر (CMD)",
            "Task Manager": "إدارة المهام",
            "Control Panel": "لوحة التحكم",
            "Registry Editor": "محرر التسجيل",
            "Notepad": "المفكرة",
            "Snipping Tool": "أداة القص",
            "Calculator": "الآلة الحاسبة",
            "PowerShell": "PowerShell",
            "Remote Desktop": "سطح المكتب البعيد",
            "Exit": "إنهاء البرنامج",
            "Add App Manually...": "إضافة ملف/تطبيق يدوياً...",
            "Rename Folder": "إعادة تسمية المجلد",
            "Change Opacity": "تغيير الشفافية",
            "Inner Opacity (10-100):": "الشفافية الداخلية (10-100):",
            "Outer Opacity (0-100):": "الشفافية الخارجية (0-100):",
            "Delete Folder": "حذف المجلد",
            "Remove App": "إزالة العنصر من المجلد",
            "New Name:": "الاسم الجديد:",
            "Rename": "إعادة تسمية",
            "Select App": "اختر تطبيقاً، ملفاً، أو مجلداً",
            "Factory Reset": "إعادة ضبط المصنع",
            "Are you sure you want to reset all folders?": "هل أنت متأكد من رغبتك في إعادة ضبط المصنع؟\nسيتم حذف جميع المجلدات الحالية!"
        }
    }
    
    @classmethod
    def tr(cls, text):
        if cls.LANG == "ar": return cls.STRINGS["ar"].get(text, text)
        return text

    @classmethod
    def reverse_tr(cls, localized_text):
        for k, v in cls.STRINGS["ar"].items():
            if v == localized_text: return k
        return localized_text

# ----------------- الوظائف الأساسية لمعالجة الملفات والأيقونات -----------------
class SHFILEINFO(ctypes.Structure):
    _fields_ = [("hIcon", wintypes.HICON), ("iIcon", ctypes.c_int), ("dwAttributes", wintypes.DWORD), ("szDisplayName", wintypes.WCHAR * 260), ("szTypeName", wintypes.WCHAR * 80)]

def extract_exe_and_args(cmd_str):
    cmd_str = str(cmd_str).strip()
    if os.path.exists(cmd_str): return cmd_str, ""
    if cmd_str.startswith('"'):
        end_quote = cmd_str.find('"', 1)
        if end_quote != -1: return cmd_str[1:end_quote], cmd_str[end_quote+1:].strip()
    parts = cmd_str.split(' ')
    for i in range(len(parts), 0, -1):
        candidate = ' '.join(parts[:i])
        if os.path.exists(candidate): return candidate, ' '.join(parts[i:])
    return cmd_str, ""

def resolve_shortcut(lnk_path):
    try:
        from win32com.client import Dispatch
        shell = Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(str(lnk_path))
        return {"target": sc.Targetpath or str(lnk_path), "args": sc.Arguments or "", "icon_location": sc.IconLocation or "", "workdir": sc.WorkingDirectory or ""}
    except Exception: return {"target": str(lnk_path), "args": "", "icon_location": "", "workdir": ""}

def get_real_path(path_str):
    p = str(path_str).strip()
    if p.lower().endswith(".lnk"):
        info = resolve_shortcut(p)
        if info.get("target") and os.path.exists(info["target"]):
            clean, _ = extract_exe_and_args(info["target"])
            return clean.lower()
    clean, _ = extract_exe_and_args(p)
    return clean.lower()

def get_native_icon(file_path, icon_location=""):
    path_str = str(file_path)
    if path_str.lower().endswith(".lnk"):
        info = resolve_shortcut(path_str)
        if info.get("target") and os.path.exists(info["target"]):
            path_str = info["target"]
            icon_location = icon_location or info.get("icon_location", "")

    clean_path, _ = extract_exe_and_args(path_str)
    lower_path = clean_path.lower()
    
    if lower_path.endswith((".png", ".jpg", ".jpeg", ".ico", ".bmp", ".svg")):
        if os.path.exists(clean_path): return QIcon(clean_path)

    if icon_location:
        try:
            icon_file, _, icon_idx_str = icon_location.rpartition(",")
            icon_idx = int(icon_idx_str) if icon_idx_str.strip().lstrip("-").isdigit() else 0
            if icon_file and os.path.exists(icon_file):
                large, small = wintypes.HICON(), wintypes.HICON()
                if ctypes.windll.shell32.ExtractIconExW(icon_file, icon_idx, ctypes.byref(large), ctypes.byref(small), 1) and large.value:
                    image = QImage.fromHICON(large.value)
                    ctypes.windll.user32.DestroyIcon(large.value)
                    if small.value: ctypes.windll.user32.DestroyIcon(small.value)
                    if not image.isNull(): return QIcon(QPixmap.fromImage(image))
        except Exception: pass

    if clean_path.lower().endswith(".url"):
        try:
            with open(clean_path, "r", encoding="utf-8", errors="ignore") as f:
                match = re.search(r"IconFile=(.+)", f.read(), re.IGNORECASE)
                if match and os.path.exists(match.group(1).strip()): return QIcon(match.group(1).strip())
        except Exception: pass

    try:
        shfi = SHFILEINFO()
        if ctypes.windll.shell32.SHGetFileInfoW(clean_path, 0, ctypes.byref(shfi), ctypes.sizeof(shfi), 0x00000100 | 0x00000000) and shfi.hIcon:
            image = QImage.fromHICON(shfi.hIcon)
            ctypes.windll.user32.DestroyIcon(shfi.hIcon)
            if not image.isNull(): return QIcon(QPixmap.fromImage(image))
    except Exception: pass

    if os.path.exists(clean_path):
        provider = QFileIconProvider()
        icon = provider.icon(QFileInfo(clean_path))
        if not icon.isNull(): return icon

    return QApplication.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)

class ConfigManager:
    @staticmethod
    def load_config():
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
            except Exception: pass
        return {"folders": [], "setup_done": False, "lang": "ar", "snap_to_grid": True}
    @staticmethod
    def save_config(data):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(data, f, indent=4, ensure_ascii=False)

# ----------------- عناصر الواجهة -----------------
class AppItem(QWidget):
    def __init__(self, name, path, parent_folder, args="", icon_location="", index=-1):
        super().__init__()
        self.name, self.path, self.args, self.icon_location, self.parent_folder, self.index = name, path, args, icon_location, parent_folder, index
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.icon_btn = QToolButton()
        self.icon_btn.setAutoRaise(True)
        self.icon_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.icon_btn.setIcon(get_native_icon(path, icon_location))
        self.icon_btn.setStyleSheet("QToolButton { background: transparent; border: none; border-radius: 12px; } QToolButton:hover { background: rgba(255, 255, 255, 0.12); }")
        self.icon_btn.clicked.connect(self.launch_app)
        self.icon_btn.installEventFilter(self)

        self.name_label = QLabel(I18N.tr(self.name))
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.name_label.setWordWrap(True)
        self.name_label.setMinimumHeight(35)
        self.name_label.setStyleSheet("color: #E0E0E0; font-family: 'Segoe UI'; font-size: 11.5px; font-weight: 600; border: none; background: transparent;")
        
        layout.addWidget(self.icon_btn, 0, Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self.name_label, 0, Qt.AlignmentFlag.AlignHCenter)

    def eventFilter(self, source, event):
        if source == self.icon_btn:
            if event.type() == QEvent.Type.MouseButtonPress and event.button() == Qt.MouseButton.LeftButton:
                self.drag_start_pos = event.pos()
            elif event.type() == QEvent.Type.MouseMove and (event.buttons() & Qt.MouseButton.LeftButton):
                if (event.pos() - self.drag_start_pos).manhattanLength() > QApplication.startDragDistance():
                    drag, mime = QDrag(self), QMimeData()
                    mime.setText(str(self.index))
                    drag.setMimeData(mime)
                    drag.exec(Qt.DropAction.MoveAction)
                    return True
        return super().eventFilter(source, event)

    def update_scale(self, icon_size, text_width):
        self.setFixedWidth(text_width + 16)
        self.icon_btn.setFixedSize(icon_size + 16, icon_size + 16)
        self.icon_btn.setIconSize(QSize(icon_size, icon_size))
        self.name_label.setFixedWidth(text_width + 16)
        self.name_label.setText(I18N.tr(self.name))

    def launch_app(self):
        try:
            if self.args: subprocess.Popen(f'"{self.path}" {self.args}', shell=True)
            elif os.path.exists(self.path): os.startfile(self.path)
            else: subprocess.Popen(self.path, shell=True)
            self.parent_folder.switch_mode(False)
        except Exception:
            try:
                subprocess.Popen(f'start "" "{self.path}" {self.args}', shell=True)
                self.parent_folder.switch_mode(False)
            except Exception: pass

    def contextMenuEvent(self, event):
        self.parent_folder.ignore_deactivate = True
        menu = QMenu(self)
        menu.setStyleSheet("QMenu { background-color: #2b2b2b; color: white; border-radius: 6px; } QMenu::item:selected { background-color: #d32f2f; }")
        action = QAction(I18N.tr("Remove App"), self)
        action.triggered.connect(lambda: self.parent_folder.remove_app(self.path))
        menu.addAction(action)
        menu.exec(QCursor.pos())
        self.parent_folder.ignore_deactivate = False


class ClickableTitleLabel(QLabel):
    def __init__(self, text, rename_callback):
        super().__init__(text)
        self.rename_callback = rename_callback
    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton: self.rename_callback()


class StartFolderWidget(QWidget):
    def __init__(self, folder_data, controller):
        super().__init__()
        self.controller = controller
        self.folder_id = folder_data.get("id")
        
        raw_title = str(folder_data.get("title", I18N.tr("New Folder")))
        self.title = I18N.reverse_tr(raw_title)
        
        self.apps = folder_data.get("apps", [])
        self.opacity_inner = folder_data.get("opacity_inner", 100)
        self.opacity_outer = folder_data.get("opacity_outer", 10)
        
        self.expanded_w, self.expanded_h = folder_data.get("width", 320), folder_data.get("height", 360)
        self.is_expanded = self.ignore_deactivate = self.is_dragging = False

        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.NoDropShadowWindowHint | Qt.WindowType.WindowStaysOnBottomHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAcceptDrops(True)
        self.setStyleSheet("StartFolderWidget, QWidget { border: none; outline: none; background: transparent; }")

        if "x" in folder_data and "y" in folder_data: self.move(folder_data["x"], folder_data["y"])
        
        self.app_items = []
        self.init_ui()
        self.switch_mode(expanded=False)

    def init_ui(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.preview_container = QWidget()
        preview_layout = QVBoxLayout(self.preview_container)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(6)
        preview_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.tile_box = QWidget()
        self.tile_box.setFixedSize(70, 70)
        self.tile_box.setCursor(Qt.CursorShape.PointingHandCursor)
        self.tile_grid = QGridLayout(self.tile_box)
        self.tile_grid.setContentsMargins(8, 8, 8, 8)
        self.tile_grid.setSpacing(4)

        self.tile_title = ClickableTitleLabel(I18N.tr(self.title), self.rename_folder)
        self.tile_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tile_title.setFixedWidth(95)
        self.tile_title.setStyleSheet("color: white; font-size: 12px; font-weight: bold; font-family: 'Segoe UI'; text-align: center; text-shadow: 1px 1px 2px black;")

        preview_layout.addWidget(self.tile_box, 0, Qt.AlignmentFlag.AlignHCenter)
        preview_layout.addWidget(self.tile_title, 0, Qt.AlignmentFlag.AlignHCenter)

        self.expanded_container = QWidget()
        self.expanded_container.setObjectName("ExpandedContainer")
        self.expanded_layout = QVBoxLayout(self.expanded_container)
        
        header_layout = QHBoxLayout()
        self.expanded_title = ClickableTitleLabel(I18N.tr(self.title), self.rename_folder)
        self.expanded_title.setStyleSheet("color: white; font-weight: bold; font-size: 14px;")
        header_layout.addWidget(self.expanded_title)
        header_layout.addStretch()
        self.expanded_layout.addLayout(header_layout)

        self.full_grid_widget = QWidget()
        self.full_grid = QGridLayout(self.full_grid_widget)
        self.full_grid.setSpacing(10)
        self.full_grid.setAlignment(Qt.AlignmentFlag.AlignTop | (Qt.AlignmentFlag.AlignRight if I18N.LANG == "ar" else Qt.AlignmentFlag.AlignLeft))
        self.expanded_layout.addWidget(self.full_grid_widget)
        self.expanded_layout.addStretch()

        self.size_grip = QSizeGrip(self.expanded_container)
        self.expanded_layout.addWidget(self.size_grip, 0, Qt.AlignmentFlag.AlignBottom | (Qt.AlignmentFlag.AlignLeft if I18N.LANG == "ar" else Qt.AlignmentFlag.AlignRight))

        self.main_layout.addWidget(self.preview_container)
        self.main_layout.addWidget(self.expanded_container)

        self.tile_box.mousePressEvent = lambda e: self.tile_mouse_evt(e, 'press')
        self.tile_box.mouseMoveEvent = lambda e: self.tile_mouse_evt(e, 'move')
        self.tile_box.mouseReleaseEvent = lambda e: self.tile_mouse_evt(e, 'release')

        self.apply_style()
        self.refresh_apps()

    def tile_mouse_evt(self, event, action):
        if action == 'press' and event.button() == Qt.MouseButton.LeftButton:
            self.old_pos = event.globalPosition().toPoint()
            self.is_dragging = False
        elif action == 'move' and event.buttons() == Qt.MouseButton.LeftButton:
            delta = QPoint(event.globalPosition().toPoint() - self.old_pos)
            if delta.manhattanLength() > 3:
                self.is_dragging = True
                self.move(self.x() + delta.x(), self.y() + delta.y())
                self.old_pos = event.globalPosition().toPoint()
        elif action == 'release' and event.button() == Qt.MouseButton.LeftButton:
            if self.is_dragging:
                new_x, new_y = self.x(), self.y()
                # شبكة أوسع 110 بكسل
                if getattr(self.controller, 'snap_to_grid', True):
                    new_x = round(new_x / 110) * 110
                    new_y = max(0, round(new_y / 110) * 110) # الحفاظ على الالتصاق بالسقف
                
                # نظام تصادم ذكي 110 بكسل
                collision = True
                while collision:
                    collision = False
                    for f in self.controller.folders:
                        if f != self and not f.is_expanded:
                            if abs(f.x() - new_x) < 55 and abs(f.y() - new_y) < 55:
                                new_x += 110
                                collision = True
                                break
                                
                self.move(new_x, new_y)
                self.controller.save_all()
                self.is_dragging = False
            else:
                self.switch_mode(True)

    def switch_mode(self, expanded):
        self.is_expanded = expanded
        if not expanded:
            self.setWindowFlag(Qt.WindowType.WindowStaysOnBottomHint, True)
            if hasattr(self, 'saved_tile_pos'): self.move(self.saved_tile_pos)
            self.expanded_container.hide()
            self.preview_container.show()
            self.setFixedSize(110, 130)
            self.show()
        else:
            for folder in self.controller.folders:
                if folder != self and folder.is_expanded: folder.switch_mode(False)
            self.saved_tile_pos = self.pos()
            new_w, new_h = max(350, self.expanded_w), max(400, self.expanded_h)
            self.setMinimumSize(350, 400)
            
            screen_geom = QApplication.primaryScreen().availableGeometry()
            new_x = self.x() - ((new_w - 110) // 2)
            new_y = self.y() + 90
            new_x = max(screen_geom.left() + 10, min(new_x, screen_geom.right() - new_w - 10))
            new_y = max(screen_geom.top() + 10, min(new_y, screen_geom.bottom() - new_h - 10))
            
            self.setWindowFlag(Qt.WindowType.WindowStaysOnBottomHint, False) 
            self.preview_container.hide()
            self.expanded_container.show()
            self.setGeometry(new_x, new_y, new_w, new_h)
            self.show()
            self.refresh_apps()
            self.raise_()
            self.activateWindow()

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.Type.ActivationChange and not self.isActiveWindow() and self.is_expanded and not self.ignore_deactivate:
            self.switch_mode(False)

    def apply_style(self):
        alpha_inner = int((self.opacity_inner / 100) * 255)
        alpha_outer = int((self.opacity_outer / 100) * 255)
        self.expanded_container.setStyleSheet(f"#ExpandedContainer {{ background-color: rgba(20, 20, 20, {alpha_inner}); border: 1px solid rgba(255, 255, 255, 0.15); border-radius: 16px; }}")
        self.tile_box.setStyleSheet(f"background-color: rgba(255, 255, 255, {alpha_outer}); border-radius: 14px;")

    def calculate_dynamic_scale(self):
        cols = max(3, self.width() // 90)
        cell_w = max(70, (self.width() - 40) // cols)
        return cols, max(36, min(64, int(cell_w * 0.55))), max(65, cell_w - 8)

    def refresh_apps(self):
        valid_apps = [app for app in self.apps if app["path"].lower().startswith("steam://") or os.path.exists(app["path"])]
        if len(valid_apps) != len(self.apps):
            self.apps = valid_apps
            self.controller.save_all()

        for i in reversed(range(self.tile_grid.count())): 
            w = self.tile_grid.itemAt(i).widget()
            if w: w.deleteLater()
            
        for idx, app in enumerate(self.apps[:4]):
            icon_lbl = QLabel()
            icon_lbl.setPixmap(get_native_icon(app["path"], app.get("icon_location", "")).pixmap(22, 22))
            icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            icon_lbl.setStyleSheet("background: transparent; border: none;")
            self.tile_grid.addWidget(icon_lbl, *divmod(idx, 2))

        for i in reversed(range(self.full_grid.count())):
            w = self.full_grid.itemAt(i).widget()
            if w: w.deleteLater()

        self.app_items.clear()
        cols, icon_size, text_w = self.calculate_dynamic_scale()

        for idx, app in enumerate(self.apps):
            item = AppItem(app["name"], app["path"], self, app.get("args", ""), app.get("icon_location", ""), idx)
            item.update_scale(icon_size, text_w)
            self.full_grid.addWidget(item, idx // cols, idx % cols)
            self.app_items.append(item)

    def add_app(self, path):
        path_str = str(path).strip()
        args, icon_location = "", ""
        if path_str.lower().endswith(".lnk"):
            info = resolve_shortcut(path_str)
            target, args, icon_location = info["target"] or path_str, info["args"], info["icon_location"]
        else: target, args = extract_exe_and_args(path_str)

        if not os.path.exists(target) and not path_str.lower().startswith("steam://"): return
        if any(app["path"] == target for app in self.apps): return
        
        self.apps.append({"name": Path(target).stem if target else "App", "path": target, "args": args, "icon_location": icon_location})
        self.refresh_apps()
        self.controller.save_all()

    def remove_app(self, path):
        self.apps = [app for app in self.apps if app["path"] != path]
        self.refresh_apps()
        self.controller.save_all()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() or event.mimeData().hasText(): event.acceptProposedAction()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile(): self.add_app(url.toLocalFile())
        elif event.mimeData().hasText():
            try:
                src_idx = int(event.mimeData().text())
                drop_pos = event.position().toPoint()
                dest_idx = len(self.apps) - 1
                for item in self.app_items:
                    if item.geometry().contains(self.full_grid_widget.mapFrom(self, drop_pos)):
                        dest_idx = item.index
                        break
                if src_idx != dest_idx:
                    self.apps.insert(dest_idx, self.apps.pop(src_idx))
                    self.refresh_apps()
                    self.controller.save_all()
            except ValueError: pass

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.is_expanded:
            self.expanded_w, self.expanded_h = self.width(), self.height()
            cols, icon_size, text_w = self.calculate_dynamic_scale()
            for item in self.app_items: item.update_scale(icon_size, text_w)

    def contextMenuEvent(self, event):
        self.ignore_deactivate = True
        menu = QMenu(self)
        menu.setStyleSheet("QMenu { background-color: #1e1e1e; color: #e0e0e0; border-radius: 8px; font-family: 'Segoe UI'; font-size: 12px; padding: 4px; } QMenu::item:selected { background-color: #0067b8; }")
        
        for name, func in [(I18N.tr("Add App Manually..."), self.manual_add_app), (I18N.tr("Rename Folder"), self.rename_folder), (I18N.tr("Change Opacity"), self.change_opacity)]:
            action = QAction(name, self)
            action.triggered.connect(func)
            menu.addAction(action)
            
        menu.addSeparator()
        del_action = QAction(I18N.tr("Delete Folder"), self)
        del_action.triggered.connect(lambda: self.controller.delete_folder(self.folder_id))
        menu.addAction(del_action)
        menu.exec(QCursor.pos())
        self.ignore_deactivate = False

    def manual_add_app(self):
        self.ignore_deactivate = True
        path, _ = QFileDialog.getOpenFileName(None, I18N.tr("Select App"), "", "All Supported (*.*)")
        if path: self.add_app(path)
        self.ignore_deactivate = False

    def rename_folder(self):
        self.ignore_deactivate = True
        new_title, ok = QInputDialog.getText(self, I18N.tr("Rename"), I18N.tr("New Name:"), text=I18N.tr(self.title))
        if ok and new_title.strip():
            self.title = I18N.reverse_tr(new_title.strip())
            self.tile_title.setText(I18N.tr(self.title))
            self.expanded_title.setText(I18N.tr(self.title))
            self.controller.save_all()
        self.ignore_deactivate = False

    def change_opacity(self):
        self.ignore_deactivate = True
        inner, ok1 = QInputDialog.getInt(None, I18N.tr("Change Opacity"), I18N.tr("Inner Opacity (10-100):"), self.opacity_inner, 10, 100, 5)
        if ok1:
            outer, ok2 = QInputDialog.getInt(None, I18N.tr("Change Opacity"), I18N.tr("Outer Opacity (0-100):"), self.opacity_outer, 0, 100, 5)
            if ok2:
                self.opacity_inner = inner
                self.opacity_outer = outer
                self.apply_style()
                self.controller.save_all()
        self.ignore_deactivate = False

    def get_data(self):
        return {"id": self.folder_id, "title": self.title, "x": self.x(), "y": self.y(), "width": self.expanded_w, "height": self.expanded_h, "opacity_inner": self.opacity_inner, "opacity_outer": self.opacity_outer, "apps": self.apps}

# ----------------- شاشة الترحيب (First Run Wizard) -----------------
class FirstRunWizard(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(I18N.tr("Welcome to DeskFoldersPro"))
        self.resize(350, 380)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet("QDialog { background-color: #202020; color: white; } QCheckBox { font-size: 14px; font-weight: bold; margin: 5px; color: #e0e0e0; } QPushButton { background: #0067b8; color: white; padding: 12px; border-radius: 6px; font-weight: bold; font-size: 14px; } QPushButton:hover { background: #0078d4; }")
        
        layout = QVBoxLayout(self)
        lbl = QLabel(I18N.tr("Select default folders to create:"))
        lbl.setStyleSheet("font-size: 15px; font-weight: bold; margin-bottom: 15px; color: #fff;")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)
        
        self.checkboxes = {}
        defaults = ["Recent", "Games", "Launchers", "Developer Tools", "Design Apps", "Office Apps", "Windows Tools"]
        for d in defaults:
            chk = QCheckBox(I18N.tr(d))
            chk.setChecked(True)
            self.checkboxes[d] = chk
            layout.addWidget(chk)
            
        btn = QPushButton(I18N.tr("Start"))
        btn.clicked.connect(self.accept)
        layout.addStretch()
        layout.addWidget(btn)
        
    def get_selected(self):
        return [k for k, v in self.checkboxes.items() if v.isChecked()]


# ----------------- شاشة الإعدادات المتقدمة -----------------
class SettingsWindow(QDialog):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.setWindowTitle(I18N.tr("Dashboard"))
        self.resize(580, 500)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet("""
            QDialog { background-color: #202020; color: #f0f0f0; }
            QTabWidget::pane { border: 1px solid #333; background: #252525; border-radius: 8px; }
            QTabBar::tab { background: #1e1e1e; padding: 12px 30px; color: #aaa; font-weight: bold; border-top-left-radius: 8px; border-top-right-radius: 8px; }
            QTabBar::tab:selected { background: #252525; color: #4facfe; }
            QPushButton { background: #0067b8; color: white; border-radius: 6px; padding: 10px; font-weight: bold; }
            QPushButton:hover { background: #0078d4; }
            QListWidget { background-color: #1e1e1e; border: 1px solid #333; border-radius: 8px; color: white; font-size: 14px; padding: 8px; }
            QListWidget::item { padding: 12px; border-bottom: 1px solid #2d2d2d; }
            QListWidget::item:selected { background-color: #0067b8; border-radius: 6px; }
            QCheckBox, QLabel { color: white; font-size: 14px; font-weight: bold; }
            QComboBox { background-color: #1e1e1e; color: white; border: 1px solid #333; padding: 5px; border-radius: 4px; font-size: 14px; }
        """)
        
        self.main_layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)
        
        self.tab_folders = QWidget()
        folders_layout = QVBoxLayout(self.tab_folders)
        self.folders_list = QListWidget()
        folders_layout.addWidget(self.folders_list)
        add_btn = QPushButton(I18N.tr("+ Create New Folder"))
        add_btn.clicked.connect(self.add_new)
        folders_layout.addWidget(add_btn)
        self.tabs.addTab(self.tab_folders, I18N.tr("Folders"))

        # قسم إدارة الملفات الافتراضية الذكي
        self.tab_defaults = QWidget()
        defaults_layout = QVBoxLayout(self.tab_defaults)
        self.defaults_list = QListWidget()
        cats = ["Recent", "Games", "Launchers", "Developer Tools", "Design Apps", "Office Apps", "Windows Tools"]
        
        for cat in cats:
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(5,2,5,2)
            lbl = QLabel(I18N.tr(cat))
            row_layout.addWidget(lbl)
            
            btn = QPushButton()
            btn.setFixedWidth(120)
            if any(f.title == cat for f in self.controller.folders):
                btn.setText(I18N.tr("Hide/Delete"))
                btn.setStyleSheet("background: #d32f2f;")
                btn.clicked.connect(lambda _, c=cat: self.controller.delete_default_folder(c))
            else:
                btn.setText(I18N.tr("Add"))
                btn.setStyleSheet("background: #388e3c;")
                btn.clicked.connect(lambda _, c=cat: self.controller.add_default_folder(c))
                
            row_layout.addWidget(btn)
            item = QListWidgetItem(self.defaults_list)
            item.setSizeHint(row.sizeHint())
            self.defaults_list.addItem(item)
            self.defaults_list.setItemWidget(item, row)
            
        defaults_layout.addWidget(self.defaults_list)
        self.tabs.addTab(self.tab_defaults, I18N.tr("Default Folders"))
        
        # قسم الإعدادات العامة
        self.tab_settings = QWidget()
        settings_layout = QFormLayout(self.tab_settings)
        settings_layout.setContentsMargins(20, 30, 20, 20)
        settings_layout.setSpacing(25)
        
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["العربية", "English"])
        self.lang_combo.setCurrentIndex(0 if I18N.LANG == "ar" else 1)
        self.lang_combo.currentIndexChanged.connect(self.change_language)

        self.grid_chk = QCheckBox(I18N.tr("Enable Grid Snapping"))
        self.grid_chk.setChecked(self.controller.snap_to_grid)
        self.grid_chk.toggled.connect(self.controller.toggle_grid)
        
        self.startup_chk = QCheckBox(I18N.tr("Run App on Windows Startup"))
        self.startup_chk.setChecked(self.controller.is_in_startup())
        self.startup_chk.toggled.connect(self.controller.toggle_startup)

        self.inner_slider = QSlider(Qt.Orientation.Horizontal)
        self.inner_slider.setRange(10, 100)
        self.outer_slider = QSlider(Qt.Orientation.Horizontal)
        self.outer_slider.setRange(0, 100)

        if self.controller.folders:
            self.inner_slider.setValue(self.controller.folders[0].opacity_inner)
            self.outer_slider.setValue(self.controller.folders[0].opacity_outer)
        else:
            self.inner_slider.setValue(100)
            self.outer_slider.setValue(10)
            
        self.inner_slider.valueChanged.connect(self.controller.change_all_inner_opacity)
        self.outer_slider.valueChanged.connect(self.controller.change_all_outer_opacity)

        self.reset_btn = QPushButton(I18N.tr("Factory Reset (Reset All Folders)"))
        self.reset_btn.setStyleSheet("background: #d32f2f; margin-top: 15px;")
        self.reset_btn.clicked.connect(self.controller.factory_reset)
        
        settings_layout.addRow(I18N.tr("Language:"), self.lang_combo)
        settings_layout.addRow(self.grid_chk)
        settings_layout.addRow(self.startup_chk)
        settings_layout.addRow(I18N.tr("Inner Opacity:"), self.inner_slider)
        settings_layout.addRow(I18N.tr("Outer Opacity:"), self.outer_slider)
        settings_layout.addRow("", self.reset_btn)
        
        self.tabs.addTab(self.tab_settings, I18N.tr("Preferences"))
        self.refresh_list()

    def refresh_list(self):
        self.folders_list.clear()
        for f in self.controller.folders: self.folders_list.addItem(I18N.tr(f.title))

    def add_new(self):
        self.controller.create_new_folder("New Folder")
        self.refresh_list()

    def change_language(self, index):
        new_lang = "ar" if index == 0 else "en"
        if new_lang != I18N.LANG:
            self.controller.config["lang"] = new_lang
            self.controller.save_all()
            I18N.LANG = new_lang
            
            # تحديث الواجهة والاتجاه بشكل فوري 100% بدون ريستارت
            self.controller.app.setLayoutDirection(Qt.LayoutDirection.RightToLeft if new_lang == "ar" else Qt.LayoutDirection.LeftToRight)
            
            for f in self.controller.folders:
                f.tile_title.setText(I18N.tr(f.title))
                f.expanded_title.setText(I18N.tr(f.title))
                for item in f.app_items: item.name_label.setText(I18N.tr(item.name))
                    
            self.close()
            self.controller.show_settings()


class MainController:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setStyle("Fusion")
        self.app.setQuitOnLastWindowClosed(False)
        self.config = ConfigManager.load_config()
        
        I18N.LANG = self.config.get("lang", "ar")
        if I18N.LANG == "ar": self.app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        
        self.folders = []
        self.snap_to_grid = self.config.get("snap_to_grid", True)

        self.setup_tray()
        self.load_folders()

    def setup_tray(self):
        self.tray = QSystemTrayIcon(self.app.style().standardIcon(QStyle.StandardPixmap.SP_DesktopIcon))
        tray_menu = QMenu()
        
        for name, func in [(I18N.tr("Dashboard"), self.show_settings), (I18N.tr("Exit"), self.quit_app)]:
            action = QAction(name, self.app)
            action.triggered.connect(func)
            tray_menu.addAction(action)
            if name == I18N.tr("Dashboard"): tray_menu.addSeparator()
            
        self.tray.setContextMenu(tray_menu)
        self.tray.show()
        self.tray.activated.connect(lambda r: self.show_settings() if r in (QSystemTrayIcon.ActivationReason.Trigger, QSystemTrayIcon.ActivationReason.DoubleClick) else None)

    def show_settings(self):
        if not hasattr(self, 'settings_win') or not self.settings_win.isVisible():
            self.settings_win = SettingsWindow(self)
            self.settings_win.show()
        else: self.settings_win.activateWindow()

    def scan_all_defaults(self, existing_paths=None):
        if existing_paths is None: existing_paths = set()
        apps = {k: [] for k in ["Games", "Launchers", "Office Apps", "Developer Tools", "Design Apps", "Windows Tools", "Recent"]}
        seen = set(existing_paths)

        def add_app(cat, name, path):
            rp = get_real_path(path)
            if rp not in seen:
                if cat: apps[cat].append({"name": name, "path": path, "args": "", "icon_location": ""})
                seen.add(rp)
                return True
            return False

        sys32 = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "System32")
        windir = os.environ.get("WINDIR", "C:\\Windows")
        win_tools = [
            ("explorer.exe", "File Explorer", windir), ("cmd.exe", "Command Prompt", sys32),
            ("Taskmgr.exe", "Task Manager", sys32), ("control.exe", "Control Panel", sys32),
            ("regedit.exe", "Registry Editor", windir), ("notepad.exe", "Notepad", sys32),
            ("mspaint.exe", "Paint", sys32), ("SnippingTool.exe", "Snipping Tool", sys32),
            ("calc.exe", "Calculator", sys32), ("msinfo32.exe", "System Information", sys32),
            ("devmgmt.msc", "Device Manager", sys32), ("diskmgmt.msc", "Disk Management", sys32),
            ("mstsc.exe", "Remote Desktop", sys32), ("WindowsPowerShell\\v1.0\\powershell.exe", "PowerShell", sys32)
        ]
        for exe, name, base in win_tools:
            p = os.path.join(base, exe)
            if os.path.exists(p): add_app("Windows Tools", name, p)
            
        for base in [r"C:\Program Files\Microsoft Office\root\Office16", r"C:\Program Files (x86)\Microsoft Office\root\Office16"]:
            if os.path.exists(base):
                for exe, name in {"WINWORD.EXE": "Word", "EXCEL.EXE": "Excel", "POWERPNT.EXE": "PowerPoint", "MSACCESS.EXE": "Access"}.items():
                    p = os.path.join(base, exe)
                    if os.path.exists(p): add_app("Office Apps", name, p)
                break

        uncat = []
        search_dirs = [
            os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
            os.path.join(os.environ.get("PUBLIC", ""), "Desktop"),
            os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs"),
            os.path.join(os.environ.get("ALLUSERSPROFILE", ""), r"Microsoft\Windows\Start Menu\Programs")
        ]

        for d in search_dirs:
            if not os.path.exists(d): continue
            for root, _, files in os.walk(d):
                for f in files:
                    path = os.path.join(root, f)
                    rp = get_real_path(path)
                    if rp in seen: continue
                    name = os.path.splitext(f)[0]
                    lower_f = f.lower()
                    
                    if lower_f.endswith(".url"):
                        try:
                            with open(path, "r", encoding="utf-8", errors="ignore") as file:
                                if any(x in file.read().lower() for x in ["steam://", "com.epicgames.launcher://", "battlenet://"]):
                                    add_app("Games", name, path)
                        except: pass
                    elif lower_f.endswith(".lnk"):
                        info = resolve_shortcut(path)
                        target = (info.get("target") or "").lower()
                        if not target or "unins" in target: continue
                        combined = os.path.basename(target).lower() + " " + target
                        
                        if any(x in combined for x in ["steam.exe", "epicgameslauncher", "upc.exe", "eadesktop", "origin.exe", "battle.net", "galaxyclient", "xboxapp", "tlauncher", "lunar client", "badlion", "curseforge", "riotclient"]):
                            add_app("Launchers", name, path)
                        elif any(x in combined for x in ["steamapps", "xboxgames", "epic games\\", "minecraft", "roblox", "genshin impact", "valorant", "league of legends", "counter-strike", "call of duty", "gta", "fivem", "assassin's creed"]) and not any(x in combined for x in ["steam.exe", "epicgameslauncher"]):
                            add_app("Games", name, path)
                        elif any(x in combined for x in ["code.exe", "devenv", "pycharm", "idea", "studio", "git", "postman", "docker", "xampp", "sublime", "notepad++", "eclipse", "unity", "unreal", "godot", "nodejs", "sourcetree"]):
                            add_app("Developer Tools", name, path)
                        elif any(x in combined for x in ["photoshop", "illustrator", "premiere", "afterfx", "lightroom", "figma", "blender", "coreldraw", "indesign", "xd", "artcraft", "krita", "clip studio", "zbrush", "cinema4d", "maya", "autocad", "sketchup", "affinity", "canva", "capcut", "davinci"]):
                            add_app("Design Apps", name, path)
                        elif any(x in combined for x in ["winword", "excel", "powerpnt", "msaccess", "outlook", "notion", "obsidian", "evernote"]):
                            add_app("Office Apps", name, path)
                        else:
                            try: mtime = os.path.getctime(path)
                            except: mtime = 0
                            uncat.append((mtime, name, path))
                            seen.add(rp)

        uncat.sort(key=lambda x: x[0], reverse=True)
        for mtime, name, path in uncat[:12]:
            apps["Recent"].append({"name": name, "path": path, "args": "", "icon_location": ""})

        return apps

    def load_folders(self):
        folders_data = self.config.get("folders", [])
        if not self.config.get("setup_done", False):
            wizard = FirstRunWizard()
            wizard.exec()
            selected = wizard.get_selected()
            
            all_apps = self.scan_all_defaults(set())
            screen = QApplication.primaryScreen().availableGeometry()
            spacing = 110
            total_width = len(selected) * spacing
            start_x = round(((screen.width() - total_width) / 2) / 110) * 110
            y_pos = 0 # سقف الشاشة تماماً
            
            for cat in selected:
                if all_apps.get(cat):
                    self.create_new_folder(cat, apps=all_apps[cat], x=start_x, y=y_pos)
                    start_x += spacing
                    
            self.config["setup_done"] = True
            self.save_all()
        else:
            for f_data in folders_data:
                w = StartFolderWidget(f_data, self)
                w.show()
                self.folders.append(w)

    def get_existing_paths(self):
        paths = set()
        for f in self.folders:
            for app in f.apps: paths.add(get_real_path(app["path"]))
        return paths

    def add_default_folder(self, cat):
        all_apps = self.scan_all_defaults(self.get_existing_paths())
        apps = all_apps.get(cat, [])
        screen = QApplication.primaryScreen().availableGeometry()
        x = round((screen.width() / 2) / 110) * 110
        self.create_new_folder(cat, apps=apps, x=x, y=0)
        if hasattr(self, 'settings_win') and self.settings_win.isVisible():
            self.settings_win.close()
            self.show_settings()

    def delete_default_folder(self, cat):
        for f in self.folders[:]:
            if f.title == cat: self.delete_folder(f.folder_id)
        if hasattr(self, 'settings_win') and self.settings_win.isVisible():
            self.settings_win.close()
            self.show_settings()

    def create_new_folder(self, title, apps=None, x=300, y=0):
        import time
        w = StartFolderWidget({"id": int(time.time()), "title": str(title), "x": x, "y": y, "width": 320, "height": 360, "opacity_inner": 100, "opacity_outer": 10, "apps": apps or []}, self)
        w.show()
        self.folders.append(w)
        self.save_all()

    def delete_folder(self, folder_id):
        for widget in self.folders:
            if widget.folder_id == folder_id:
                widget.close()
                self.folders.remove(widget)
                break
        self.save_all()

    def change_all_inner_opacity(self, value):
        for widget in self.folders:
            widget.opacity_inner = value
            widget.apply_style()
        self.save_all()
        
    def change_all_outer_opacity(self, value):
        for widget in self.folders:
            widget.opacity_outer = value
            widget.apply_style()
        self.save_all()

    def factory_reset(self):
        parent = self.settings_win if hasattr(self, 'settings_win') and self.settings_win.isVisible() else None
        msg = QMessageBox(parent)
        msg.setWindowTitle(I18N.tr("Factory Reset"))
        msg.setText(I18N.tr("Are you sure you want to reset all folders?"))
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg.setWindowFlags(msg.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        
        if msg.exec() == QMessageBox.StandardButton.Yes:
            for widget in self.folders: widget.close()
            self.folders.clear()
            self.config = {"first_run": False, "setup_done": False, "snap_to_grid": True, "folders": [], "lang": I18N.LANG}
            self.save_all()
            if hasattr(self, 'settings_win'): self.settings_win.close()
            self.load_folders()
            if hasattr(self, 'settings_win') and self.settings_win.isVisible():
                self.settings_win.refresh_list()
                self.settings_win.inner_slider.setValue(100)
                self.settings_win.outer_slider.setValue(10)

    def toggle_grid(self, checked):
        self.snap_to_grid = checked
        if checked:
            for f in self.folders:
                if not f.is_expanded: f.move(round(f.x() / 110) * 110, round(f.y() / 110) * 110)
        self.save_all()

    def save_all(self):
        self.config.update({"snap_to_grid": self.snap_to_grid, "folders": [f.get_data() for f in self.folders]})
        ConfigManager.save_config(self.config)

    def is_in_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.CloseKey(key)
            return True
        except FileNotFoundError: return False

    def toggle_startup(self, checked):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            if checked: winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, sys.executable if getattr(sys, 'frozen', False) else f'"{sys.executable}" "{os.path.abspath(__file__)}"')
            else: winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
        except Exception: pass

    def quit_app(self):
        self.save_all()
        self.tray.hide()
        self.app.quit()

    def run(self):
        sys.exit(self.app.exec())

if __name__ == "__main__":
    MainController().run()