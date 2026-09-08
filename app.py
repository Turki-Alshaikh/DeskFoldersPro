import sys
import os
import json
import winreg
import ctypes
import re
import subprocess
from ctypes import wintypes
from pathlib import Path
from PyQt6.QtCore import Qt, QPoint, QFileInfo, QSize, QEvent, QPropertyAnimation, QEasingCurve, QRect
from PyQt6.QtGui import QIcon, QAction, QCursor, QFontMetrics, QImage, QPixmap
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QGridLayout, QLabel,
    QPushButton, QToolButton, QMenu, QSystemTrayIcon, QFileDialog, QInputDialog,
    QMessageBox, QSizeGrip, QHBoxLayout, QTabWidget, QFormLayout, QCheckBox
)



APP_NAME = "DeskFoldersPro"
CONFIG_DIR = Path.home() / "AppData" / "Roaming" / APP_NAME
CONFIG_FILE = CONFIG_DIR / "config.json"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)


class SHFILEINFO(ctypes.Structure):
    _fields_ = [
        ("hIcon", wintypes.HICON),
        ("iIcon", ctypes.c_int),
        ("dwAttributes", wintypes.DWORD),
        ("szDisplayName", wintypes.WCHAR * 260),
        ("szTypeName", wintypes.WCHAR * 80)
    ]


def extract_exe_and_args(cmd_str):
    """فصل مسار البرنامج عن الـ arguments (مثل --productcode=pro) لدعم كل اختصارات الألعاب"""
    cmd_str = str(cmd_str).strip()

    # 1. إذا كان المسار داخل علامات تنصيص "C:\path\exe" --args
    if cmd_str.startswith('"'):
        end_quote = cmd_str.find('"', 1)
        if end_quote != -1:
            exe_path = cmd_str[1:end_quote]
            args = cmd_str[end_quote+1:].strip()
            return exe_path, args

    # 2. إذا كان النص كاملاً يمثل ملفاً موجوداً
    if os.path.exists(cmd_str):
        return cmd_str, ""

    # 3. محاولة تفكيك السطر من اليمين لإيجاد ملف الـ exe الصحيح
    parts = cmd_str.split(' ')
    for i in range(len(parts), 0, -1):
        candidate = ' '.join(parts[:i])
        if os.path.exists(candidate):
            return candidate, ' '.join(parts[i:])

    return cmd_str, ""


def resolve_shortcut(lnk_path):
    """يفك تشفير اختصار .lnk ويرجع المسار الحقيقي + الـ arguments + مصدر الأيقونة الأصلي"""
    try:
        from win32com.client import Dispatch
        shell = Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(str(lnk_path))
        return {
            "target": sc.Targetpath or str(lnk_path),
            "args": sc.Arguments or "",
            "icon_location": sc.IconLocation or "",
            "workdir": sc.WorkingDirectory or ""
        }
    except Exception:
        return {"target": str(lnk_path), "args": "", "icon_location": "", "workdir": ""}


def get_native_icon(file_path, icon_location=""):
    """جلب الأيقونة الأصلية مع دعم الألعاب، Steam، اختصارات .lnk/.url، والمسارات المركبة مع arguments"""
    path_str = str(file_path)
    clean_path, _ = extract_exe_and_args(path_str)

    # أولوية قصوى: الأيقونة المحددة داخل بيانات الاختصار (.lnk) نفسه، تماماً مثل سطح المكتب
    if icon_location:
        try:
            icon_file, _, icon_idx_str = icon_location.rpartition(",")
            icon_idx = int(icon_idx_str) if icon_idx_str.strip().lstrip("-").isdigit() else 0
            if icon_file and os.path.exists(icon_file):
                large, small = wintypes.HICON(), wintypes.HICON()
                n = ctypes.windll.shell32.ExtractIconExW(
                    icon_file, icon_idx, ctypes.byref(large), ctypes.byref(small), 1
                )
                if n and large.value:
                    image = QImage.fromHICON(large.value)
                    ctypes.windll.user32.DestroyIcon(large.value)
                    if small.value:
                        ctypes.windll.user32.DestroyIcon(small.value)
                    if not image.isNull():
                        return QIcon(QPixmap.fromImage(image))
        except Exception:
            pass

    # دعم اختصارات إنترنت / ستيم (.url)
    if clean_path.lower().endswith(".url"):
        try:
            with open(clean_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            match = re.search(r"IconFile=(.+)", content, re.IGNORECASE)
            if match:
                icon_path = match.group(1).strip()
                if os.path.exists(icon_path):
                    return QIcon(icon_path)
        except Exception:
            pass

    # استخراج أيقونة ويندوز عالية الدقة من الملف النظيف
    try:
        shfi = SHFILEINFO()
        SHGFI_ICON = 0x00000100
        SHGFI_LARGEICON = 0x00000000

        res = ctypes.windll.shell32.SHGetFileInfoW(
            clean_path,
            0,
            ctypes.byref(shfi),
            ctypes.sizeof(shfi),
            SHGFI_ICON | SHGFI_LARGEICON
        )
        if res and shfi.hIcon:
            image = QImage.fromHICON(shfi.hIcon)
            ctypes.windll.user32.DestroyIcon(shfi.hIcon)
            if not image.isNull():
                return QIcon(QPixmap.fromImage(image))
    except Exception:
        pass

    return QIcon()


class ConfigManager:
    @staticmethod
    def load_config():
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"folders": []}

    @staticmethod
    def save_config(data):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)


class AppItem(QWidget):
    """عنصر التطبيق/اللعبة داخل المجلد مع تشغيل آمن للـ arguments ودعم كامل للاختصارات"""
    def __init__(self, name, path, parent_folder, args="", icon_location=""):
        super().__init__()
        self.name = name
        self.path = path
        self.args = args
        self.icon_location = icon_location
        self.parent_folder = parent_folder

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_btn = QToolButton()
        self.icon_btn.setAutoRaise(True)
        self.icon_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.icon_btn.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)

        self.icon = get_native_icon(path, icon_location)
        self.icon_btn.setIcon(self.icon)

        # بدون أي حدود بيضاء أو خطوط تحديد
        self.icon_btn.setStyleSheet("""
            QToolButton {
                background-color: transparent;
                border: none;
                outline: none;
                border-radius: 12px;
            }
            QToolButton:hover {
                background-color: rgba(255, 255, 255, 0.12);
            }
            QToolButton:pressed {
                background-color: rgba(255, 255, 255, 0.08);
            }
        """)
        self.icon_btn.clicked.connect(self.launch_app)

        self.name_label = QLabel(self.name)
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.name_label.setWordWrap(True)
        self.name_label.setMinimumHeight(35)
        self.name_label.setStyleSheet("color: #E0E0E0; font-family: 'Segoe UI'; font-size: 11.5px; font-weight: 600; border: none; background: transparent; text-align: center;")

        layout.addWidget(self.icon_btn)
        layout.addWidget(self.name_label)

    def update_scale(self, icon_size, text_width):
        widget_width = text_width + 16
        self.setFixedWidth(widget_width) # إجبار الحاوية على عرض ثابت لتوسيط النص
        self.icon_btn.setFixedSize(icon_size + 16, icon_size + 16)
        self.icon_btn.setIconSize(QSize(icon_size, icon_size))
        self.name_label.setFixedWidth(widget_width)
        self.name_label.setText(self.name)

    def launch_app(self):
        try:
            if self.args:
                subprocess.Popen(f'"{self.path}" {self.args}', shell=True)
            elif os.path.exists(self.path):
                os.startfile(self.path)
            else:
                subprocess.Popen(self.path, shell=True)
            self.parent_folder.switch_mode(False)
        except Exception as e:
            try:
                subprocess.Popen(f'start "" "{self.path}" {self.args}', shell=True)
                self.parent_folder.switch_mode(False)
            except Exception as e2:
                QMessageBox.warning(None, "Error", f"Failed to launch file:\n{str(e2)}")

    def contextMenuEvent(self, event):
        self.parent_folder.ignore_deactivate = True
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: #2b2b2b; color: white; border: none; border-radius: 6px; }
            QMenu::item:selected { background-color: #3a3a3a; }
        """)
        delete_action = QAction("Remove from Folder", self)
        delete_action.triggered.connect(lambda: self.parent_folder.remove_app(self.path))
        menu.addAction(delete_action)
        menu.exec(QCursor.pos())
        self.parent_folder.ignore_deactivate = False


class ClickableTitleLabel(QLabel):
    """تسمية تدعم الضغط المزدوج (Double Click) لتغيير اسم المجلد فوراً"""
    def __init__(self, text, rename_callback):
        super().__init__(text)
        self.rename_callback = rename_callback

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.rename_callback()


class StartFolderWidget(QWidget):
    """مجلد Start التفاعلي بدون حدود بيضاء ومع فيزيائية Hold & Drag"""
    def __init__(self, folder_data, controller):
        super().__init__()
        self.controller = controller
        self.folder_id = folder_data.get("id")
        self.title = str(folder_data.get("title", "Games"))
        self.apps = folder_data.get("apps", [])
        self.opacity_level = folder_data.get("opacity", 85)

        self.expanded_w = folder_data.get("width", 320)
        self.expanded_h = folder_data.get("height", 360)
        self.is_expanded = False
        self.ignore_deactivate = False

        self.old_pos = self.pos()
        self.is_dragging = False

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.Tool
            | Qt.WindowType.NoDropShadowWindowHint
            | Qt.WindowType.WindowStaysOnBottomHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAcceptDrops(True)

        # تصفير كل الحدود الخارجية على مستوى النافذة الرئيسية
        self.setStyleSheet("StartFolderWidget, QWidget { border: none; outline: none; background: transparent; }")

        if "x" in folder_data and "y" in folder_data:
            self.move(folder_data["x"], folder_data["y"])

        self.app_items = []
        self.init_ui()
        self.switch_mode(expanded=False)
    def init_ui(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(10, 10, 10, 10)

        # 1. وضع الأيقونة المصغرة على سطح المكتب (بدون أي حدود بيضاء)
        self.preview_container = QWidget()
        preview_layout = QVBoxLayout(self.preview_container)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(6)
        preview_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.tile_box = QWidget()
        self.tile_box.setObjectName("TileBox")
        self.tile_box.setFixedSize(68, 68)
        self.tile_box.setCursor(Qt.CursorShape.PointingHandCursor)
        self.tile_box.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.tile_box.setStyleSheet("""
            #TileBox {
                background-color: rgba(255, 255, 255, 0.08);
                border: none;
                border-radius: 14px;
            }
            #TileBox:hover {
                background-color: rgba(255, 255, 255, 0.14);
            }
        """)
        self.tile_grid = QGridLayout(self.tile_box)
        self.tile_grid.setContentsMargins(8, 8, 8, 8)
        self.tile_grid.setSpacing(4)

        self.tile_title = ClickableTitleLabel(self.title, self.rename_folder)
        self.tile_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tile_title.setFixedWidth(90)
        self.tile_title.setStyleSheet("color: white; font-size: 12px; font-weight: bold; font-family: 'Segoe UI'; border: none; background: transparent;")

        preview_layout.addWidget(self.tile_box)
        preview_layout.addWidget(self.tile_title)

        # 2. وضع المجلد المفتوح (بدون إطار أبيض)
        self.expanded_container = QWidget()
        self.expanded_container.setObjectName("ExpandedContainer")
        self.expanded_container.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.expanded_layout = QVBoxLayout(self.expanded_container)
        self.expanded_layout.setContentsMargins(15, 15, 15, 15)

        header_layout = QHBoxLayout()
        self.expanded_title = ClickableTitleLabel(self.title, self.rename_folder)
        self.expanded_title.setStyleSheet("color: white; font-weight: bold; font-size: 14px; background: transparent; border: none;")

        header_layout.addWidget(self.expanded_title)
        header_layout.addStretch()
        self.expanded_layout.addLayout(header_layout)

        self.full_grid_widget = QWidget()
        self.full_grid = QGridLayout(self.full_grid_widget)
        self.full_grid.setSpacing(10)
        self.full_grid.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.expanded_layout.addWidget(self.full_grid_widget)
        self.expanded_layout.addStretch()

        self.size_grip = QSizeGrip(self.expanded_container)
        self.size_grip.setStyleSheet("width: 14px; height: 14px; background: transparent;")
        self.expanded_layout.addWidget(self.size_grip, 0, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)

        self.main_layout.addWidget(self.preview_container)
        self.main_layout.addWidget(self.expanded_container)

        self.tile_box.mousePressEvent = self.tile_mouse_press
        self.tile_box.mouseMoveEvent = self.tile_mouse_move
        self.tile_box.mouseReleaseEvent = self.tile_mouse_release

        self.apply_style()
        self.refresh_apps()

    # --- Hold and Drag في الوضع المصغر ---
    def tile_mouse_press(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.old_pos = event.globalPosition().toPoint()
            self.is_dragging = False

    def tile_mouse_move(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton:
            delta = QPoint(event.globalPosition().toPoint() - self.old_pos)
            if delta.manhattanLength() > 3:
                self.is_dragging = True
                self.move(self.x() + delta.x(), self.y() + delta.y())
                self.old_pos = event.globalPosition().toPoint()

    def tile_mouse_release(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if self.is_dragging:
                if getattr(self.controller, 'snap_to_grid', True):
                    self.move(round(self.x() / 80) * 80, round(self.y() / 80) * 80)
                self.controller.save_all()
                self.is_dragging = False
            else:
                self.switch_mode(True)

    def switch_mode(self, expanded):
        self.is_expanded = expanded
        if not expanded:
            if hasattr(self, 'saved_tile_pos'):
                self.move(self.saved_tile_pos)
            self.expanded_container.hide()
            self.preview_container.show()
            self.setFixedSize(95, 115)
        else:
            for folder in self.controller.folders:
                if folder != self and folder.is_expanded:
                    folder.switch_mode(False)

            self.saved_tile_pos = self.pos()
            
            new_w = max(350, self.expanded_w)
            new_h = max(400, self.expanded_h)
            self.setMinimumSize(350, 400)
            self.setMaximumSize(1200, 1200)
            
            screen_geom = QApplication.primaryScreen().availableGeometry()
            offset_x = (new_w - 95) // 2
            offset_y = 70  # زيادة النزول لأسفل لتجنب التداخل مع الأيقونات الأخرى
            
            new_x = self.x() - offset_x
            new_y = self.y() + offset_y
            
            if new_x < screen_geom.left(): new_x = screen_geom.left() + 10
            if new_y < screen_geom.top(): new_y = screen_geom.top() + 10
            if new_x + new_w > screen_geom.right(): new_x = screen_geom.right() - new_w - 10
            if new_y + new_h > screen_geom.bottom(): new_y = screen_geom.bottom() - new_h - 10
            
            self.preview_container.hide()
            self.expanded_container.show()
            
            if not hasattr(self, 'anim'):
                self.anim = QPropertyAnimation(self, b"geometry")
                self.anim.setDuration(220)
                self.anim.setEasingCurve(QEasingCurve.Type.OutCubic)
                
            self.anim.setStartValue(QRect(self.x(), self.y(), 95, 115))
            self.anim.setEndValue(QRect(new_x, new_y, new_w, new_h))
            self.anim.start()
            
            self.refresh_apps()
            self.raise_()
            self.activateWindow()
    def changeEvent(self, event):
        super().changeEvent(event)
        if (event.type() == QEvent.Type.ActivationChange
                and not self.isActiveWindow()
                and self.is_expanded
                and not self.ignore_deactivate):
            self.switch_mode(False)
    def hideEvent(self, event):
        # محاولة منع ويندوز من إخفاء النافذة عند الضغط على Win+D
        if not self.isHidden():
            event.ignore()
            self.show()
        else:
            super().hideEvent(event)
    def apply_style(self):
        alpha = int((self.opacity_level / 100) * 255)
        self.expanded_container.setStyleSheet(f"""
            #ExpandedContainer {{
                background-color: rgba(20, 20, 20, {alpha});
                border: 1px solid rgba(255, 255, 255, 0.15);
                border-radius: 16px;
            }}
        """)

    def calculate_dynamic_scale(self):
        folder_w = self.width()
        cols = max(3, folder_w // 90)
        cell_w = max(70, (folder_w - 40) // cols)
        icon_size = max(36, min(64, int(cell_w * 0.55)))
        text_w = max(65, cell_w - 8)
        return cols, icon_size, text_w

    def refresh_apps(self):
        valid_apps = []
        for app in self.apps:
            path = app["path"]
            if path.lower().startswith("steam://") or path.lower().endswith(".url") or os.path.exists(path):
                valid_apps.append(app)
        
        if len(valid_apps) != len(self.apps):
            self.apps = valid_apps
            self.controller.save_all()

        for idx, app in enumerate(self.apps[:4]):
            r, c = divmod(idx, 2)
            icon = get_native_icon(app["path"], app.get("icon_location", ""))
            icon_lbl = QLabel()
            icon_lbl.setPixmap(icon.pixmap(22, 22))
            icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            icon_lbl.setStyleSheet("background: transparent; border: none;")
            self.tile_grid.addWidget(icon_lbl, r, c)

        for i in reversed(range(self.full_grid.count())):
            w = self.full_grid.itemAt(i).widget()
            if w: w.deleteLater()

        self.app_items.clear()
        cols, icon_size, text_w = self.calculate_dynamic_scale()

        row, col = 0, 0
        for app in self.apps:
            item = AppItem(
                app["name"], app["path"], self,
                app.get("args", ""), app.get("icon_location", "")
            )
            item.update_scale(icon_size, text_w)
            self.full_grid.addWidget(item, row, col)
            self.app_items.append(item)
            col += 1
            if col >= cols:
                col = 0
                row += 1

    def add_app(self, path):
        path_str = str(path).strip()

        args = ""
        icon_location = ""

        if path_str.lower().endswith(".lnk"):
            info = resolve_shortcut(path_str)
            target = info["target"] or path_str
            args = info["args"]
            icon_location = info["icon_location"]
        else:
            target, args = extract_exe_and_args(path_str)

        if not os.path.exists(target) and not path_str.lower().startswith("steam://"):
            return
        name = Path(target).stem if target else "App"
        if any(app["path"] == target for app in self.apps):
            return
        self.apps.append({
            "name": name,
            "path": target,
            "args": args,
            "icon_location": icon_location
        })
        self.refresh_apps()
        self.controller.save_all()

    def remove_app(self, path):
        self.apps = [app for app in self.apps if app["path"] != path]
        self.refresh_apps()
        self.controller.save_all()

    def mousePressEvent(self, event):
        if self.is_expanded: return # منع التحريك
        if event.button() == Qt.MouseButton.LeftButton:
            self.old_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self.is_expanded: return # منع التحريك
        if event.buttons() == Qt.MouseButton.LeftButton:
            delta = QPoint(event.globalPosition().toPoint() - self.old_pos)
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.old_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event):
        if self.is_expanded: return
        if getattr(self.controller, 'snap_to_grid', True):
            self.move(round(self.x() / 80) * 80, round(self.y() / 80) * 80)
        self.controller.save_all()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path:
                self.add_app(file_path)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.is_expanded:
            self.expanded_w = self.width()
            self.expanded_h = self.height()

            cols, icon_size, text_w = self.calculate_dynamic_scale()
            for item in self.app_items:
                item.update_scale(icon_size, text_w)

            self.controller.save_all()

    def contextMenuEvent(self, event):
        self.ignore_deactivate = True
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: #1e1e1e; color: #e0e0e0; border: 1px solid #333; border-radius: 8px; font-family: 'Segoe UI'; font-size: 12px; padding: 4px; }
            QMenu::item { padding: 8px 30px 8px 20px; margin: 2px; border-radius: 4px; }
            QMenu::item:selected { background-color: #0067b8; color: white; }
            QMenu::separator { height: 1px; background: #333; margin: 4px 10px; }
        """)

        toggle_action = QAction("Expand / Collapse", self)
        toggle_action.triggered.connect(lambda: self.switch_mode(not self.is_expanded))
        menu.addAction(toggle_action)

        add_action = QAction("Add App Manually...", self)
        add_action.triggered.connect(self.manual_add_app)
        menu.addAction(add_action)

        rename_action = QAction("Rename Folder", self)
        rename_action.triggered.connect(self.rename_folder)
        menu.addAction(rename_action)

        opacity_action = QAction("Change Opacity", self)
        opacity_action.triggered.connect(self.change_opacity)
        menu.addAction(opacity_action)

        menu.addSeparator()

        del_action = QAction("Delete Folder", self)
        del_action.triggered.connect(lambda: self.controller.delete_folder(self.folder_id))
        menu.addAction(del_action)

        menu.exec(QCursor.pos())
        self.ignore_deactivate = False

    def manual_add_app(self):
        self.ignore_deactivate = True
        path, _ = QFileDialog.getOpenFileName(
            None, "Select an App or Game", "",
            "All Supported (*.exe;*.lnk;*.url);;Executables (*.exe);;Shortcuts (*.lnk;*.url);;All Files (*.*)"
        )
        if path: self.add_app(path)
        self.ignore_deactivate = False

    def rename_folder(self):
        self.ignore_deactivate = True
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Rename Folder")
        dialog.setLabelText("New Name:")
        dialog.setTextValue(self.title)
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        ok = dialog.exec()
        new_title = dialog.textValue()
        if ok and new_title.strip():
            self.title = new_title.strip()
            self.tile_title.setText(self.title)
            self.expanded_title.setText(self.title)
            self.controller.save_all()
        self.ignore_deactivate = False

    def change_opacity(self):
        self.ignore_deactivate = True
        val, ok = QInputDialog.getInt(None, "Opacity", "Select Opacity (10-100):", self.opacity_level, 10, 100, 5)
        if ok:
            self.opacity_level = val
            self.apply_style()
            self.controller.save_all()
        self.ignore_deactivate = False

    def get_data(self):
        return {
            "id": self.folder_id,
            "title": self.title,
            "x": self.x(),
            "y": self.y(),
            "width": self.expanded_w,
            "height": self.expanded_h,
            "opacity": self.opacity_level,
            "apps": self.apps
        }

from PyQt6.QtWidgets import QDialog, QListWidget

class SettingsWindow(QDialog):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.setWindowTitle("DeskFoldersPro Dashboard")
        self.resize(500, 380)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        
        # تصميم CSS احترافي للوحة التحكم
        self.setStyleSheet("""
            QDialog { background-color: #161616; color: #ffffff; }
            QTabWidget::pane { border: 1px solid #333; border-radius: 6px; background: #1e1e1e; }
            QTabBar::tab { background: #161616; border: 1px solid #333; padding: 10px 25px; color: #888; border-top-left-radius: 6px; border-top-right-radius: 6px; font-weight: bold; font-family: 'Segoe UI'; }
            QTabBar::tab:selected { background: #1e1e1e; color: #0078d4; border-bottom-color: #1e1e1e; }
            QPushButton { background-color: #0067b8; color: white; border: none; padding: 10px 15px; border-radius: 6px; font-weight: bold; font-family: 'Segoe UI'; font-size: 13px; }
            QPushButton:hover { background-color: #0078d4; }
            QListWidget { background-color: #161616; border: 1px solid #333; border-radius: 6px; color: white; padding: 5px; font-size: 13px; }
            QListWidget::item { padding: 10px; border-radius: 4px; border-bottom: 1px solid #222; }
            QListWidget::item:selected { background-color: #0067b8; font-weight: bold; }
            QCheckBox { color: #ccc; font-size: 14px; font-family: 'Segoe UI'; }
            QCheckBox::indicator { width: 18px; height: 18px; border-radius: 4px; border: 1px solid #555; background: #1e1e1e; }
            QCheckBox::indicator:checked { background: #0067b8; border: 1px solid #0067b8; }
        """)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # --- تبويب إدارة المجلدات ---
        self.tab_folders = QWidget()
        folders_layout = QVBoxLayout(self.tab_folders)
        folders_layout.setSpacing(15)
        
        self.folders_list = QListWidget()
        folders_layout.addWidget(self.folders_list)
        
        btn_layout = QHBoxLayout()
        add_btn = QPushButton(" Create New Folder")
        add_btn.clicked.connect(self.add_new)
        btn_layout.addWidget(add_btn)
        folders_layout.addLayout(btn_layout)
        
        self.tabs.addTab(self.tab_folders, " Folders")
        
        # --- تبويب الإعدادات المتقدمة ---
        self.tab_settings = QWidget()
        settings_layout = QFormLayout(self.tab_settings)
        settings_layout.setContentsMargins(20, 25, 20, 20)
        settings_layout.setSpacing(20)
        
        self.grid_chk = QCheckBox("Enable Grid Snapping (Align to Desktop Grid)")
        self.grid_chk.setChecked(self.controller.snap_to_grid)
        self.grid_chk.toggled.connect(self.controller.toggle_grid)
        
        settings_layout.addRow("", self.grid_chk)
        self.tabs.addTab(self.tab_settings, " Preferences")
        
        self.refresh_list()

    def refresh_list(self):
        self.folders_list.clear()
        for f in self.controller.folders:
            self.folders_list.addItem(f.title)

    def add_new(self):
        self.controller.create_new_folder("New Folder")
        self.refresh_list()
class MainController:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setStyle("Fusion")
        self.app.setQuitOnLastWindowClosed(False)
        self.folders = []
        self.config = ConfigManager.load_config()
        self.snap_to_grid = self.config.get("snap_to_grid", True)

        self.setup_tray()
        self.load_folders()

    def setup_tray(self):
        self.tray = QSystemTrayIcon()
        self.tray.setIcon(self.app.style().standardIcon(self.app.style().StandardPixmap.SP_DesktopIcon))

        tray_menu = QMenu()
        
        settings_action = QAction("Settings & Folders", self.app)
        settings_action.triggered.connect(self.show_settings)
        tray_menu.addAction(settings_action)

        self.grid_action = QAction("Enable Grid Snap", self.app, checkable=True)
        self.grid_action.setChecked(self.snap_to_grid)
        self.grid_action.triggered.connect(self.toggle_grid)
        tray_menu.addAction(self.grid_action)

        tray_menu.addSeparator()

        self.startup_action = QAction("Run at Startup", self.app, checkable=True)
        self.startup_action.setChecked(self.is_in_startup())
        self.startup_action.triggered.connect(self.toggle_startup)
        tray_menu.addAction(self.startup_action)

        tray_menu.addSeparator()

        quit_action = QAction("Quit", self.app)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)

        self.tray.setContextMenu(tray_menu)
        self.tray.show()
        
        self.tray.activated.connect(self.on_tray_click)

    def on_tray_click(self, reason):
        if reason in (QSystemTrayIcon.ActivationReason.Trigger, QSystemTrayIcon.ActivationReason.DoubleClick):
            self.show_settings()

    def show_settings(self):
        if not hasattr(self, 'settings_win') or not self.settings_win.isVisible():
            self.settings_win = SettingsWindow(self)
            self.settings_win.show()
        else:
            self.settings_win.activateWindow()

    def load_folders(self):
        folders_data = self.config.get("folders", [])
        if not folders_data:
            self.create_new_folder("Games")
        else:
            for f_data in folders_data:
                widget = StartFolderWidget(f_data, self)
                widget.show()
                self.folders.append(widget)

    def create_new_folder(self, title="New Folder"):
        import time
        folder_data = {
            "id": int(time.time()), "title": str(title),
            "x": 300, "y": 200, "width": 320, "height": 360,
            "opacity": 85, "apps": []
        }
        widget = StartFolderWidget(folder_data, self)
        widget.show()
        self.folders.append(widget)
        self.save_all()

    def delete_folder(self, folder_id):
        for widget in self.folders:
            if widget.folder_id == folder_id:
                widget.close()
                self.folders.remove(widget)
                break
        self.save_all()
    def toggle_grid(self, checked):
        self.snap_to_grid = checked
        if checked:
            for f in self.folders:
                if not f.is_expanded:
                    f.move(round(f.x() / 80) * 80, round(f.y() / 80) * 80)
        self.save_all()
    def save_all(self):
        ConfigManager.save_config({
            "snap_to_grid": getattr(self, 'snap_to_grid', True),
            "folders": [f.get_data() for f in self.folders]
        })

    def is_in_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, APP_NAME)
            winreg.CloseKey(key)
            return True
        except FileNotFoundError: return False

    def toggle_startup(self, checked):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            if checked:
                exec_path = sys.executable if getattr(sys, 'frozen', False) else f'"{sys.executable}" "{os.path.abspath(__file__)}"'
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, exec_path)
            else: winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
        except Exception as e:
            QMessageBox.warning(None, "Error", f"Startup Settings Error:\n{str(e)}")

    def quit_app(self):
        self.save_all()
        self.tray.hide()
        self.app.quit()

    def run(self):
        sys.exit(self.app.exec())


if __name__ == "__main__":
    controller = MainController()
    controller.run()