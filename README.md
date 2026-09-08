# 📁 DeskFoldersPro

**A modern, elegant, and native-feeling desktop organizer for Windows.**  
*Bring iOS-like interactive folders straight to your Windows desktop without compromising system resources.*

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![PyQt6](https://img.shields.io/badge/UI-PyQt6-green.svg?style=for-the-badge&logo=qt&logoColor=white)](https://pypi.org/project/PyQt6/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)
[![Windows](https://img.shields.io/badge/Platform-Windows_10_|_11-0078d7.svg?style=for-the-badge&logo=windows&logoColor=white)]()

<br/>
<img src="https://via.placeholder.com/800x400.png?text=Add+a+Screenshot+of+DeskFoldersPro+Here" alt="DeskFoldersPro Screenshot" width="800"/>
<br/>

</div>

---

## ✨ Features

- **🍎 iOS-Style Folders:** Interactive hold-and-drag mechanics with smooth, eased expansion animations.
- **🎨 Glassmorphism UI:** Modern translucent design with adjustable opacity that perfectly blends with Windows 11 aesthetics.
- **⚡ Native Integration:** Anchors natively to your desktop layer—fully bypasses the `Win + D` hiding issue (Uses `WindowStaysOnBottomHint`).
- **🎮 Game & Steam Support:** Automatically fetches native high-res icons for `.exe`, `.lnk`, `.url`, and specific Steam game shortcuts.
- **🎛️ Dashboard Settings:** An intuitive, tabbed tray-based dashboard to manage folders and preferences.
- **📐 Grid Snapping:** Keep your desktop clean with auto-alignment to the desktop grid.
- **🚀 Auto-Startup:** Run seamlessly in the background via Windows Registry.

---

## 📥 Installation

### Option 1: Quick Install (For Regular Users)
The easiest way to get started is by downloading the standalone pre-compiled executable.

1. Navigate to the [Releases](../../releases) page.
2. Download the latest `DeskFoldersPro.exe`.
3. Run the executable (Portable - No installation required!).
4. *Tip: Right-click the system tray icon to enable "Run at Startup".*

### Option 2: Build from Source (For Developers)
To run or build the project from the source code:

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/Turki-Alshaikh/DeskFoldersPro.git](https://github.com/Turki-Alshaikh/DeskFoldersPro.git)
   cd DeskFoldersPro

```

2. **Install dependencies:**
```bash
pip install PyQt6 pywin32

```


3. **Run the app directly:**
```bash
python app.py

```


4. **Build the silent executable (Using PyInstaller):**
```bash
pyinstaller --noconsole --onefile --windowed app.py

```



---

## 🛠️ Usage Guide

1. **Create Folders:** Right-click the system tray icon and open the **Settings Dashboard** to create a new folder.
2. **Add Apps:** Drag and drop any shortcut (`.lnk`, `.url`, or `.exe`) directly into an expanded folder, or right-click the folder and select "Add App Manually".
3. **Move Folders:** Click and hold the folder icon on your desktop to drag it around. Ensure **Enable Grid Snap** is checked in the settings for precise alignment.
4. **Rename & Customize:** Double-click the folder title at any time to quickly rename it, or right-click to adjust its opacity.

---

## 🤝 Contributing

Contributions, issues, and feature requests are highly welcome!
If you have suggestions for improving the pixel-perfect alignment or adding new animations, feel free to check the [issues page](https://www.google.com/search?q=../../issues) or submit a Pull Request.

---

## 📝 License

This project is licensed under the **MIT License** - see the [LICENSE](https://www.google.com/search?q=LICENSE) file for details.

