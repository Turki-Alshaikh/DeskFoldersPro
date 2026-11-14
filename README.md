# 📁 DeskFoldersPro

<div align="center">
  <img src=".ico/App_Icon.ico" alt="DeskFoldersPro Logo" width="120" height="120">
  <br>
  <strong>A Smart, Beautiful, and Intelligent Desktop Organizer for Windows.</strong>
  <br><br>
  
  [![Windows](https://img.shields.io/badge/Platform-Windows-0078d4.svg?logo=windows)](#)
  [![Python](https://img.shields.io/badge/Python-3.8+-3776AB.svg?logo=python)](#)
  [![PyQt6](https://img.shields.io/badge/GUI-PyQt6-41CD52.svg)](#)
  [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
</div>

---

## 🌟 Overview

**DeskFoldersPro** is a powerful, highly customizable desktop organization tool for Windows. It allows you to group your shortcuts, apps, and files into sleek, semi-transparent folders that sit directly on your desktop. 

Powered by a smart categorization engine, it automatically detects your installed applications and groups them into intelligent folders like *Games, Developer Tools, Design Apps, and Windows Tools*. Say goodbye to a cluttered desktop!

## ✨ Key Features

*   🧠 **Smart Categorization Engine:** Automatically scans and groups your apps into predefined categories (Games, Launchers, Dev Tools, Design, Office, Windows Tools, Recent).
*   🎨 **Customizable Glass UI:** Independent opacity controls for both the folder widget (Outer Opacity) and the expanded view (Inner Opacity) for a beautiful Acrylic/Glassmorphism look.
*   📐 **Smart Grid System & Collision Detection:** Folders snap perfectly to a 110px grid. The built-in collision engine ensures folders never overlap.
*   🌐 **Instant Bilingual Support:** Fully supports English and Arabic (RTL/LTR). Switch languages in real-time without restarting the app.
*   🖼️ **Native Icon Extraction:** Bypasses messy shortcut arrows! It extracts high-quality icons directly from `.exe`, `.lnk`, `.url`, or uses custom image files.
*   🖱️ **Drag & Drop:** Easily reorder apps within a folder, or drop new files/apps directly into an open folder.
*   🚀 **Run on Startup:** Seamless integration with Windows startup via the system tray dashboard.

## 📸 Screenshots

*(Replace these placeholders with your actual screenshots)*

| Dashboard / Settings | Desktop Folders (Collapsed) | Folder (Expanded) |
| :---: | :---: | :---: |
| `<img src="path/to/settings.png" width="250">` | `<img src="path/to/collapsed.png" width="250">` | `<img src="path/to/expanded.png" width="250">` |

## 🚀 Installation & Usage

### For Regular Users (Executable)
1. Download the latest `DeskFoldersPro.exe` from the [Releases](#) page.
2. Run the executable. 
3. On the first run, the **Welcome Wizard** will help you generate your smart default folders.
4. Access the **Dashboard** at any time by double-clicking the app icon in your system tray (bottom right of your screen).

### For Developers (Run from Source)
1. Clone the repository:
   ```bash
   git clone https://github.com/YourUsername/DeskFoldersPro.git
   cd DeskFoldersPro
   ```
2. Install the required dependencies:
   ```bash
   pip install PyQt6 pywin32
   ```
3. Run the application:
   ```bash
   python DeskFoldersPro.py
   ```

## 🛠️ Building the `.exe`

To package this app into a standalone Windows executable using PyInstaller:

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --icon=".ico\App_Icon.ico" DeskFoldersPro.py
```
*The output file will be located in the `dist/` folder.*

## ⚙️ Configuration
All settings and folder states are securely saved in your local AppData folder:
`C:\Users\YourUsername\AppData\Roaming\DeskFoldersPro\config.json`

You can easily Factory Reset the app from the Dashboard if you want to start over.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! 
Feel free to check [issues page](#). 

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

Distributed under the MIT License. See `LICENSE` for more information.

---
<div align="center">
  Made with ❤️ using Python & PyQt6
</div>
