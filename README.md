# PPT to PDF Converter

A lightweight utility that converts Microsoft PowerPoint files (`.ppt`, `.pptx`) to PDF. Supports both Windows and Linux operating systems.

- **Windows**: Uses PowerPoint's native COM automation, ensuring 100% accurate layouts, fonts, and styling.
- **Linux**: Uses LibreOffice in headless mode for conversion.

## 📥 Download

### Windows
1. Go to the [Releases](https://github.com/rishit537/PPTtoPDF/releases/) page.
2. Download the latest `PPTtoPDF.exe` file.
3. Place it on your Desktop or anywhere convenient.

### Linux
1. Clone the repository or download the source code.
2. Ensure you have Python installed (Python 3.x recommended).
3. Install LibreOffice if not already installed:
   - Ubuntu/Debian: `sudo apt install libreoffice`
   - Fedora: `sudo dnf install libreoffice`
   - Arch: `sudo pacman -S libreoffice-fresh`

## 🚀 Usage

### 1. Drag and Drop (Windows Only)
Simply select one or more `.ppt`/`.pptx` files—or a whole folder—and **drag and drop them directly onto the `.exe` file**.

### 2. Interactive Mode
Run the script/executable without arguments. A console window will open and ask you to enter the path to your file or folder.

### 3. Command Line
You can also run it from the command prompt or terminal:

**Windows:**
```cmd
PPTtoPDF.exe "C:\path\to\your\presentations.pptx" ...
```

**Linux:**
```bash
python main.py "/path/to/your/presentations.pptx" ...
```

## ⚙️ Features
- **Cross-Platform:** Works on both Windows and Linux.
- **Batch Processing:** Provide a folder path to convert all PowerPoint files inside it automatically.
- **Drag-and-Drop (Windows):** Windows natively supports dragging files onto `.exe` programs.
- **Optional Cleanup (Windows):** Built-in prompt to delete the original PowerPoint files after conversion.
- **High Accuracy:** 
  - Windows: Automates the actual Microsoft PowerPoint application.
  - Linux: Uses LibreOffice for reliable conversion.

## ⚠️ Prerequisites

### Windows
- **Microsoft PowerPoint** must be installed on your system.

### Linux
- **LibreOffice** must be installed on your system.

## Notes:
1. **Windows only:** After a successful PDF conversion, the tool will ask if you want to automatically delete the original PowerPoint files to save space.
2. When providing the path to a folder, the tool converts all the PowerPoint presentations to PDF inside the folder. Other files remain untouched.

---

## 👨‍💻 For Developers (Building from Source)

If you want to run the Python script directly or build the `.exe` yourself:

1. Clone the repository and navigate to the folder.
2. Install the required dependencies using `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the script:
   ```bash
   python main.py [path/to/file or directory]
   ```
4. **To build the executable yourself (Windows only)** using PyInstaller:
   ```bash
   pip install pyinstaller
   pyinstaller --onefile --icon=NONE main.py
   ```
   *Your compiled executable will be located in the `dist/` folder.*

## 📄 License
MIT License. Feel free to modify and distribute as needed.