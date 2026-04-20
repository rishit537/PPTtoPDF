# PPT to PDF Converter

A lightweight, drag-and-drop Windows utility that converts Microsoft PowerPoint files (`.ppt`, `.pptx`) to PDF. It uses PowerPoint's native COM automation, ensuring 100% accurate layouts, fonts, and styling. 

## 📥 Download

1. Go to the [Releases](https://github.com/rishit537/PPTtoPDF/releases/) page.
2. Download the latest `PPTtoPDF.exe` file.
3. Place it on your Desktop or anywhere convenient.

## 🚀 Usage

### 1. Drag and Drop (Easiest)
Simply select one or more `.ppt`/`.pptx` files—or a whole folder—and **drag and drop them directly onto the `.exe` file**.

### 2. Interactive Mode
Double-click the `.exe` file. A console window will open and ask you to paste the path to your file or folder.

### 3. Command Line
You can also run it from the command prompt or terminal:
```cmd
PPTtoPDF.exe "C:\path\to\your\presentations.pptx" ... 
```


## ⚙️ Features
- **Batch Processing:** Drop a whole folder on the executable to convert all PowerPoint files inside it automatically.
- **Drag-and-Drop Ready:** Windows natively supports dragging files onto `.exe` programs.
- **Optional Cleanup:** Built-in prompt to delete the original PowerPoint files after conversion.
- **Native Accuracy:** Automates the actual Microsoft PowerPoint application in the background, so your PDFs look exactly like your slides.

## ⚠️ Prerequisites
- **Windows OS** is required.
- **Microsoft PowerPoint** must be installed on your system.

## Notes:
1. After a successful PDf conversion, the tool will ask if you want to automatically delete the original PowerPoint files to save space.
2. When dropping or providing the path to a folder, the tool converts all the PowerPoint presentations to PDF inside the folder. Other files remain untouched.

---

## 👨‍💻 For Developers (Building from Source)

If you want to run the Python script directly or build the `.exe` yourself:

1. Clone the repository and navigate to the folder.
2. Install the required dependencies using `requirements.txt`:
   ```bash
   pip install -r .\requirements.txt
   ```
3. Run the script:
   ```bash
   python main.py
   ```
4. **To build the executable yourself** using PyInstaller:
   ```bash
   pip install pyinstaller
   pyinstaller --onefile --icon=NONE main.py
   ```
   *Your compiled executable will be located in the `dist/` folder.*

## 📄 License
MIT License. Feel free to modify and distribute as needed.
