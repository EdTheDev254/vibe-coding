# 🖼️ Screenshot Text Extractor

A simple app that lets you **grab a part of your screen and extract text from it** using OCR (Tesseract). Works with a click, highlights the area, and copies the text to your clipboard—easy and quick.

### 🔧 What It Does
- You click a button.
- Select an area of the screen.
- It captures that region, enhances the image, and runs OCR.
- Text gets shown in the app **and copied to your clipboard** automatically.

### 🧠 Behind the Scenes
- Built with **Tkinter** for the GUI.
- Uses **Pillow + Tesseract (via pytesseract)** for image handling & OCR.
- Can find a bundled Tesseract or use the one installed in your system.
- Includes basic preprocessing: grayscale, thresholding, resizing to improve OCR accuracy.

### 📦 Requirements
- Python 3
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed (or bundled)
- Python packages:
  - `pytesseract`
  - `Pillow`
  - `pyperclip`

### ⚙️ Run It
```bash
python text-extractor.py
