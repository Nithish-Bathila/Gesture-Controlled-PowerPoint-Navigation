
# Gesture Controlled PowerPoint Navigation

🚀 **A touchless system to navigate PowerPoint slides using hand gestures detected from your webcam.**

---

## ✨ Features

### ✅ Real-Time Gesture Detection

Detects:

- ✋ **Palm**: Start slideshow
- ✊ **Fist**: End slideshow
- ☝️ **One finger**: Next slide
- ✌️ **Two fingers**: Previous slide

---

### ✅ PowerPoint Control

Uses **pywin32 COM** to automate:

- Starting slideshow
- Ending slideshow
- Moving to next or previous slides

---

### ✅ System Tray Integration

Runs silently with tray options:

- Show/hide camera preview
- View gesture reference
- About window
- Exit application

---

### ✅ Hold Confirmation

Gesture must be **held for 2 seconds** to avoid accidental triggers.

---

## 🛠 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Nithish-Bathila/gesture-controlled-presentation.git
cd gesture-controlled-presentation
```

---

### 2. Create Virtual Environment & Install Dependencies

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

---

## 🎮 Usage

1. **Ensure PowerPoint is open** with your presentation.  
2. Run:

```bash
python main.py
```

✅ The app will:

- Start webcam and detect gestures.
- Control PowerPoint slides accordingly.
- Show a **system tray icon** for controls.

---

## 📂 Project Structure

```
main.py                 # Entry point with system tray integration
controller.py           # PowerPoint control module (pywin32)
gesture_detector.py     # Gesture detection module (MediaPipe)
gestures.png            # Gesture reference image
icon.png / icon.ico     # Application icons
install_script.iss      # Inno Setup installer script
requirements.txt        # Python dependencies
.gitignore
```

---

## 💻 Technologies Used

- **Python 3.11**
- **OpenCV**
- **MediaPipe**
- **pywin32**
- **pystray**
- **tkinter**

---

## ⚠️ Limitations

- Requires good lighting for accurate detection.
- Only supports **right hand gestures** as implemented.
- Works on **Windows only** due to PowerPoint COM integration.

---
