# 📖 Study Tracker

A simple study tracker with a Tkinter GUI and Excel logging (via `openpyxl`).

---

## ✨ Features
- Start and stop study sessions with one click
- Logs data in `study_log.xlsx`
- Tracks daily totals + comparison to previous day
- Clean, minimal GUI (Tkinter)

---

## 🚀 How to Run

### Option 1: Run with Python
1. Clone this repo:
```bash
   git clone https://github.com/athular-me/study-tracker.git
   cd study-tracker 
```
2.Install dependencies:
```bash
    pip install -r requirements.txt
```
3.Run the script:
```bash
python study_tracker.py
```
### option 2: Build a Windows Executable
1.Open CMD in the repo folder:
```bash
cd C:\Users\athul\OneDrive\Desktop\study-tracker
```
2.Build the .exe:
```bash
pyinstaller --onefile --windowed study_tracker.py
```
3.Find your app inside:
```bash
dist/study_tracker.exe
```
4.(Optional) Pin the .exe to your Taskbar for quick access.