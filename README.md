# 🛠️ AK System Cleaner v4.3

> A powerful all-in-one Windows system maintenance, cleanup, and optimization tool with a modern Dark Mode GUI.

**Developed BY Ahmad_Khaled**

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-4.3-brightgreen)

---

## ✨ Features

### 🧹 System Cleaning
- **Remove Virus Remnants** — Scans all drives and removes legacy USB virus files
- **Clean Temp Files** — Deletes temporary files from Windows Temp and Prefetch
- **Empty Recycle Bin** — Automatically clears the Recycle Bin

### ⚡ Performance Optimization
- **Clear DNS Cache** — Flushes DNS for faster browsing
- **Disable Startup Programs** — Lists and manages startup applications
- **Disable Telemetry** — Stops Windows data collection services

### 🔧 System Repair
- **SFC Scan** — Finds and repairs corrupted system files
- **DISM Repair** — Repairs the Windows system image
- **Schedule CHKDSK** — Schedules full disk check on next reboot

### 🎨 User Experience
- **Show Hidden Files** — Makes hidden and system files visible
- **Enable Dark Mode** — Switches Windows to Dark Mode
- **Restart Explorer** — Refreshes Explorer to apply changes

### 🔌 USB Tools
- **USB Cleaner** — Detects USB drives and removes virus files (No Format)

---

## 🖥️ GUI Highlights

| Feature | Details |
|---|---|
| 🌑 Dark Mode UI | Clean dark interface with green accent |
| ✅ Selective Steps | Choose which steps to run |
| 📊 Progress Bar | Visual progress with percentage |
| ⏱️ Timer | Live elapsed time |
| 🎨 Color-Coded Log | Green, Yellow, Red, Blue colors |
| 📋 Export Log | Save to Excel or Text |

---

## 🚀 Getting Started

### Option 1: Download .exe (Easiest)
1. Download `AK_Cleaner_v4.3.exe` from this repository
2. Right-click → Run as Administrator
3. Select steps and click Run Cleaner

### Option 2: Run from Source
```bash
git clone https://github.com/YOUR_USERNAME/AK-System-Cleaner.git
cd AK-System-Cleaner
pip install openpyxl
python AK_Cleaner_v4.3.py
