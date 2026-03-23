# 🛠️ AK System Cleaner v4.3

> A powerful all-in-one Windows system maintenance, cleanup, and optimization tool with a modern Dark Mode GUI.

**Developed BY Ahmad_Khaled**

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-4.3-brightgreen)

---

## 📸 Screenshot

<!-- Add your screenshot here -->
![AK System Cleaner](screenshot.png)

---

## ✨ Features

### 🧹 System Cleaning
- **Remove Virus Remnants** — Scans all drives and removes legacy USB virus files (autorun.inf, copy.exe, host.exe, rose.exe, etc.)
- **Clean Temp Files** — Deletes temporary files from Windows\Temp, Prefetch, and user AppData
- **Empty Recycle Bin** — Automatically clears the Recycle Bin

### ⚡ Performance Optimization
- **Clear DNS Cache** — Flushes DNS resolver cache for faster browsing
- **Disable Startup Programs** — Lists and helps manage startup applications
- **Disable Telemetry** — Stops Windows data collection services for better privacy and performance

### 🔧 System Repair
- **SFC Scan** — Runs System File Checker to find and repair corrupted system files
- **DISM Repair** — Repairs the Windows system image when SFC can't fix the issue
- **Schedule CHKDSK** — Schedules a full disk check for bad sectors on next reboot

### 🎨 User Experience
- **Show Hidden Files** — Makes all hidden and system files visible in Explorer
- **Enable Dark Mode** — Switches Windows to system-wide Dark Mode
- **Restart Explorer** — Instantly refreshes Windows Explorer to apply changes

### 🔌 USB Tools
- **USB Cleaner** — Detects connected USB drives, removes virus files, and resets hidden file attributes (No Format)

### 📊 Logging & Export
- **Live Output** — Real-time color-coded log with step-by-step details
- **Export to Excel** — Save the full log as a formatted .xlsx file with colors
- **Export to Text** — Fallback export to .txt if Excel is not available

---

## 🖥️ GUI Highlights

| Feature | Details |
|---|---|
| 🌑 Dark Mode UI | Clean dark interface with `#0D0D0D` background and `#00C896` accent |
| ✅ Selective Steps | Choose exactly which steps to run via checkboxes |
| 📊 Progress Bar | Visual progress tracking with percentage |
| ⏱️ Timer | Live elapsed time counter |
| 🎨 Color-Coded Log | Green (success), Yellow (running), Red (error), Blue (section), Gray (skipped) |
| 📋 Export Log | Save results to Excel (.xlsx) or Text (.txt) |

---

## 🚀 Getting Started

### Prerequisites
- **Windows 10/11**
- **Python 3.8+** (for running from source)
- **Administrator privileges** (required for system operations)

### Option 1: Run from Source
```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/AK-System-Cleaner.git
cd AK-System-Cleaner

# Install dependencies
pip install openpyxl

# Run the cleaner
python AK_Cleaner_v4.3.py
