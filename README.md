# 🤖 Markus Voice Assistant

**Markus** is a fully offline, always-listening voice assistant designed for Windows 10/11. Built for privacy and speed, Markus uses **Vosk** for local speech recognition—meaning zero data ever leaves your machine.

---

## ✨ Key Features
*   🏠 **130+ Voice Commands**: Control your apps, folders, and windows with simple phrases.
*   🔒 **Fully Offline**: No API keys, no internet required, and no subscription fees.
*   🧠 **Smart Prefix Normalization**: Say **"Markus, [command]"** or just the command naturally—he'll understand.
*   🔊 **Precise Volume Control**: Commands like "lower the volume" are tuned for an exact 30% reduction.
*   🎬 **Media Focused**: Pre-mapped to prioritize your **:\Movies** directory and media controls.
*   💻 **Windows Master**: Launch any app, minimize windows, take screenshots, or even shutdown your PC.

---

## 🚀 Setup & Installation

### 1. Prerequisites
*   **Python 3.8+**
*   **Vosk Model**: Download the [vosk-model-small-en-us-0.15](https://alphacephei.com/vosk/models) and place it in a `models/` folder in the project directory.

### 2. Install Dependencies
```powershell
python -m venv venv
.\venv\Scripts\pip install -r requirements.txt
```

### 3. Start Markus
```powershell
# Run in foreground (to see output)
.\venv\Scripts\python.exe markus.py

# Run in background (Silent Mode)
.\start_markus.cmd
```

---

## 🎙️ Popular Commands

| Category | Commands Examples |
| :--- | :--- |
| **Media** | "play some music", "next song", "lower the volume (30%)", "unmute" |
| **Folders** | "let's watch some movies" (Opens D:\Movies), "open downloads", "disk D" |
| **System** | "take a screenshot", "lock the pc", "shutdown the pc", "system info" |
| **Apps** | "open chrome", "open spotify", "open discord", "open task manager" |
| **Productivity** | "study mode", "meeting mode", "morning routine", "focus mode" |

---

## 🛠️ Configuration
You can override default settings using environment variables:
*   `WAKE_PHRASE`: Change his name (Default: `Markus`)
*   `WAKE_DEVICE`: Select a specific microphone by name.
*   `WAKE_DEBUG`: Set to `1` to see everything Markus hears in the terminal.

---

## 📄 License
This project is open-source and free for all! 

**Created with ❤️ by nottherajyk**
