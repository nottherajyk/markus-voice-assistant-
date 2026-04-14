#!/usr/bin/env python3
"""Markus — Offline Voice-Controlled PC Assistant.

A fully offline, always-listening voice assistant for Windows 10/11.
Uses Vosk for local speech recognition and sounddevice for mic input.
Data-driven dispatch table architecture: one builder, one dict, one loop.
"""
from __future__ import annotations

import ctypes
import json
import os
import re
import subprocess
import sys
import time
import webbrowser
from pathlib import Path
from typing import Callable

# ---------------------------------------------------------------------------
# Vosk log suppression (before import)
# ---------------------------------------------------------------------------
os.environ.setdefault("VOSK_LOG_LEVEL", "-1")

import sounddevice as sd  # noqa: E402
from vosk import KaldiRecognizer, Model, SetLogLevel  # noqa: E402

SetLogLevel(-1)

# ===========================================================================
#                        CONFIGURATION (env-overridable)
# ===========================================================================
SAMPLE_RATE = int(os.getenv("WAKE_SAMPLE_RATE", "16000"))
BLOCK_DURATION = float(os.getenv("WAKE_BLOCK_DURATION", "0.25"))
TRIGGER_COOLDOWN = float(os.getenv("WAKE_TRIGGER_COOLDOWN", "2.0"))
DEVICE_NAME = os.getenv("WAKE_DEVICE")
DEBUG_RECOGNITION = bool(os.getenv("WAKE_DEBUG", ""))
WAKE_PHRASE = os.getenv("WAKE_PHRASE", "Markus")

MODEL_DIR = Path(__file__).resolve().parent / "models" / "vosk-model-small-en-us-0.15"

# User folders
HOME = Path.home()
PICTURES = HOME / "Pictures"
SCREENSHOTS = PICTURES / "Screenshots"
VIDEOS = HOME / "Videos"
CAPTURES = VIDEOS / "Captures"
DOWNLOADS = HOME / "Downloads"
DOCUMENTS = HOME / "Documents"
DESKTOP = HOME / "Desktop"
MUSIC_DIR = HOME / "Music"

# ===========================================================================
#                        APP PATH DETECTION
# ===========================================================================

def _find_exe(*candidates: str) -> str | None:
    """Return the first existing path from candidates or env override."""
    for c in candidates:
        p = Path(os.path.expandvars(c))
        if p.exists():
            return str(p)
    return None


CHROME_PATH = os.getenv("CHROME_APP_PATH") or _find_exe(
    r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
    r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
    r"%LocalAppData%\Google\Chrome\Application\chrome.exe",
) or "chrome.exe"  # fallback to PATH

BRAVE_PATH = os.getenv("BRAVE_APP_PATH") or _find_exe(
    r"%ProgramFiles%\BraveSoftware\Brave-Browser\Application\brave.exe",
    r"%LocalAppData%\BraveSoftware\Brave-Browser\Application\brave.exe",
) or "brave.exe"

SPOTIFY_PATH = os.getenv("SPOTIFY_APP_PATH") or _find_exe(
    r"%AppData%\Spotify\Spotify.exe",
    r"%LocalAppData%\Microsoft\WindowsApps\Spotify.exe",
) or "spotify.exe"

VSCODE_PATH = os.getenv("VSCODE_APP_PATH") or _find_exe(
    r"%LocalAppData%\Programs\Microsoft VS Code\Code.exe",
    r"%ProgramFiles%\Microsoft VS Code\Code.exe",
) or "code.exe"

TEAMS_PATH = os.getenv("TEAMS_APP_PATH") or _find_exe(
    r"%LocalAppData%\Microsoft\Teams\Update.exe",
    r"%ProgramFiles%\Microsoft\Teams\current\Teams.exe",
    r"%LocalAppData%\Microsoft\WindowsApps\ms-teams.exe",
) or "ms-teams.exe"

STEAM_PATH = os.getenv("STEAM_APP_PATH") or _find_exe(
    r"%ProgramFiles(x86)%\Steam\steam.exe",
    r"%ProgramFiles%\Steam\steam.exe",
) or "steam.exe"

GFE_PATH = os.getenv("GFE_APP_PATH") or _find_exe(
    r"%ProgramFiles%\NVIDIA Corporation\NVIDIA GeForce Experience\NVIDIA GeForce Experience.exe",
)

ANTIGRAVITY_PATH = os.getenv("ANTIGRAVITY_APP_PATH") or _find_exe(
    r"%LocalAppData%\Programs\antigravity\antigravity.exe",
)

CODEX_PATH = os.getenv("CODEX_APP_PATH") or _find_exe(
    r"%LocalAppData%\Programs\codex\codex.exe",
)


# ===========================================================================
#                        PHRASE DEFAULTS (Mapping Keys)
# ===========================================================================
_P: dict[str, str] = {
    "wake":             os.getenv("WAKE_UP_PHRASE",          "markus wake up"),
    "study":            os.getenv("STUDY_MODE_PHRASE",       "markus study mode"),
    "wind_down":        os.getenv("WIND_DOWN_PHRASE",        "markus wind down"),
    "watch":            os.getenv("WATCH_PHRASE",            "markus watch something"),
    "watch_movies":     os.getenv("WATCH_MOVIES_PHRASE",     "lets watch some movies"),
    "play_music":       os.getenv("PLAY_MUSIC_PHRASE",       "markus play some music"),
    "my_playlist":      os.getenv("MY_PLAYLIST_PHRASE",      "markus play my playlist"),
    "pause":            os.getenv("PAUSE_PHRASE",            "markus pause"),
    "play":             os.getenv("PLAY_PHRASE",             "markus play"),
    "next_track":       os.getenv("NEXT_TRACK_PHRASE",       "markus next song"),
    "prev_track":       os.getenv("PREV_TRACK_PHRASE",       "markus previous song"),
    "vol_up":           os.getenv("VOL_UP_PHRASE",           "markus volume up"),
    "vol_down":         os.getenv("VOL_DOWN_PHRASE",         "lower the volume"),
    "mute":             os.getenv("MUTE_PHRASE",             "markus mute"),
    "unmute":           os.getenv("UNMUTE_PHRASE",           "markus unmute"),
    "close_window":     os.getenv("CLOSE_WINDOW_PHRASE",     "markus close this"),
    "hide_all":         os.getenv("HIDE_ALL_PHRASE",         "markus hide everything"),
    "show_all":         os.getenv("SHOW_ALL_PHRASE",         "markus restore all"),
    "task_view":        os.getenv("TASK_VIEW_PHRASE",        "markus show my tasks"),
    "screenshot":       os.getenv("SCREENSHOT_PHRASE",       "markus take a screenshot"),
    "partial_ss":       os.getenv("PARTIAL_SS_PHRASE",       "markus partial screenshot"),
    "show_ss":          os.getenv("SHOW_SS_PHRASE",          "markus show screenshots"),
    "start_rec":        os.getenv("START_REC_PHRASE",        "markus start recording"),
    "show_rec":         os.getenv("SHOW_REC_PHRASE",         "markus show recordings"),
    "open_camera":      os.getenv("OPEN_CAMERA_PHRASE",      "markus open camera"),
    "lock":             os.getenv("LOCK_PHRASE",             "markus lock the pc"),
    "sleep":            os.getenv("SLEEP_PHRASE",            "markus go to sleep"),
    "open_files":       os.getenv("OPEN_FILES_PHRASE",       "markus open files"),
    "open_disk_d":      os.getenv("OPEN_DISK_D_PHRASE",      "markus open disk d"),
    "recycle_bin":      os.getenv("RECYCLE_BIN_PHRASE",      "markus open recycle bin"),
    "open_spotify":     os.getenv("OPEN_SPOTIFY_PHRASE",     "markus open spotify"),
    "open_teams":       os.getenv("OPEN_TEAMS_PHRASE",       "markus open teams"),
    "gaming":           os.getenv("GAMING_MODE_PHRASE",      "markus gaming mode"),
    "deactivate":       os.getenv("DEACTIVATE_PHRASE",       "markus deactivate"),
    # App launchers
    "open_chrome":      os.getenv("OPEN_CHROME_PHRASE",      "markus open chrome"),
    "open_notepad":     os.getenv("OPEN_NOTEPAD_PHRASE",     "markus open notepad"),
    "open_paint":       os.getenv("OPEN_PAINT_PHRASE",       "markus open paint"),
    "open_word":        os.getenv("OPEN_WORD_PHRASE",        "markus open word"),
    "open_excel":       os.getenv("OPEN_EXCEL_PHRASE",       "markus open excel"),
    "open_ppt":         os.getenv("OPEN_PPT_PHRASE",         "markus open powerpoint"),
    "open_discord":     os.getenv("OPEN_DISCORD_PHRASE",     "markus open discord"),
    "open_telegram":    os.getenv("OPEN_TELEGRAM_PHRASE",    "markus open telegram"),
    "open_whatsapp":    os.getenv("OPEN_WHATSAPP_PHRASE",    "markus open whatsapp"),
    "open_terminal":    os.getenv("OPEN_TERMINAL_PHRASE",    "markus open terminal"),
    "open_taskmgr":     os.getenv("OPEN_TASKMGR_PHRASE",    "markus open task manager"),
    "open_calc":        os.getenv("OPEN_CALC_PHRASE",       "markus open calculator"),
    "open_calendar":    os.getenv("OPEN_CALENDAR_PHRASE",   "markus open calendar"),
    "open_clock":       os.getenv("OPEN_CLOCK_PHRASE",      "markus open clock"),
    "open_maps":        os.getenv("OPEN_MAPS_PHRASE",       "markus open maps"),
    "open_weather":     os.getenv("OPEN_WEATHER_PHRASE",    "markus open weather"),
    "open_mail":        os.getenv("OPEN_MAIL_PHRASE",       "markus open mail"),
    "open_store":       os.getenv("OPEN_STORE_PHRASE",      "markus open store"),
    "restart":          os.getenv("RESTART_PHRASE",         "markus restart the pc"),
    "shutdown":         os.getenv("SHUTDOWN_PHRASE",        "markus shut down the pc"),
    "cancel_shutdown":  os.getenv("CANCEL_SHUTDOWN_PHRASE", "markus cancel shutdown"),
    "hibernate":        os.getenv("HIBERNATE_PHRASE",       "markus hibernate"),
    "logout":           os.getenv("LOGOUT_PHRASE",          "markus log out"),
    "schedule_shutdown":os.getenv("SCHEDULE_SHUTDOWN_PHRASE", "markus schedule shutdown"),
    "open_downloads":   os.getenv("OPEN_DOWNLOADS_PHRASE",  "markus open downloads"),
    "open_documents":   os.getenv("OPEN_DOCUMENTS_PHRASE",  "markus open documents"),
    "open_desktop":     os.getenv("OPEN_DESKTOP_PHRASE",    "markus open desktop"),
    "open_pictures":    os.getenv("OPEN_PICTURES_PHRASE",   "markus open pictures"),
    "open_music_dir":   os.getenv("OPEN_MUSIC_DIR_PHRASE",  "markus open my music"),
    "open_videos":      os.getenv("OPEN_VIDEOS_PHRASE",     "markus open my videos"),
    "open_temp":        os.getenv("OPEN_TEMP_PHRASE",       "markus open temp folder"),
    "open_startup":     os.getenv("OPEN_STARTUP_PHRASE",    "markus open startup folder"),
    "open_appdata":     os.getenv("OPEN_APPDATA_PHRASE",    "markus open appdata"),
    "brightness_up":    os.getenv("BRIGHTNESS_UP_PHRASE",   "markus increase brightness"),
    "brightness_down":  os.getenv("BRIGHTNESS_DOWN_PHRASE", "markus decrease brightness"),
    "night_light":      os.getenv("NIGHT_LIGHT_PHRASE",     "markus toggle night light"),
    "display_settings": os.getenv("DISPLAY_SETTINGS_PHRASE","markus open display settings"),
    "project_screen":   os.getenv("PROJECT_SCREEN_PHRASE",  "markus project screen"),
    "toggle_wifi":      os.getenv("TOGGLE_WIFI_PHRASE",     "markus toggle wifi"),
    "bluetooth":        os.getenv("BLUETOOTH_PHRASE",       "markus open bluetooth settings"),
    "network_settings": os.getenv("NETWORK_SETTINGS_PHRASE","markus network settings"),
    "airplane_mode":    os.getenv("AIRPLANE_MODE_PHRASE",   "markus toggle airplane mode"),
    "vpn_settings":     os.getenv("VPN_SETTINGS_PHRASE",    "markus vpn settings"),
    "hotspot":          os.getenv("HOTSPOT_PHRASE",         "markus toggle hotspot"),
    "speed_test":       os.getenv("SPEED_TEST_PHRASE",      "markus run speed test"),
    "open_settings":    os.getenv("OPEN_SETTINGS_PHRASE",   "markus open settings"),
    "system_info":      os.getenv("SYSTEM_INFO_PHRASE",     "markus system info"),
    "storage_settings": os.getenv("STORAGE_SETTINGS_PHRASE","markus storage settings"),
    "device_manager":   os.getenv("DEVICE_MANAGER_PHRASE",  "markus device manager"),
    "disk_cleanup":     os.getenv("DISK_CLEANUP_PHRASE",    "markus disk cleanup"),
    "windows_update":   os.getenv("WINDOWS_UPDATE_PHRASE",  "markus check for updates"),
    "sound_settings":   os.getenv("SOUND_SETTINGS_PHRASE",  "markus sound settings"),
    "battery_info":     os.getenv("BATTERY_INFO_PHRASE",    "markus battery settings"),
    "date_time":        os.getenv("DATE_TIME_PHRASE",       "markus open date and time"),
    "privacy_settings": os.getenv("PRIVACY_SETTINGS_PHRASE","markus privacy settings"),
    "clear_clipboard":  os.getenv("CLEAR_CLIPBOARD_PHRASE", "markus clear clipboard"),
    "empty_trash":      os.getenv("EMPTY_TRASH_PHRASE",     "markus empty recycle bin"),
    "flush_dns":        os.getenv("FLUSH_DNS_PHRASE",       "markus flush dns"),
    "emoji":            os.getenv("EMOJI_PHRASE",           "markus show emojis"),
    "clipboard_hist":   os.getenv("CLIPBOARD_HIST_PHRASE",  "markus clipboard history"),
    "snip":             os.getenv("SNIP_PHRASE",            "markus snipping tool"),
    "notifications":    os.getenv("NOTIFICATIONS_PHRASE",   "markus show notifications"),
    "quick_settings":   os.getenv("QUICK_SETTINGS_PHRASE",  "markus quick settings"),
    "widgets":          os.getenv("WIDGETS_PHRASE",         "markus widgets"),
    "magnifier_on":     os.getenv("MAGNIFIER_ON_PHRASE",    "markus turn on magnifier"),
    "magnifier_off":    os.getenv("MAGNIFIER_OFF_PHRASE",   "markus turn off magnifier"),
    "undo":             os.getenv("UNDO_PHRASE",            "markus undo"),
    "redo":             os.getenv("REDO_PHRASE",            "markus redo"),
    "select_all":       os.getenv("SELECT_ALL_PHRASE",      "markus select all"),
    "copy":             os.getenv("COPY_PHRASE",            "markus copy"),
    "paste":            os.getenv("PASTE_PHRASE",           "markus paste"),
    "cut":              os.getenv("CUT_PHRASE",             "markus cut"),
    "find":             os.getenv("FIND_PHRASE",            "markus search"),
    "save":             os.getenv("SAVE_PHRASE",            "markus save"),
    "new_tab":          os.getenv("NEW_TAB_PHRASE",         "markus new tab"),
    "close_tab":        os.getenv("CLOSE_TAB_PHRASE",       "markus close tab"),
    "refresh":          os.getenv("REFRESH_PHRASE",         "markus refresh"),
    "delete":           os.getenv("DELETE_PHRASE",          "markus delete"),
    "narrator":         os.getenv("NARRATOR_PHRASE",        "markus narrate this"),
    "high_contrast":    os.getenv("HIGH_CONTRAST_PHRASE",   "markus toggle high contrast"),
    "text_size":        os.getenv("TEXT_SIZE_PHRASE",       "markus increase text size"),
    "color_filters":    os.getenv("COLOR_FILTERS_PHRASE",    "markus color filters"),
    "osk":              os.getenv("OSK_PHRASE",             "markus on screen keyboard"),
    "morning":          os.getenv("MORNING_PHRASE",         "markus good morning"),
    "focus_mode":       os.getenv("FOCUS_MODE_PHRASE",      "markus focus mode"),
    "end_focus":        os.getenv("END_FOCUS_PHRASE",       "markus end focus"),
    "presentation":     os.getenv("PRESENTATION_PHRASE",    "markus presentation mode"),
    "break_time":       os.getenv("BREAK_TIME_PHRASE",      "markus take a break"),
    "night_mode":       os.getenv("NIGHT_MODE_PHRASE",      "markus night mode"),
    "code_mode":        os.getenv("CODE_MODE_PHRASE",       "markus codex mode"),
    "meeting_mode":     os.getenv("MEETING_MODE_PHRASE",     "markus meeting mode"),
    "clean_up":         os.getenv("CLEAN_UP_PHRASE",         "markus clean up"),
    # Info / TTS
    "what_time":        os.getenv("WHAT_TIME_PHRASE",        "markus what time is it"),
    "what_day":         os.getenv("WHAT_DAY_PHRASE",         "markus what day is it"),
    "battery_level":    os.getenv("BATTERY_LEVEL_PHRASE",    "markus battery level"),
    "my_ip":            os.getenv("MY_IP_PHRASE",            "markus what is my ip"),
}

# ===========================================================================
#                        VARIANT EXTRAS
# ===========================================================================
VARIANT_EXTRAS: dict[str, list[str]] = {
    "wake":             ["markus wake up", "wake up"],
    "study":            ["markus let's study", "markus work mode", "markus study mode", "lets do some work"],
    "wind_down":        ["markus wind down", "call it a day", "markus shut everything down"],
    "watch":            ["markus watch something", "lets watch something"],
    "watch_movies":     ["markus movies", "watch some movies", "lets watch some movies"],
    "play_music":       ["play some music", "markus music"],
    "my_playlist":      ["play my playlist", "markus play my playlist"],
    "pause":            ["pause", "markus stop"],
    "play":             ["play", "markus resume"],
    "next_track":       ["next song", "skip this song", "change the song", "markus next", "markus skip"],
    "prev_track":       ["previous song", "markus previous", "repeat the last song"],
    "vol_up":           ["louder", "increase the volume", "markus volume up", "turn it up"],
    "vol_down":         ["lower the volume", "decrease the volume", "markus volume down", "turn it down"],
    "mute":             ["mute", "silence", "shh"],
    "unmute":           ["unmute", "umute", "volume on", "mute off"],
    "close_window":     ["close it", "close this window", "close application", "shut it down", "markus close this"],
    "hide_all":         ["minimize all", "show desktop", "hide everything", "markus minimize"],
    "show_all":         ["restore all", "show me everything", "markus restore"],
    "task_view":        ["task view", "show my tasks", "markus tasks"],
    "screenshot":       ["take a screenshot", "full screenshot", "markus screenshot"],
    "partial_ss":       ["partial screenshot", "snip screen", "markus snip screen"],
    "show_ss":          ["show my screenshots", "show screenshots"],
    "start_rec":        ["start recording", "record screen", "markus record"],
    "show_rec":         ["show me the recording", "show recordings", "markus recordings"],
    "open_camera":      ["open the camera", "markus camera"],
    "lock":             ["protect yourself", "markus lock", "lock the pc"],
    "sleep":            ["go to sleep", "put the pc to sleep", "markus sleep"],
    "open_files":       ["files", "open explorer", "markus explorer"],
    "open_disk_d":      ["open my disk", "disk d", "markus disk"],
    "recycle_bin":      ["open recycle bin", "trash bin", "markus trash", "markus recycle bin"],
    "open_spotify":     ["open spotify"],
    "open_teams":       ["open teams"],
    "gaming":           ["lets do some gaming", "markus gaming", "gaming mode"],
    "deactivate":       ["deactivate yourself", "markus shut down", "markus exit", "markus quit"],
    # App launchers
    "open_chrome":      ["open chrome"],
    "open_notepad":     ["open notepad"],
    "open_paint":       ["open paint"],
    "open_word":        ["open word"],
    "open_excel":       ["open excel"],
    "open_ppt":         ["open powerpoint"],
    "open_discord":     ["open discord"],
    "open_telegram":    ["open telegram"],
    "open_whatsapp":    ["open whatsapp"],
    "open_terminal":    ["open terminal", "open powershell", "open cmd"],
    "open_taskmgr":     ["open task manager", "show running processes"],
    "open_calc":        ["calculator", "markus calc"],
    "open_downloads":   ["downloads"],
    "delete":           ["delete that", "remove it", "markus delete"],
}


# ===========================================================================
#                          HELPER UTILITIES
# ===========================================================================

def _ps(command: str) -> None:
    """Execute a PowerShell command silently."""
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            check=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except subprocess.CalledProcessError as e:
        print(f"PS Error: {e}", file=sys.stderr)

def _keybd(*keys: int) -> None:
    """Simulate keybd_event sequence via PowerShell (Robust)."""
    lines = []
    lines.append(
        "Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; "
        "public class Win32 { [DllImport(\"user32.dll\")] public static extern void "
        "keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo); }'"
    )
    # Key Down
    for k in keys:
        lines.append(f"[Win32.KBD]::keybd_event(0x{k:02X}, 0, 0, [UIntPtr]::Zero)")
    # Key Up (Reverse)
    for k in reversed(keys):
        lines.append(f"[Win32.KBD]::keybd_event(0x{k:02X}, 0, 2, [UIntPtr]::Zero)")
    _ps("; ".join(lines))

# ---------------------------------------------------------------------------
#                           ACTIONS ENGINE
# ---------------------------------------------------------------------------

def _open(target: str) -> None:
    """Silent os.startfile wrapper."""
    try:
        os.startfile(target)
    except OSError as e:
        print(f"Cannot open {target}: {e}", file=sys.stderr)

def _open_url(url: str) -> None:
    """Open a URL in the default browser."""
    webbrowser.open(url)

# App Launchers
def open_chrome() -> None:
    if CHROME_PATH: _open(CHROME_PATH)

def open_notepad() -> None:
    _open("notepad.exe")

def open_paint() -> None:
    _open("mspaint.exe")

def open_word() -> None:
    _open_url("ms-word:")

def open_excel() -> None:
    _open_url("ms-excel:")

def open_ppt() -> None:
    _open_url("ms-powerpoint:")

def open_discord() -> None:
    if DISCORD_PATH: _open(DISCORD_PATH)

def open_telegram() -> None:
    _open(r"%AppData%\Telegram Desktop\Telegram.exe")

def open_whatsapp() -> None:
    _open("whatsapp:")

def open_terminal() -> None:
    _open("wt.exe")

def open_taskmgr() -> None:
    _ps("Start-Process taskmgr.exe")

def open_calc() -> None:
    _open("calc.exe")

def open_calendar() -> None:
    _open("outlookcal:")

def open_clock() -> None:
    _open("ms-clock:")

def open_maps() -> None:
    _open("bingmaps:")

def open_weather() -> None:
    _open("bingweather:")

def open_mail() -> None:
    _open("outlookmail:")

def open_store() -> None:
    _open("ms-windows-store:")

# Folder Navigation
def open_downloads() -> None:
    _open("shell:Downloads")

def open_documents() -> None:
    _open("shell:Personal")

def open_desktop() -> None:
    _open("shell:Desktop")

def open_pictures() -> None:
    _open("shell:My Pictures")

def open_music_dir() -> None:
    _open("shell:My Music")

def open_videos() -> None:
    _open("shell:My Video")

def open_temp() -> None:
    _open(os.environ.get("TEMP", r"C:\Windows\Temp"))

def open_startup() -> None:
    _open("shell:Startup")

def open_appdata() -> None:
    _open(os.environ.get("APPDATA", ""))

def watch_movies() -> None:
    """Open the local movies folder."""
    candidates = [
        r"D:\Movies",
        os.path.join(os.environ.get("USERPROFILE", ""), "Videos"),
        r"D:\Videos",
        r"E:\Videos",
    ]
    for c in candidates:
        if os.path.isdir(c):
            _open(c)
            return
    _open("shell:My Video")

def _tts(text: str) -> None:
    """Speak text using Windows built-in Speech Synthesizer."""
    # Use double quotes and escape existing double quotes for PS safety
    safe_text = text.replace('"', '`"')
    _ps(
        "Add-Type -AssemblyName System.Speech; "
        f"$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; $s.Speak(\"{safe_text}\")"
    )

# ===========================================================================
#                        COMMAND ACTIONS
# ===========================================================================

def wake_up() -> None:
    """Open Brave, Antigravity, Codex."""
    if BRAVE_PATH:
        _open(BRAVE_PATH)
    if ANTIGRAVITY_PATH:
        _open(ANTIGRAVITY_PATH)
    if CODEX_PATH:
        _open(CODEX_PATH)

def study_mode() -> None:
    if VSCODE_PATH: _open(VSCODE_PATH)
    if CHROME_PATH: _open(CHROME_PATH)
    _ps("Start-Process 'https://music.youtube.com' -WindowStyle Minimized")

def wind_down() -> None:
    _tts("Markus is signing off. Closing windows and checking the day's logs.")
    _keybd(0x5B, 0x44)  # Win + D (Minimize all)
    _keybd(0xB3)  # Play/Pause

def watch_something() -> None:
    urls = ["https://net22.cc/home", "https://cinehd.cc/home", "https://www.youtube.com/"]
    for url in urls:
        _open_url(url)

def play_music() -> None:
    _open_url("https://music.youtube.com/search?q=lofi")

def my_playlist() -> None:
    _open_url("https://music.youtube.com/playlist?list=LM")

def media_play_pause() -> None: _keybd(0xB3)
def media_next() -> None:       _keybd(0xB0)
def media_prev() -> None:       _keybd(0xB1)
def volume_up() -> None:        _keybd(0xAF)

def volume_down() -> None:
    # 15 pulses of volume down = roughly 30% reduction (Windows uses increments of 2)
    for _ in range(15):
        _keybd(0xAE)

def mute() -> None:             _keybd(0xAD)

def unmute() -> None:
    # Sending 'Volume Up' reliably unmutes without toggling.
    _keybd(0xAF)

def close_current_window() -> None: _keybd(0x12, 0x73)  # Alt + F4
def hide_all_windows() -> None:     _keybd(0x5B, 0x44)  # Win + D
def show_all_windows() -> None:     _keybd(0x5B, 0x44)  # Win + D (Toggle)
def task_view() -> None:            _keybd(0x5B, 0x09)  # Win + Tab

def take_screenshot() -> None:      _keybd(0x5B, 0x2C)  # Win + PrtSc
def partial_screenshot() -> None:   _keybd(0x5B, 0x10, 0x53)  # Win + Shift + S

def show_screenshots() -> None:
    if SCREENSHOTS.exists(): _open(str(SCREENSHOTS))

def start_recording() -> None:      _keybd(0x5B, 0x12, 0x52)  # Win + Alt + R

def show_recordings() -> None:
    if CAPTURES.exists(): _open(str(CAPTURES))

def open_camera() -> None:          _ps("Start-Process 'microsoft.windows.camera:'")

def lock_pc() -> None:              _ps("rundll32.exe user32.dll,LockWorkStation")

def sleep_pc() -> None:             _ps("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def open_files() -> None:           _open("explorer.exe")
def open_disk_d() -> None:          _open("D:\\")
def open_recycle_bin() -> None:     _open("shell:RecycleBinFolder")
def open_spotify() -> None:         _open(SPOTIFY_PATH) if SPOTIFY_PATH else None
def open_teams() -> None:           _open(TEAMS_PATH) if TEAMS_PATH else None

def gaming_mode() -> None:
    if STEAM_PATH: _open(STEAM_PATH)
    if GFE_PATH: _open(GFE_PATH)
    _keybd(0x5B, 0x44)

def deactivate() -> None:
    _tts("Deactivating. Goodbye.")
    sys.exit(0)

def restart_pc() -> None:    _ps("Restart-Computer -Force")
def shutdown_pc() -> None:   _ps("Stop-Computer -Force")
def cancel_shutdown() -> None: _ps("shutdown -a")
def hibernate_pc() -> None: _ps("shutdown -h")
def logout_pc() -> None: _ps("logoff")
def schedule_shutdown() -> None: _ps("shutdown -s -t 3600")

def brightness_up() -> None:   _ps("(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness + 10)")
def brightness_down() -> None: _ps("(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness - 10)")
def night_light() -> None:    _ps("Start-Process 'ms-settings:nightlight'")
def display_settings() -> None: _ps("Start-Process 'ms-settings:display'")
def project_screen() -> None:  _keybd(0x5B, 0x50)  # Win + P

def toggle_wifi() -> None:    _ps("if ((Get-NetAdapter -Name 'Wi-Fi').Status -eq 'Up') { Disable-NetAdapter -Name 'Wi-Fi' -Confirm:$false } else { Enable-NetAdapter -Name 'Wi-Fi' -Confirm:$false }")
def bluetooth_settings() -> None: _ps("Start-Process 'ms-settings:bluetooth'")
def network_settings() -> None:   _ps("Start-Process 'ms-settings:network'")
def airplane_mode() -> None:      _ps("Start-Process 'ms-settings:network-airplanemode'")
def vpn_settings() -> None:       _ps("Start-Process 'ms-settings:network-vpn'")
def hotspot_settings() -> None:   _ps("Start-Process 'ms-settings:network-mobilehotspot'")
def speed_test() -> None:         _open_url("https://www.google.com/search?q=speedtest")

def open_settings() -> None:    _ps("Start-Process 'ms-settings:'")
def system_info() -> None:      _ps("msinfo32.exe")
def storage_settings() -> None: _ps("Start-Process 'ms-settings:storagesense'")
def device_manager() -> None:   _ps("devmgmt.msc")
def disk_cleanup() -> None:     _ps("cleanmgr.exe")
def windows_update() -> None:   _ps("Start-Process 'ms-settings:windowsupdate'")
def sound_settings() -> None:   _ps("Start-Process 'ms-settings:sound'")

def battery_info() -> None:     _ps("powercfg /batteryreport; Start-Process 'battery-report.html'")
def date_time() -> None:        _ps("Start-Process 'ms-settings:dateandtime'")
def privacy_settings() -> None: _ps("Start-Process 'ms-settings:privacy'")

def clear_clipboard() -> None:  _ps("Clear-Clipboard")
def empty_trash() -> None:      _ps("$rb = New-Object -ComObject Shell.Application; $rb.NameSpace(0x0a).Items() | ForEach-Object { Remove-Item $_.Path -Recurse -Force }")
def flush_dns() -> None:        _ps("ipconfig /flushdns")

def emoji_picker() -> None:     _keybd(0x5B, 0xBE)  # Win + .
def clipboard_history() -> None: _keybd(0x5B, 0x56)  # Win + V
def snipping_tool() -> None:    _ps("Start-Process 'snippingtool.exe'")
def action_center() -> None:    _keybd(0x5B, 0x41)  # Win + A
def quick_settings() -> None:   _keybd(0x5B, 0x41)  # Duplicate Win+A for Win11
def widgets_panel() -> None:    _keybd(0x5B, 0x57)  # Win + W

def magnifier_on() -> None:     _keybd(0x5B, 0xBB)  # Win + Plus
def magnifier_off() -> None:    _keybd(0x5B, 0x1B)  # Win + Esc

def do_undo() -> None:       _keybd(0x11, 0x5A)  # Ctrl + Z
def do_redo() -> None:       _keybd(0x11, 0x59)  # Ctrl + Y
def do_select_all() -> None: _keybd(0x11, 0x41)  # Ctrl + A
def do_copy() -> None:       _keybd(0x11, 0x43)  # Ctrl + C
def do_paste() -> None:      _keybd(0x11, 0x56)  # Ctrl + V
def do_cut() -> None:        _keybd(0x11, 0x58)  # Ctrl + X
def do_find() -> None:       _keybd(0x11, 0x46)  # Ctrl + F
def do_save() -> None:       _keybd(0x11, 0x53)  # Ctrl + S
def do_new_tab() -> None:    _keybd(0x11, 0x54)  # Ctrl + T
def do_close_tab() -> None:  _keybd(0x11, 0x57)  # Ctrl + W
def do_refresh() -> None:    _keybd(0x74)        # F5
def do_delete() -> None:     _keybd(0x2E)        # Delete key

def narrator() -> None:         _keybd(0x5B, 0x11, 0x0D)  # Win + Ctrl + Enter
def high_contrast() -> None:    _ps("Start-Process 'ms-settings:easeofaccess-highcontrast'")
def text_size() -> None:        _ps("Start-Process 'ms-settings:easeofaccess-display'")
def color_filters() -> None:    _ps("Start-Process 'ms-settings:easeofaccess-colorfilter'")
def on_screen_keyboard() -> None: _ps("osk.exe")

def morning_routine() -> None:
    _tts("Good morning! Checking your workspace.")
    wake_up()
    system_info()
    time.sleep(1)
    _open_url("https://calendar.google.com")

def focus_mode() -> None:
    _tts("Focus mode engaged. Silencing notifications.")
    _ps("Start-Process 'ms-settings:quiethours'")
    _keybd(0xB3)  # Pause music
    time.sleep(1)
    _open_url("https://music.youtube.com/search?q=focus+music")

def end_focus() -> None:
    _tts("Focus mode deactivated. Welcome back.")
    _ps("Start-Process 'ms-settings:quiethours'")

def presentation_mode() -> None:
    _tts("Presentation mode. Cleaning desktop.")
    hide_all_windows()
    _ps(r"Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'HideIcons' -Value 1")
    _ps(r"Stop-Process -Name explorer -Force") # Refresh icons

def break_time() -> None:
    _tts("Time for a break. Screen will lock in 5 seconds.")
    _ps(r"Start-Process 'https://www.youtube.com/watch?v=5qap5aO4i9A'") # Lofi
    time.sleep(5)
    lock_pc()

def night_mode_action() -> None:
    brightness_down()
    _ps(r"Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\UI\Colors' -Name 'SystemUsesLightTheme' -Value 0")
    _tts("Night mode active.")

def code_mode() -> None:
    if VSCODE_PATH: _open(VSCODE_PATH)
    if CHROME_PATH: _open(CHROME_PATH)
    time.sleep(1)
    _open_url("https://open.spotify.com/playlist/0vvXsWCC9xrXsKd4FyS8kM")

def meeting_mode() -> None:
    if TEAMS_PATH: _open(TEAMS_PATH)
    time.sleep(0.5)
    _keybd(0x5B, 0x44)  # Minimize all
    _keybd(0xB3)  # Pause spotify

def clean_up() -> None:
    wind_down()
    time.sleep(1)
    clear_clipboard()
    empty_trash()


# Info / TTS
def tell_time() -> None:
    _ps(
        "Add-Type -AssemblyName System.Speech; $s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
        "$s.Speak('The time is ' + (Get-Date).ToString('h:mm tt'))"
    )

def tell_day() -> None:
    _ps(
        "Add-Type -AssemblyName System.Speech; $s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
        "$s.Speak('Today is ' + (Get-Date).DayOfWeek)"
    )

def tell_battery() -> None:
    _ps(
        "$b = Get-CimInstance -ClassName Win32_Battery; $p = $b.EstimatedChargeRemaining; "
        "Add-Type -AssemblyName System.Speech; $s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
        "$s.Speak('Battery is at ' + $p + ' percent')"
    )

def tell_ip() -> None:
    _ps(
        "$ip = (Invoke-WebRequest -uri 'https://api.ipify.org').Content; "
        "Add-Type -AssemblyName System.Speech; $s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
        "$s.Speak('Your external IP is ' + $ip)"
    )


# ===========================================================================
#                          LOGIC & WAKE ENGINE
# ===========================================================================

def _normalize(text: str) -> str:
    """Standardize voice input for matching."""
    txt = text.lower().strip()
    # Strip common punctuation
    txt = re.sub(r"[^\w\s]", "", txt)
    # Smart Wake-Word Stripping:
    # If the user says "Markus, [command]", we want to match just "[command]".
    wake = WAKE_PHRASE.lower()
    if txt.startswith(wake):
        # Remove wake word and handle potential comma/space (e.g., "markus open files")
        txt = txt[len(wake):].strip()
    return txt

def _build_variant_sets() -> dict[str, set[str]]:
    """Merge primary phrases with variants into sets."""
    final: dict[str, set[str]] = {}
    for key, primary in _P.items():
        variants = {primary.lower()}
        if key in VARIANT_EXTRAS:
            for v in VARIANT_EXTRAS[key]:
                variants.add(v.lower())
        # Add stripped versions of everything (in case normalization strips Markus)
        stripped = set()
        for v in variants:
            s = _normalize(v)
            if s: stripped.add(s)
        variants.update(stripped)
        final[key] = variants
    return final

def _collect_grammar_words(variant_sets: dict[str, set[str]]) -> list[str]:
    """Extract unique words for the Vosk grammar constraint."""
    words: set[str] = set()
    for phrases in variant_sets.values():
        for phrase in phrases:
            words.update(phrase.split())
    return sorted(words)


# ===========================================================================
#                          MAIN LOOP
# ===========================================================================
def main() -> None:
    """Entry point — build dispatch table, start listening."""
    # --- Parse CLI flags ---
    debug = DEBUG_RECOGNITION or "--debug-recognition" in sys.argv

    # --- Load model ---
    if not MODEL_DIR.exists():
        print(
            f"ERROR: Vosk model not found at {MODEL_DIR}\n"
            "Download vosk-model-small-en-us-0.15 from "
            "https://alphacephei.com/vosk/models and extract into models/",
            file=sys.stderr,
        )
        sys.exit(1)

    print("Loading Vosk model...", flush=True)
    model = Model(str(MODEL_DIR))

    # --- Build variants & dispatch ---
    vs = _build_variant_sets()

    # Map command keys → (label, action)
    CMD_ACTIONS: dict[str, tuple[str, Callable[[], None]]] = {
        "wake":             ("Wake Up",             wake_up),
        "study":            ("Study Mode",          study_mode),
        "wind_down":        ("Wind Down",           wind_down),
        "watch":            ("Watch Something",     watch_something),
        "watch_movies":     ("Watch Movies",        watch_movies),
        "play_music":       ("Play Music",          play_music),
        "my_playlist":      ("My Playlist",         my_playlist),
        "pause":            ("Pause",               media_play_pause),
        "play":             ("Play",                media_play_pause),
        "next_track":       ("Next Track",          media_next),
        "prev_track":       ("Previous Track",      media_prev),
        "vol_up":           ("Volume Up",           volume_up),
        "vol_down":         ("Volume Down",         volume_down),
        "mute":             ("Mute",                mute),
        "unmute":           ("Unmute",              unmute),
        "close_window":     ("Close Window",        close_current_window),
        "hide_all":         ("Hide Everything",     hide_all_windows),
        "show_all":         ("Show Everything",     show_all_windows),
        "task_view":        ("Task View",           task_view),
        "screenshot":       ("Screenshot",          take_screenshot),
        "partial_ss":       ("Partial Screenshot",  partial_screenshot),
        "show_ss":          ("Show Screenshots",    show_screenshots),
        "start_rec":        ("Start Recording",     start_recording),
        "show_rec":         ("Show Recordings",     show_recordings),
        "open_camera":      ("Open Camera",         open_camera),
        "lock":             ("Lock PC",             lock_pc),
        "sleep":            ("Sleep",               sleep_pc),
        "open_files":       ("Open Files",          open_files),
        "open_disk_d":      ("Open Disk D",         open_disk_d),
        "recycle_bin":      ("Recycle Bin",          open_recycle_bin),
        "open_spotify":     ("Open Spotify",        open_spotify),
        "open_teams":       ("Open Teams",          open_teams),
        "gaming":           ("Gaming Mode",         gaming_mode),
        "deactivate":       ("Deactivate",          deactivate),
        "open_chrome":      ("Open Chrome",         open_chrome),
        "open_notepad":     ("Open Notepad",        open_notepad),
        "open_paint":       ("Open Paint",          open_paint),
        "open_word":        ("Open Word",           open_word),
        "open_excel":       ("Open Excel",          open_excel),
        "open_ppt":         ("Open PowerPoint",     open_ppt),
        "open_discord":     ("Open Discord",        open_discord),
        "open_telegram":    ("Open Telegram",       open_telegram),
        "open_whatsapp":    ("Open WhatsApp",       open_whatsapp),
        "open_terminal":    ("Open Terminal",       open_terminal),
        "open_taskmgr":     ("Task Manager",        open_taskmgr),
        "open_calc":        ("Calculator",          open_calc),
        "open_calendar":    ("Calendar",            open_calendar),
        "open_clock":       ("Clock",               open_clock),
        "open_maps":        ("Maps",                open_maps),
        "open_weather":     ("Weather",             open_weather),
        "open_mail":        ("Mail",                open_mail),
        "open_store":       ("Store",               open_store),
        "restart":          ("Restart PC",          restart_pc),
        "shutdown":         ("Shutdown PC",         shutdown_pc),
        "cancel_shutdown":  ("Cancel Shutdown",     cancel_shutdown),
        "hibernate":        ("Hibernate",           hibernate_pc),
        "logout":           ("Log Out",             logout_pc),
        "schedule_shutdown":("Schedule Shutdown",   schedule_shutdown),
        "open_downloads":   ("Open Downloads",      open_downloads),
        "open_documents":   ("Open Documents",      open_documents),
        "open_desktop":     ("Open Desktop",        open_desktop),
        "open_pictures":    ("Open Pictures",       open_pictures),
        "open_music_dir":   ("Open Music Folder",   open_music_dir),
        "open_videos":      ("Open Videos",         open_videos),
        "open_temp":        ("Open Temp",           open_temp),
        "open_startup":     ("Open Startup",        open_startup),
        "open_appdata":     ("Open AppData",        open_appdata),
        "brightness_up":    ("Brightness Up",       brightness_up),
        "brightness_down":  ("Brightness Down",     brightness_down),
        "night_light":      ("Night Light",         night_light),
        "display_settings": ("Display Settings",    display_settings),
        "project_screen":   ("Project Screen",      project_screen),
        "toggle_wifi":      ("Toggle Wi-Fi",        toggle_wifi),
        "bluetooth":        ("Bluetooth",           bluetooth_settings),
        "network_settings": ("Network Settings",    network_settings),
        "airplane_mode":    ("Airplane Mode",       airplane_mode),
        "vpn_settings":     ("VPN Settings",        vpn_settings),
        "hotspot":          ("Hotspot",             hotspot_settings),
        "speed_test":       ("Speed Test",          speed_test),
        "open_settings":    ("Open Settings",       open_settings),
        "system_info":      ("System Info",         system_info),
        "storage_settings": ("Storage Settings",    storage_settings),
        "device_manager":   ("Device Manager",      device_manager),
        "disk_cleanup":     ("Disk Cleanup",        disk_cleanup),
        "windows_update":   ("Windows Update",      windows_update),
        "sound_settings":   ("Sound Settings",      sound_settings),
        "battery_info":     ("Battery Info",        battery_info),
        "date_time":        ("Date & Time",         date_time),
        "privacy_settings": ("Privacy Settings",    privacy_settings),
        "clear_clipboard":  ("Clear Clipboard",     clear_clipboard),
        "empty_trash":      ("Empty Trash",         empty_trash),
        "flush_dns":        ("Flush DNS",           flush_dns),
        "emoji":            ("Emoji Picker",        emoji_picker),
        "clipboard_hist":   ("Clipboard History",   clipboard_history),
        "snip":             ("Snipping Tool",       snipping_tool),
        "notifications":    ("Notifications",       action_center),
        "quick_settings":   ("Quick Settings",      quick_settings),
        "widgets":          ("Widgets",             widgets_panel),
        "magnifier_on":     ("Magnifier On",        magnifier_on),
        "magnifier_off":    ("Magnifier Off",       magnifier_off),
        "undo":             ("Undo",                do_undo),
        "redo":             ("Redo",                do_redo),
        "select_all":       ("Select All",          do_select_all),
        "copy":             ("Copy",                do_copy),
        "paste":            ("Paste",               do_paste),
        "cut":              ("Cut",                 do_cut),
        "find":             ("Find",                do_find),
        "save":             ("Save",                do_save),
        "new_tab":          ("New Tab",             do_new_tab),
        "close_tab":        ("Close Tab",           do_close_tab),
        "refresh":          ("Refresh",             do_refresh),
        "delete":           ("Delete",              do_delete),
        "narrator":         ("Narrator",            narrator),
        "high_contrast":    ("High Contrast",       high_contrast),
        "text_size":        ("Text Size",           text_size),
        "color_filters":    ("Color Filters",       color_filters),
        "osk":              ("On-Screen Keyboard",  on_screen_keyboard),
        "morning":          ("Morning Routine",     morning_routine),
        "focus_mode":       ("Focus Mode",          focus_mode),
        "end_focus":        ("End Focus",           end_focus),
        "presentation":     ("Presentation Mode",   presentation_mode),
        "break_time":       ("Break Time",          break_time),
        "night_mode":       ("Night Mode",          night_mode_action),
        "code_mode":        ("Code Mode",           code_mode),
        "meeting_mode":     ("Meeting Mode",        meeting_mode),
        "clean_up":         ("Clean Up",            clean_up),
        "what_time":        ("Tell Time",           tell_time),
        "what_day":         ("Tell Day",            tell_day),
        "battery_level":    ("Battery Level",       tell_battery),
        "my_ip":            ("My IP",               tell_ip),
    }

    # Build the final dispatch table: list[(variant_set, label, action)]
    dispatch_table: list[tuple[set[str], str, Callable[[], None]]] = []
    for key, phrase_set in vs.items():
        if key in CMD_ACTIONS:
            label, action = CMD_ACTIONS[key]
            dispatch_table.append((phrase_set, label, action))

    # --- Build Vosk grammar ---
    grammar_words = _collect_grammar_words(vs)
    grammar_json = json.dumps(grammar_words)

    rec = KaldiRecognizer(model, SAMPLE_RATE, grammar_json)
    rec.SetWords(True)

    # --- Resolve audio device ---
    device_id = None
    if DEVICE_NAME:
        devices = sd.query_devices()
        for i, d in enumerate(devices):
            if DEVICE_NAME.lower() in d["name"].lower() and d["max_input_channels"] > 0:
                device_id = i
                break

    block_size = int(SAMPLE_RATE * BLOCK_DURATION)
    last_trigger_time = 0.0

    print(f"Markus is listening ({len(dispatch_table)} commands loaded)...", flush=True)
    if debug:
        print(f"  Grammar: {len(grammar_words)} words", flush=True)

    # --- Audio stream ---
    try:
        with sd.RawInputStream(
            samplerate=SAMPLE_RATE,
            blocksize=block_size,
            dtype="int16",
            channels=1,
            device=device_id,
        ) as stream:
            while True:
                data, _ = stream.read(block_size)
                if rec.AcceptWaveform(bytes(data)):
                    result = json.loads(rec.Result())
                    text = result.get("text", "")
                    normalized = _normalize(text)
                    if not normalized:
                        continue
                    if debug:
                        print(f"  [heard] \"{normalized}\"", flush=True)

                    for cmd_set, label, action in dispatch_table:
                        if normalized in cmd_set:
                            now = time.monotonic()
                            if now - last_trigger_time >= TRIGGER_COOLDOWN:
                                last_trigger_time = now
                                print(f"{label} detected.", flush=True)
                                try:
                                    action()
                                except Exception as exc:
                                    print(
                                        f"Error in {label}: {exc}",
                                        file=sys.stderr,
                                    )
                            elif debug:
                                print(
                                    f"{label} heard, cooldown active.",
                                    flush=True,
                                )
                            break
                elif debug:
                    partial = json.loads(rec.PartialResult())
                    p = partial.get("partial", "")
                    if p:
                        print(f"  [partial] \"{p}\"", end="\r", flush=True)
    except KeyboardInterrupt:
        print("\nMarkus signing off.", flush=True)


if __name__ == "__main__":
    main()
