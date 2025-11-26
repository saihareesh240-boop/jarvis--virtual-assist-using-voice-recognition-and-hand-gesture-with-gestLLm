# jarvis--virtual-assist-using-voice-recognition-and-hand-gesture-with-gestLLmimport os
import sys
import time
import math
import queue
import re
import webbrowser
import threading
import subprocess
import datetime
from pathlib import Path
from random import choice
from urllib.parse import quote_plus

# ---------------------------
# GUI import up-front (fix: RobotGUI needs tk available at class definition)
# ---------------------------
try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    TK_AVAILABLE = True
except Exception:
    tk = None
    ttk = None
    messagebox = None
    TK_AVAILABLE = False

# ---------------------------
# Optional third-party imports (graceful)
# ---------------------------
try:
    import pyttsx3
    TTS_AVAILABLE = True
except Exception:
    pyttsx3 = None
    TTS_AVAILABLE = False

try:
    import pywhatkit
    HAS_PYWHATKIT = True
except Exception:
    pywhatkit = None
    HAS_PYWHATKIT = False

try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except Exception:
    pyautogui = None
    PYAUTOGUI_AVAILABLE = False

try:
    from PIL import Image, ImageTk, ImageGrab
    PIL_AVAILABLE = True
except Exception:
    Image = None
    ImageTk = None
    ImageGrab = None
    PIL_AVAILABLE = False

try:
    import cv2
    CV2_AVAILABLE = True
except Exception:
    cv2 = None
    CV2_AVAILABLE = False

try:
    import mediapipe as mp
    MP_AVAILABLE = True
except Exception:
    mp = None
    MP_AVAILABLE = False

try:
    import speech_recognition as sr
    SR_AVAILABLE = True
except Exception:
    sr = None
    SR_AVAILABLE = False

# optional
try:
    import sounddevice as sd
    SD_AVAILABLE = True
except Exception:
    sd = None
    SD_AVAILABLE = False

try:
    import vosk
    VOSK_AVAILABLE = True
except Exception:
    vosk = None
    VOSK_AVAILABLE = False

# ---------------------------
# Configuration
# ---------------------------
SAMPLE_RATE = 16000
CAMERA_ID = 0
CAM_RESOLUTION = (640, 480)
SMOOTHING_BUFFER = 6
PINCH_THRESHOLD_PIXELS = 42
CLICK_COOLDOWN = 0.25
MOVE_DURATION = 0
WAKE_WORD = "jarvis"   # Wake word (case-insensitive)
HELP_IMAGE_PATH = "gesture_help.png"  # optional help image
HISTORY_FILE = Path("voice_history.txt")
MAX_HISTORY = 20

# --- UI theme: Cyber Blue ---
BG_MAIN = "#020617"       # overall window background (very dark navy)
BG_SIDE = "#020617"       # right panel
BG_CARD = "#02091f"       # panels / canvas background
BG_CANVAS = "#020817"
EYE_WHITE = "#e5f3ff"
EYE_PUPIL = "#0f172a"
TEXT_MAIN = "#e0f2fe"
TEXT_SOFT = "#38bdf8"
TEXT_MUTED = "#64748b"
TEXT_WARN = "#fca5a5"
ACCENT = "#22d3ee"
ACCENT_STRONG = "#0ea5e9"
LIST_BG = "#020617"
LIST_SELECT_BG = "#0f172a"

# ---------------------------
# Shared state
# ---------------------------
program_should_exit = threading.Event()
assistant_should_run = threading.Event()   # controls assistant without exiting app
gesture_enabled = threading.Event()
gesture_enabled.set()
overlay_msg_q = queue.Queue(maxsize=8)
last_voice_text = {"text": ""}

MIRROR_CAMERA = True
MIRROR_CURSOR = True

# placeholder for vosk autodetect (kept for UI)
DETECTED_VOSK_MODEL = None

# ---------------------------
# TTS (pyttsx3)
# ---------------------------
if TTS_AVAILABLE:
    try:
        tts_engine = pyttsx3.init()
        try:
            rate = tts_engine.getProperty("rate") or 200
            tts_engine.setProperty("rate", int(rate * 0.95))
        except Exception:
            pass
    except Exception:
        tts_engine = None
        TTS_AVAILABLE = False
else:
    tts_engine = None


def speak(text: str):
    print("[Assistant]:", text)
    try:
        if TTS_AVAILABLE and tts_engine:
            tts_engine.say(str(text))
            tts_engine.runAndWait()
    except Exception as e:
        print("TTS error:", e)


# ---------------------------
# Small helpers
# ---------------------------
def np_interpolate(val, src, dst):
    (a, b), (c, d) = src, dst
    if b == a:
        return dst[0]
    return c + (d - c) * ((val - a) / (b - a))


# ---------------------------
# History helpers
# ---------------------------
def load_history():
    items = []
    try:
        if HISTORY_FILE.exists():
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        items.append(line)
    except Exception:
        pass
    return items[-MAX_HISTORY:]


def append_history(cmd):
    try:
        with open(HISTORY_FILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now().isoformat()} | {cmd}\n")
    except Exception: pass


# ---------------------------
# WhatsApp helpers
# ---------------------------
def _attempt_auto_send(wait_for_load=8, message=None):
    if not PYAUTOGUI_AVAILABLE:
        return False
    try:
        speak(f"Waiting {wait_for_load} seconds for WhatsApp Web to load...")
        time.sleep(wait_for_load)
        try:
            if sys.platform.startswith("darwin"):
                pyautogui.keyDown("command")
                pyautogui.press("tab")
                pyautogui.keyUp("command")
            else:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                pyautogui.keyUp("alt")
        except Exception:
            pass
        time.sleep(0.8)
        pyautogui.press("enter")
        time.sleep(0.7)
        screen_w, screen_h = pyautogui.size()
        click_x = int(screen_w * 0.5)
        click_y = int(screen_h * 0.92)
        pyautogui.click(click_x, click_y)
        time.sleep(0.3)
        if message:
            try:
                pyautogui.typewrite(message, interval=0.01)
            except Exception:
                pass
        pyautogui.press("enter")
        speak("Auto-send attempted using pyautogui. Check WhatsApp Web to confirm.")
        return True
    except Exception as e:
        speak("Auto-send via pyautogui failed: " + str(e))
        return False


def open_whatsapp(number: str = None, message: str = None, auto_send: bool = False, wait_for_load: float = 6.0):
    try:
        if number:
            num = re.sub(r"\D+", "", number)
            url = f"https://web.whatsapp.com/send?phone={num}"
            if message:
                url += f"&text={quote_plus(message)}"
            speak("Opening WhatsApp Web...")
            webbrowser.open(url)
            if auto_send and PYAUTOGUI_AVAILABLE:
                _attempt_auto_send(wait_for_load=wait_for_load, message=message)
        else:
            speak("Opening WhatsApp Web...")
            webbrowser.open("https://web.whatsapp.com")
    except Exception as e:
        speak("Failed to open WhatsApp: " + str(e))


def send_whatsapp(number: str, message: str, use_pywhatkit=True, auto_send: bool = True, wait_for_load: float = 8.0):
    try:
        if use_pywhatkit and HAS_PYWHATKIT:
            speak(f"Attempting to send via pywhatkit to {number}...")
            try:
                pywhatkit.sendwhatmsg_instantly(number, message, wait_time=10, tab_close=True)
                speak("pywhatkit attempted to send the message (verify in browser).")
                return
            except Exception as e:
                speak("pywhatkit failed: " + str(e))
    except Exception as e:
        speak("pywhatkit path error: " + str(e))

    try:
        num = re.sub(r"\D+", "", number)
        url = f"https://web.whatsapp.com/send?phone={num}&text={quote_plus(message)}"
        speak("Opening WhatsApp Web with prefilled message. I will try to auto-send if possible.")
        webbrowser.open(url)
        if auto_send and PYAUTOGUI_AVAILABLE:
            _attempt_auto_send(wait_for_load=wait_for_load, message=message)
        elif auto_send:
            speak("Auto-send not available (pyautogui missing). Please send manually in browser.")
    except Exception as e:
        speak("Fallback open failed: " + str(e))


# ---------------------------
# Close page helper (new)
# ---------------------------
def close_page():
    try:
        if PYAUTOGUI_AVAILABLE:
            speak("Closing the current page...")
            if sys.platform.startswith("darwin"):
                pyautogui.hotkey("command", "w")
            else:
                pyautogui.hotkey("ctrl", "w")
            return True
        else:
            speak("I can't close the page automatically because 'pyautogui' is not installed.")
            speak("Please press Ctrl+W (or Command+W on Mac) to close the current tab/window.")
            return False
    except Exception as e:
        speak("Failed to close page automatically: " + str(e))
        return False


def close_target(target_name: str):
    tn = target_name.lower()
    try:
        if any(x in tn for x in ["edge", "microsoft edge"]):
            speak("Closing Microsoft Edge...")
            if sys.platform.startswith("win"):
                os.system("taskkill /f /im msedge.exe")
            else:
                try:
                    subprocess.Popen(["pkill", "-f", "msedge"])
                except Exception:
                    speak("Couldn't kill msedge automatically.")
            return True
        if "chrome" in tn:
            speak("Closing Chrome...")
            if sys.platform.startswith("win"):
                os.system("taskkill /f /im chrome.exe")
            else:
                try:
                    subprocess.Popen(["pkill", "-f", "chrome"])
                except Exception:
                    speak("Couldn't kill chrome automatically.")
            return True
        if any(x in tn for x in ["whatsapp", "youtube", "web.whatsapp.com", "youtube.com"]):
            if PYAUTOGUI_AVAILABLE:
                speak(f"Trying to close the browser tab for {target_name} (Ctrl+W)...")
                try:
                    if sys.platform.startswith("darwin"):
                        pyautogui.hotkey("command", "w")
                    else:
                        pyautogui.hotkey("ctrl", "w")
                    return True
                except Exception as e:
                    speak("Auto tab-close failed: " + str(e))
                    return False
            else:
                speak(f"I can't automatically close the {target_name} tab because pyautogui is not installed.")
                speak("Please switch to the browser tab and press Ctrl+W (Cmd+W on Mac).")
                return False
        speak(f"I don't have a specific close action for '{target_name}'. Try 'close chrome' or 'close edge' or 'close whatsapp'.")
        return False
    except Exception as e:
        speak("Error in close_target: " + str(e))
        return False


# ---------------------------
# Robot GUI
# ---------------------------
class RobotGUI:
    def __init__(self, root):
        if not TK_AVAILABLE:
            raise RuntimeError("Tkinter not available on this system.")
        self.root = root
        root.title("Facetelligence — Cyber Blue Assistant")
        root.protocol("WM_DELETE_WINDOW", self.on_quit)

        # geometry + min/max enforcement + resizable
        root.configure(bg=BG_MAIN)
        root.geometry("1040x680")
        root.minsize(980, 640)
        root.maxsize(1400, 980)
        root.resizable(True, True)

        left = tk.Frame(root, width=740, height=680, bg=BG_MAIN, highlightthickness=0)
        left.pack(side="left", fill="both", expand=True)
        right = tk.Frame(root, width=300, height=680, bg=BG_SIDE, highlightthickness=0)
        right.pack(side="right", fill="y")

        # Robot face + camera
        self.face_bg = tk.Canvas(
            left,
            width=700,
            height=560,
            bg=BG_CARD,
            highlightthickness=2,
            highlightbackground=ACCENT,
        )
        self.face_bg.place(x=10, y=10)

        self.camera_panel = tk.Label(
            left,
            bg=BG_CANVAS,
            highlightthickness=1,
            highlightbackground=ACCENT_STRONG,
        )
        self.camera_panel.place(x=50, y=40, width=600, height=420)

        # eyes
        self.left_eye_center = (210, 110)
        self.right_eye_center = (430, 110)
        self.eye_radius = 34
        self.pupil_radius = 12

        self.face_bg.create_oval(
            self.left_eye_center[0] - self.eye_radius,
            self.left_eye_center[1] - self.eye_radius,
            self.left_eye_center[0] + self.eye_radius,
            self.left_eye_center[1] + self.eye_radius,
            fill=EYE_WHITE,
            outline=ACCENT_STRONG,
            width=2,
        )
        self.face_bg.create_oval(
            self.right_eye_center[0] - self.eye_radius,
            self.right_eye_center[1] - self.eye_radius,
            self.right_eye_center[0] + self.eye_radius,
            self.right_eye_center[1] + self.eye_radius,
            fill=EYE_WHITE,
            outline=ACCENT_STRONG,
            width=2,
        )
        self.left_pupil = self.face_bg.create_oval(
            self.left_eye_center[0] - self.pupil_radius,
            self.left_eye_center[1] - self.pupil_radius,
            self.left_eye_center[0] + self.pupil_radius,
            self.left_eye_center[1] + self.pupil_radius,
            fill=EYE_PUPIL,
            outline="",
        )
        self.right_pupil = self.face_bg.create_oval(
            self.right_eye_center[0] - self.pupil_radius,
            self.right_eye_center[1] - self.pupil_radius,
            self.right_eye_center[0] + self.pupil_radius,
            self.right_eye_center[1] + self.pupil_radius,
            fill=EYE_PUPIL,
            outline="",
        )

        self.status_text_id = self.face_bg.create_text(
            350,
            520,
            text="Initializing...",
            fill=TEXT_SOFT,
            font=("Consolas", 11, "bold"),
        )

        lbl = tk.Label(
            right,
            text="CONTROL PANEL",
            fg=ACCENT,
            bg=BG_SIDE,
            font=("Consolas", 12, "bold"),
        )
        lbl.pack(pady=(12, 6))

        model_text = "Yes" if (VOSK_AVAILABLE and DETECTED_VOSK_MODEL) else "No"
        self.status_list = tk.Label(
            right,
            text=f"Model: {model_text}\nListening: —\nGestures: ON\nLast: —",
            justify="left",
            fg=TEXT_MAIN,
            bg=BG_SIDE,
            font=("Consolas", 9),
        )
        self.status_list.pack(pady=(0, 8))

        # WhatsApp panel
        wa_frame = tk.LabelFrame(
            right,
            text=" WhatsApp ",
            fg=ACCENT,
            bg=BG_SIDE,
            bd=1,
            relief="solid",
            labelanchor="n",
            font=("Consolas", 9, "bold"),
        )
        wa_frame.pack(padx=8, pady=6, fill="x")

        tk.Label(
            wa_frame,
            text="Phone (+country):",
            bg=BG_SIDE,
            fg=TEXT_SOFT,
            font=("Consolas", 9),
        ).pack(anchor="w", padx=6, pady=(6, 0))
        self.phone_entry = tk.Entry(wa_frame, width=28, bg="#020617", fg=TEXT_MAIN, insertbackground=TEXT_MAIN)
        self.phone_entry.pack(padx=6, pady=2)

        tk.Label(
            wa_frame,
            text="Message:",
            bg=BG_SIDE,
            fg=TEXT_SOFT,
            font=("Consolas", 9),
        ).pack(anchor="w", padx=6, pady=(4, 0))
        self.msg_text = tk.Text(wa_frame, width=30, height=4, bg="#020617", fg=TEXT_MAIN, insertbackground=TEXT_MAIN)
        self.msg_text.pack(padx=6, pady=2)

        self.auto_send_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            wa_frame,
            text="Auto-send (pyautogui)",
            variable=self.auto_send_var,
            bg=BG_SIDE,
            fg=TEXT_MAIN,
            selectcolor=BG_SIDE,
            font=("Consolas", 8),
        ).pack(anchor="w", padx=6, pady=(2, 6))

        btns = tk.Frame(wa_frame, bg=BG_SIDE)
        btns.pack(fill="x", padx=6, pady=(0, 6))

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Cyber.TButton",
            background="#020617",
            foreground=TEXT_MAIN,
            borderwidth=1,
            focusthickness=0,
            padding=3,
        )
        style.map("Cyber.TButton", background=[("active", "#0b1120")])

        self.open_btn = ttk.Button(
            btns,
            text="Open Chat",
            width=12,
            style="Cyber.TButton",
            command=self.ui_open_whatsapp,
        )
        self.open_btn.pack(side="left", padx=(0, 6))

        self.send_btn = ttk.Button(
            btns,
            text="Send Now",
            width=12,
            style="Cyber.TButton",
            command=self.ui_send_whatsapp,
        )
        self.send_btn.pack(side="left")

        self.wa_status = tk.Label(
            wa_frame,
            text="Status: idle",
            bg=BG_SIDE,
            fg=TEXT_MUTED,
            font=("Consolas", 8),
        )
        self.wa_status.pack(anchor="w", padx=6, pady=(2, 6))

        # controls frame
        ctl_frame = tk.Frame(right, bg=BG_SIDE)
        ctl_frame.pack(pady=(6, 0), fill="x")

        # --- Window size controls (min / max) ---
        size_frame = tk.Frame(right, bg=BG_SIDE)
        size_frame.pack(padx=8, pady=(6, 4), fill="x")

        tk.Label(
            size_frame,
            text="Min Size (W x H):",
            bg=BG_SIDE,
            fg=TEXT_SOFT,
            font=("Consolas", 8),
        ).grid(row=0, column=0, sticky="w", padx=6, pady=2)

        self.min_w_var = tk.StringVar(value=str(root.winfo_reqwidth()))
        self.min_h_var = tk.StringVar(value=str(root.winfo_reqheight()))

        tk.Entry(size_frame, textvariable=self.min_w_var, width=6, bg="#020617", fg=TEXT_MAIN, insertbackground=TEXT_MAIN).grid(
            row=0, column=1, padx=(0, 6)
        )
        tk.Entry(size_frame, textvariable=self.min_h_var, width=6, bg="#020617", fg=TEXT_MAIN, insertbackground=TEXT_MAIN).grid(
            row=0, column=2, padx=(0, 6)
        )

        tk.Label(
            size_frame,
            text="Max Size (W x H):",
            bg=BG_SIDE,
            fg=TEXT_SOFT,
            font=("Consolas", 8),
        ).grid(row=1, column=0, sticky="w", padx=6, pady=2)

        self.max_w_var = tk.StringVar(value="1400")
        self.max_h_var = tk.StringVar(value="980")

        tk.Entry(size_frame, textvariable=self.max_w_var, width=6, bg="#020617", fg=TEXT_MAIN, insertbackground=TEXT_MAIN).grid(
            row=1, column=1, padx=(0, 6)
        )
        tk.Entry(size_frame, textvariable=self.max_h_var, width=6, bg="#020617", fg=TEXT_MAIN, insertbackground=TEXT_MAIN).grid(
            row=1, column=2, padx=(0, 6)
        )

        def apply_sizes():
            try:
                mw = int(self.min_w_var.get())
                mh = int(self.min_h_var.get())
                Mw = int(self.max_w_var.get())
                Mh = int(self.max_h_var.get())
                if mw > Mw or mh > Mh:
                    self.show_message("Min must be <= Max")
                    speak("Minimum size must be less than or equal to maximum size.")
                    return
                root.minsize(mw, mh)
                root.maxsize(Mw, Mh)
                self.show_message(f"Size applied: min {mw}x{mh}, max {Mw}x{Mh}")
            except Exception:
                self.show_message("Invalid size values")
                speak("Invalid size values. Please enter integers.")

        ttk.Button(
            size_frame,
            text="Apply",
            style="Cyber.TButton",
            command=apply_sizes,
            width=7,
        ).grid(row=0, column=3, rowspan=2, padx=8)

        # misc controls (after size controls)
        self.toggle_g_btn = ttk.Button(
            ctl_frame,
            text="Toggle Gestures (G)",
            style="Cyber.TButton",
            command=self.toggle_gestures,
        )
        self.toggle_g_btn.pack(fill="x", padx=6, pady=(4, 6))

        self.mirror_cursor_var = tk.BooleanVar(value=MIRROR_CURSOR)
        tk.Checkbutton(
            ctl_frame,
            text="Mirror Cursor",
            variable=self.mirror_cursor_var,
            onvalue=True,
            offvalue=False,
            bg=BG_SIDE,
            fg=TEXT_MAIN,
            selectcolor=BG_SIDE,
            font=("Consolas", 8),
            command=self.on_toggle_cursor_mirror,
        ).pack(padx=6, pady=(0, 6))

        ttk.Button(
            ctl_frame,
            text="Show Help",
            style="Cyber.TButton",
            command=self.show_help_image,
        ).pack(fill="x", padx=6, pady=(0, 6))

        ttk.Button(
            ctl_frame,
            text="Close Page (C)",
            style="Cyber.TButton",
            command=close_page,
        ).pack(fill="x", padx=6, pady=(0, 6))

        ttk.Button(
            ctl_frame,
            text="Close WhatsApp Tab",
            style="Cyber.TButton",
            command=lambda: close_target("whatsapp"),
        ).pack(fill="x", padx=6, pady=(0, 6))

        ttk.Button(
            ctl_frame,
            text="Close YouTube Tab",
            style="Cyber.TButton",
            command=lambda: close_target("youtube"),
        ).pack(fill="x", padx=6, pady=(0, 6))

        ttk.Button(
            ctl_frame,
            text="Quit",
            style="Cyber.TButton",
            command=self.on_quit,
        ).pack(fill="x", padx=6, pady=(2, 6))

        # assistant controls
        ttk.Separator(ctl_frame, orient="horizontal").pack(fill="x", pady=(6, 6))

        self.auto_start_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            ctl_frame,
            text="Auto-start assistant",
            variable=self.auto_start_var,
            bg=BG_SIDE,
            fg=TEXT_MUTED,
            selectcolor=BG_SIDE,
            font=("Consolas", 8),
        ).pack(anchor="w", padx=6)

        ttk.Button(
            ctl_frame,
            text="Voice Typing (single-line) (V)",
            style="Cyber.TButton",
            command=lambda: threading.Thread(
                target=voice_typing,
                kwargs={"single_line": True, "press_enter": False},
                daemon=True,
            ).start(),
        ).pack(fill="x", padx=6, pady=(6, 4))

        ttk.Button(
            ctl_frame,
            text="Voice Typing (Enter)",
            style="Cyber.TButton",
            command=lambda: threading.Thread(
                target=voice_typing,
                kwargs={"single_line": True, "press_enter": True},
                daemon=True,
            ).start(),
        ).pack(fill="x", padx=6, pady=(0, 6))

        ttk.Button(
            ctl_frame,
            text="Start Assistant (S)",
            style="Cyber.TButton",
            command=self.start_assistant,
        ).pack(fill="x", padx=6, pady=(4, 6))

        ttk.Button(
            ctl_frame,
            text="Stop Assistant (X)",
            style="Cyber.TButton",
            command=self.stop_assistant,
        ).pack(fill="x", padx=6, pady=(0, 6))

        # History panel
        hist_frame = tk.LabelFrame(
            right,
            text=" Voice History (double-click to rerun) ",
            fg=ACCENT,
            bg=BG_SIDE,
            bd=1,
            relief="solid",
            font=("Consolas", 9, "bold"),
        )
        hist_frame.pack(padx=8, pady=6, fill="both", expand=True)

        self.history_list = tk.Listbox(
            hist_frame,
            height=8,
            bg=LIST_BG,
            fg=TEXT_MAIN,
            selectbackground=LIST_SELECT_BG,
            activestyle="none",
            borderwidth=0,
            highlightthickness=0,
            font=("Consolas", 9),
        )
        self.history_list.pack(fill="both", expand=True, padx=6, pady=6)
        self.history_list.bind("<Double-Button-1>", self.on_history_double)

        self.latest_camera_pil = None
        self.latest_finger_pos = None
        self._assistant_running = False
        self._assistant_thread = None
        self.root.after(80, self._periodic_update)

        # keyboard bindings
        root.bind("<s>", lambda e: self.start_assistant())
        root.bind("<S>", lambda e: self.start_assistant())
        root.bind("<x>", lambda e: self.stop_assistant())
        root.bind("<X>", lambda e: self.stop_assistant())
        root.bind("<g>", lambda e: self.toggle_gestures())
        root.bind("<G>", lambda e: self.toggle_gestures())
        root.bind(
            "<v>",
            lambda e: threading.Thread(
                target=voice_typing, kwargs={"single_line": True, "press_enter": False}, daemon=True
            ).start(),
        )
        root.bind(
            "<V>",
            lambda e: threading.Thread(
                target=voice_typing, kwargs={"single_line": True, "press_enter": False}, daemon=True
            ).start(),
        )
        root.bind("<c>", lambda e: close_page())
        root.bind("<C>", lambda e: close_page())

        # populate history
        self.refresh_history()

    def refresh_history(self):
        self.history_list.delete(0, tk.END)
        items = load_history()
        for line in reversed(items):
            display = line.split("|", 1)[-1].strip()
            if len(display) > 60:
                display = display[:57] + "..."
            self.history_list.insert(0, display)

    def on_history_double(self, event):
        sel = self.history_list.curselection()
        if not sel:
            return
        idx = sel[0]
        text = self.history_list.get(idx)
        threading.Thread(target=self.run_text_command, args=(text,), daemon=True).start()

    def run_text_command(self, text):
        last_voice_text["text"] = text
        append_history(f"RERUN | {text}")
        self.show_message("Rerun: " + (text[:40] + "..." if len(text) > 40 else text))
        t = text.lower()
        if any(p in t for p in ["close page", "close tab"]):
            close_page()
            return
        if t.startswith("open "):
            open_software(t.replace("open ", "", 1).strip())
            return
        if t.startswith("close "):
            close_software(t.replace("close ", "", 1).strip())
            return
        if t.startswith("play "):
            play_song(t.replace("play ", "", 1).strip())
            return
        web_search(text)

    def ui_open_whatsapp(self):
        phone = self.phone_entry.get().strip()
        message = self.msg_text.get("1.0", "end").strip()
        self.wa_status.config(text="Status: opening...")
        open_whatsapp(phone if phone else None, message if message else None, auto_send=self.auto_send_var.get())
        self.wa_status.config(text="Status: opened")

    def ui_send_whatsapp(self):
        phone = self.phone_entry.get().strip()
        message = self.msg_text.get("1.0", "end").strip()
        if not phone:
            self.wa_status.config(text="Status: enter phone")
            return
        if not message:
            self.wa_status.config(text="Status: enter message")
            return
        self.wa_status.config(text="Status: sending...")
        send_whatsapp(phone, message, use_pywhatkit=True, auto_send=self.auto_send_var.get())
        self.wa_status.config(text="Status: attempted")

    def toggle_gestures(self):
        if gesture_enabled.is_set():
            gesture_enabled.clear()
            self.show_message("Gestures disabled")
            speak("Gesture control disabled")
        else:
            gesture_enabled.set()
            self.show_message("Gestures enabled")
            speak("Gesture control enabled")
        self.refresh_history()

    def on_toggle_cursor_mirror(self):
        global MIRROR_CURSOR
        MIRROR_CURSOR = bool(self.mirror_cursor_var.get())
        self.show_message(f"Mirror cursor: {MIRROR_CURSOR}")

    def show_help_image(self):
        if not os.path.isfile(HELP_IMAGE_PATH):
            self.show_message("Help image not found")
            return
        w = tk.Toplevel(self.root)
        w.title("Gesture Help")
        w.configure(bg=BG_MAIN)
        try:
            if PIL_AVAILABLE:
                img = Image.open(HELP_IMAGE_PATH)
                img = img.resize((700, 520), Image.LANCZOS)
                imgtk = ImageTk.PhotoImage(img)
                lbl = tk.Label(w, image=imgtk, bg=BG_MAIN)
                lbl.image = imgtk
                lbl.pack()
            else:
                tk.Label(w, text="Pillow not installed; can't load image.", bg=BG_MAIN, fg=TEXT_MAIN).pack()
        except Exception as e:
            tk.Label(w, text=f"Failed to load help image: {e}", bg=BG_MAIN, fg=TEXT_WARN).pack()

    def start_assistant(self):
        if self._assistant_running:
            self.show_message("Assistant already running")
            return
        assistant_should_run.set()
        self._assistant_thread = threading.Thread(target=lambda: assistant_loop(gui=self), daemon=True)
        self._assistant_thread.start()
        self._assistant_running = True
        self.show_message("Assistant started")
        speak("Assistant started")
        append_history("ASSISTANT STARTED")
        self.set_status(
            listening=SR_AVAILABLE,
            gestures_on=gesture_enabled.is_set(),
            last_text=last_voice_text.get("text", ""),
        )

    def stop_assistant(self):
        if not self._assistant_running:
            self.show_message("Assistant not running")
            return
        assistant_should_run.clear()
        self._assistant_running = False
        self.show_message("Assistant stopping")
        speak("Stopping assistant")
        append_history("ASSISTANT STOPPED")

    def on_quit(self):
        if messagebox:
            if messagebox.askokcancel("Quit", "Stop assistant and quit?"):
                program_should_exit.set()
                assistant_should_run.clear()
                self.root.destroy()
        else:
            program_should_exit.set()
            assistant_should_run.clear()
            self.root.destroy()

    def show_message(self, text):
        try:
            overlay_msg_q.put_nowait(text)
        except queue.Full:
            pass

    def set_status(self, listening=False, gestures_on=False, last_text=""):
        model_status = "Yes" if (VOSK_AVAILABLE and DETECTED_VOSK_MODEL) else "No"
        s = (
            f"Model: {model_status}\n"
            f"Listening: {'Yes' if listening else 'No'}\n"
            f"Gestures: {'ON' if gestures_on else 'OFF'}\n"
            f"Last: {last_text[:28]}{'...' if len(last_text) > 28 else ''}"
        )
        self.status_list.config(text=s)
        try:
            self.face_bg.itemconfigure(
                self.status_text_id,
                text="Listening"
                if listening
                else ("Gestures ON" if gestures_on else "Idle"),
            )
        except Exception:
            pass

    def update_camera(self, cv2_frame):
        if cv2_frame is None:
            return
        if not PIL_AVAILABLE:
            return
        try:
            frame = cv2.resize(cv2_frame, (600, 420))
            img = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            pil = Image.fromarray(img)
            imgtk = ImageTk.PhotoImage(image=pil)
            self.camera_panel.imgtk = imgtk
            self.camera_panel.configure(image=imgtk)
        except Exception:
            pass

    def update_eyes(self, target_screen_pos):
        face_w = 600
        face_h = 420
        if target_screen_pos is None:
            tx, ty = face_w / 2, face_h / 3
        else:
            sx, sy = target_screen_pos
            try:
                sw, sh = pyautogui.size() if PYAUTOGUI_AVAILABLE else (1920, 1080)
            except Exception:
                sw, sh = (1920, 1080)
            tx = np_interpolate(sx, [0, sw], [80, 500])
            ty = np_interpolate(sy, [0, sh], [60, 180])

        def place_pupil(center, tx, ty):
            cx, cy = center
            dx = tx - cx
            dy = ty - cy
            maxd = self.eye_radius - self.pupil_radius - 4
            d = math.hypot(dx, dy)
            if d > maxd and d != 0:
                scale = maxd / d
                dx *= scale
                dy *= scale
            return (int(cx + dx), int(cy + dy))

        lx, ly = place_pupil(self.left_eye_center, tx, ty)
        rx, ry = place_pupil(self.right_eye_center, tx, ty)
        self.face_bg.coords(
            self.left_pupil,
            lx - self.pupil_radius,
            ly - self.pupil_radius,
            lx + self.pupil_radius,
            ly + self.pupil_radius,
        )
        self.face_bg.coords(
            self.right_pupil,
            rx - self.pupil_radius,
            ry - self.pupil_radius,
            rx + self.pupil_radius,
            ry + self.pupil_radius,
        )

    def _periodic_update(self):
        listening = SR_AVAILABLE and assistant_should_run.is_set()
        self.set_status(
            listening=listening,
            gestures_on=gesture_enabled.is_set(),
            last_text=last_voice_text.get("text", ""),
        )
        try:
            msg = overlay_msg_q.get_nowait()
            self.face_bg.itemconfigure(self.status_text_id, text=msg)
        except queue.Empty:
            pass
        self.update_eyes(self.latest_finger_pos)
        if not program_should_exit.is_set():
            self.root.after(120, self._periodic_update)


# ---------------------------
# Gesture thread (MediaPipe + pyautogui)
# ---------------------------
def gesture_thread_fn(gui: RobotGUI):
    if not (CV2_AVAILABLE and MP_AVAILABLE and PYAUTOGUI_AVAILABLE):
        print("[Gesture] Missing cv2/mediapipe/pyautogui - gesture disabled.")
        return
    mp_hands = mp.solutions.hands
    hands = mp_hands.Hands(
        static_image_mode=False,
        max_num_hands=1,
        min_detection_confidence=0.6,
        min_tracking_confidence=0.6,
    )
    cap = cv2.VideoCapture(CAMERA_ID, cv2.CAP_DSHOW if sys.platform.startswith("win") else 0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAM_RESOLUTION[0])
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAM_RESOLUTION[1])
    SCREEN_W, SCREEN_H = pyautogui.size()
    pts_x = []
    pts_y = []
    last_click_time = 0
    cam_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    cam_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    print(f"[Gesture] Camera {cam_w}x{cam_h}, Screen {SCREEN_W}x{SCREEN_H}")

    while not program_should_exit.is_set():
        ret, frame = cap.read()
        if not ret:
            break
        if MIRROR_CAMERA:
            frame = cv2.flip(frame, 1)
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = hands.process(frame_rgb)
        gui.update_camera(frame)
        if results.multi_hand_landmarks and gesture_enabled.is_set():
            hand = results.multi_hand_landmarks[0]
            ix = int(hand.landmark[8].x * cam_w)
            iy = int(hand.landmark[8].y * cam_h)
            tx = int(hand.landmark[4].x * cam_w)
            ty = int(hand.landmark[4].y * cam_h)
            if MIRROR_CAMERA and MIRROR_CURSOR:
                ix = cam_w - ix
                tx = cam_w - tx
            screen_x = np_interpolate(ix, [0, cam_w], [0, SCREEN_W])
            screen_y = np_interpolate(iy, [0, cam_h], [0, SCREEN_H])
            pts_x.append(screen_x)
            pts_y.append(screen_y)
            if len(pts_x) > SMOOTHING_BUFFER:
                pts_x.pop(0)
                pts_y.pop(0)
            smooth_x = int(sum(pts_x) / len(pts_x))
            smooth_y = int(sum(pts_y) / len(pts_y))
            try:
                pyautogui.moveTo(smooth_x, smooth_y, duration=MOVE_DURATION)
            except Exception as e:
                print("[Gesture] move error", e)
            pinch_dist = math.hypot(ix - tx, iy - ty)
            now = time.time()
            if pinch_dist < PINCH_THRESHOLD_PIXELS and (now - last_click_time) > CLICK_COOLDOWN:
                last_click_time = now
                try:
                    pyautogui.click()
                    overlay_msg_q.put("Click")
                except Exception as e:
                    print("[Gesture] click error", e)
            gui.latest_finger_pos = (smooth_x, smooth_y)
        else:
            gui.latest_finger_pos = None
        time.sleep(0.01)
    cap.release()
    print("[Gesture] exit")


# ---------------------------
# Voice command handling (single-command-per-wake)
# ---------------------------
recognizer = sr.Recognizer() if SR_AVAILABLE else None

# helper utilities (screenshots, camera, selfie, file ops, timers, jokes)
def take_screenshot():
    if not PIL_AVAILABLE:
        speak("Screenshot feature requires Pillow. Please install pillow (pip install Pillow).")
        return
    screenshots_dir = Path("screenshots")
    screenshots_dir.mkdir(exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = screenshots_dir / f"screenshot_{timestamp}.png"
    try:
        img = ImageGrab.grab()
        img.save(str(filename))
        speak(f"Screenshot saved as {filename}")
    except Exception as e:
        speak("Failed to take screenshot: " + str(e))


camera_thread = None
camera_running = False


def camera_loop():
    global camera_running
    if not CV2_AVAILABLE:
        speak("Camera feature requires OpenCV (cv2). Install it with 'pip install opencv-python'.")
        return
    try:
        cap = cv2.VideoCapture(0)
        speak("Camera opened. Press 'q' in the window to close (or say 'close camera').")
        camera_running = True
        while camera_running and cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                break
            cv2.imshow("Camera Preview", frame)
            if cv2.waitKey(1) & 0xFF == ord("q"):
                break
        cap.release()
        cv2.destroyAllWindows()
    except Exception as e:
        speak("Camera error: " + str(e))
    camera_running = False


def open_camera():
    global camera_thread, camera_running
    if camera_running:
        speak("Camera already running.")
        return
    camera_thread = threading.Thread(target=camera_loop, daemon=True)
    camera_thread.start()


def close_camera():
    global camera_running
    if camera_running:
        camera_running = False
        speak("Closing camera...")
    else:
        speak("Camera is not running.")


def open_software(software_name):
    software_name = software_name.lower()
    try:
        if "chrome" in software_name:
            speak("Opening Chrome...")
            if sys.platform.startswith("win"):
                program = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                try:
                    subprocess.Popen([program])
                    return
                except Exception:
                    pass
            webbrowser.open("https://www.google.com")
        elif "edge" in software_name or "microsoft edge" in software_name:
            speak("Opening Microsoft Edge...")
            if sys.platform.startswith("win"):
                program = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                try:
                    subprocess.Popen([program])
                    return
                except Exception:
                    pass
            webbrowser.open("https://www.bing.com")
        elif "notepad" in software_name:
            speak("Opening Notepad...")
            if sys.platform.startswith("win"):
                subprocess.Popen(["notepad.exe"])
                return
            speak("Notepad only supported on Windows.")
        elif "calculator" in software_name:
            speak("Opening Calculator...")
            if sys.platform.startswith("win"):
                subprocess.Popen(["calc.exe"])
                return
            speak("Calculator opening not implemented for this platform.")
        elif software_name.startswith("play "):
            query = software_name.replace("play ", "").strip()
            play_song(query)
        else:
            speak(f"I couldn't find the software {software_name}")
    except Exception as e:
        speak("Failed to open software: " + str(e))


def close_software(software_name):
    software_name = software_name.lower()
    try:
        if "chrome" in software_name:
            speak("Closing Chrome...")
            if sys.platform.startswith("win"):
                os.system("taskkill /f /im chrome.exe")
            else:
                subprocess.Popen(["pkill", "-f", "chrome"])
        elif "edge" in software_name:
            speak("Closing Edge...")
            if sys.platform.startswith("win"):
                os.system("taskkill /f /im msedge.exe")
            else:
                subprocess.Popen(["pkill", "-f", "edge"])
        elif "notepad" in software_name:
            speak("Closing Notepad...")
            if sys.platform.startswith("win"):
                os.system("taskkill /f /im notepad.exe")
        elif "calculator" in software_name:
            speak("Closing Calculator...")
            if sys.platform.startswith("win"):
                os.system("taskkill /f /im calculator.exe")
        else:
            speak(f"I couldn't find any open software named {software_name}")
    except Exception as e:
        speak("Failed to close software: " + str(e))


def web_search(query):
    try:
        speak(f"Searching web for {query}")
        if HAS_PYWHATKIT:
            pywhatkit.search(query)
        else:
            webbrowser.open(f"https://www.google.com/search?q={quote_plus(query)}")
    except Exception as e:
        speak("Search failed: " + str(e))


def play_song(query):
    try:
        speak(f"Playing {query} on YouTube")
        if HAS_PYWHATKIT:
            pywhatkit.playonyt(query)
        else:
            webbrowser.open(f"https://www.youtube.com/results?search_query={quote_plus(query)}")
    except Exception as e:
        speak("Could not play song: " + str(e))


def open_chrome_and_search(query):
    try:
        if not query or not query.strip():
            speak("Please tell me what to search for.")
            return
        q = quote_plus(query.strip())
        search_url = f"https://www.google.com/search?q={q}"
        webbrowser.open(search_url)
        speak(f"Searching for {query}")
    except Exception as e:
        speak("Failed to open Chrome and search: " + str(e))


def take_selfie(filename: str = None, show_preview: bool = False):
    if not CV2_AVAILABLE:
        speak("Selfie requires OpenCV. Install opencv-python.")
        return
    try:
        selfies_dir = Path("selfies")
        selfies_dir.mkdir(exist_ok=True)
        if not filename:
            filename = datetime.datetime.now().strftime("selfie_%Y%m%d_%H%M%S.png")
        filepath = selfies_dir / filename
        cap = cv2.VideoCapture(0)
        time.sleep(0.5)
        ret, frame = cap.read()
        if not ret:
            speak("Failed to capture image.")
            cap.release()
            return
        cv2.imwrite(str(filepath), frame)
        cap.release()
        speak(f"Selfie saved as {filepath}")
        if show_preview:
            try:
                cv2.imshow("Selfie Preview", frame)
                cv2.waitKey(1500)
                cv2.destroyAllWindows()
            except Exception:
                pass
    except Exception as e:
        speak("Take selfie failed: " + str(e))


JOKES = [
    "Why did the programmer quit his job? Because he didn't get arrays.",
    "Why do programmers prefer dark mode? Because light attracts bugs.",
    "I told my computer I needed a break, and it said 'No problem — I'll go to sleep.'",
]


def tell_joke():
    speak(choice(JOKES))


def create_file(name):
    try:
        if not name.endswith(".txt"):
            name = name + ".txt"
        path = Path(name)
        if path.exists():
            speak(f"{name} already exists.")
            return
        path.write_text("")
        speak(f"Created file {name}")
    except Exception as e:
        speak("Failed to create file: " + str(e))


def read_file(name):
    try:
        if not name.endswith(".txt"):
            name = name + ".txt"
        path = Path(name)
        if not path.exists():
            speak(f"File {name} does not exist.")
            return
        content = path.read_text()
        if not content.strip():
            speak("File is empty.")
        else:
            speak("File content is:")
            speak(content)
    except Exception as e:
        speak("Failed to read file: " + str(e))


def confirm_and_execute(action):
    speak(f"Are you sure you want to {action}? Say yes to confirm.")
    if not SR_AVAILABLE:
        speak("Speech confirmation requires SpeechRecognition; skipping.")
        return False
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=4)
            ans = recognizer.recognize_google(audio, language="en_US").lower()
            if "yes" in ans or "yeah" in ans or "yep" in ans:
                return True
            else:
                speak("Cancelled.")
                return False
    except Exception:
        speak("No confirmation heard. Cancelled.")
        return False


def set_timer(minutes):
    def timer_thread(m):
        try:
            seconds = max(1, int(float(m) * 60))
        except Exception:
            speak("Couldn't parse the timer duration.")
            return
        speak(f"Timer set for {m} minutes.")
        time.sleep(seconds)
        speak(f"Timer finished: {m} minutes are up.")

    t = threading.Thread(target=timer_thread, args=(minutes,), daemon=True)
    t.start()


# ---------------------------
# Voice command: single-per-wake
# ---------------------------
continuous_dictation = False
dictation_thread = None


def voice_typing(single_line=True, press_enter=False, timeout=8):
    if not PYAUTOGUI_AVAILABLE:
        speak("Voice-typing requires pyautogui. Install it with 'pip install pyautogui'.")
        return
    if not SR_AVAILABLE:
        speak("SpeechRecognition is required for voice typing.")
        return
    speak("Voice typing started. Speak now...")
    while True:
        try:
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=0.4)
                audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=10)
            try:
                text = recognizer.recognize_google(audio, language="en_US").strip()
                if not text:
                    speak("I didn't catch anything.")
                    return
                if text.lower() in ["stop dictation", "stop typing", "stop"]:
                    speak("Stopping dictation.")
                    return
                try:
                    pyautogui.typewrite(text)
                    if press_enter:
                        pyautogui.press("enter")
                except Exception as e:
                    speak("pyautogui typing error: " + str(e))
                speak("Typed: " + text)
                append_history(text)
                if single_line:
                    return
                else:
                    speak("Continue speaking or say 'stop dictation' to stop.")
            except sr.UnknownValueError:
                speak("Sorry, I didn't understand.")
                if single_line:
                    return
            except Exception as e:
                speak("Recognition error: " + str(e))
                return
        except Exception as e:
            speak("Microphone error: " + str(e))
            return


def cmd_once(gui=None):
    if not SR_AVAILABLE:
        speak("SpeechRecognition not available; commands disabled.")
        if gui:
            gui.show_message("Voice disabled")
        return True
    with sr.Microphone() as source:
        try:
            recognizer.adjust_for_ambient_noise(source, duration=0.4)
            if gui:
                gui.show_message("Listening for command...")
            speak("Listening for command...")
            audio = recognizer.listen(source, timeout=7, phrase_time_limit=12)
        except Exception as e:
            speak("Listening error: " + str(e))
            if gui:
                gui.show_message("Listen error")
            return True
    try:
        text = recognizer.recognize_google(audio, language="en_US").lower().strip()
        last_voice_text["text"] = text
        append_history(text)
        print("Heard command:", text)
        if gui:
            gui.show_message("Heard: " + (text[:28] + "..." if len(text) > 28 else text))
            gui.set_status(
                listening=False,
                gestures_on=gesture_enabled.is_set(),
                last_text=text,
            )
            gui.refresh_history()
    except Exception as ex:
        print("Recognition error:", ex)
        speak("Sorry, I didn't catch that.")
        if gui:
            gui.show_message("Recognition error")
        return True

    # quick handlers
    if any(p in text for p in ["close page", "close tab", "close this page", "close the page"]):
        close_page()
        return True

    m_close = re.match(r"close (.+)", text)
    if m_close:
        target = m_close.group(1).strip()
        if target in ["camera", "close camera", "stop camera"]:
            close_camera()
            return True
        if any(x in target for x in ["edge", "microsoft edge", "chrome", "browser"]):
            close_target(target)
            return True
        if any(x in target for x in ["whatsapp", "youtube"]):
            close_target(target)
            return True
        close_software(target)
        return True

    if "shutdown program" in text or ("stop" in text and "program" in text):
        speak("Stopping the program. Goodbye!")
        program_should_exit.set()
        assistant_should_run.clear()
        return True

    if text.startswith("open "):
        software_name = text.replace("open ", "", 1).strip()
        if software_name.startswith("whatsapp"):
            m = re.search(r"open whatsapp(?: to )?\s*([+\d\s-]+)?(?: message (.+))?", software_name)
            if m:
                num = m.group(1)
                msg = m.group(2)
                open_whatsapp(num.strip(), msg.strip() if msg else None) if num else open_whatsapp()
            else:
                open_whatsapp()
        else:
            open_software(software_name)
        return True

    if "time" in text or "what time" in text:
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        speak(current_time)
        return True

    if "screenshot" in text:
        take_screenshot()
        return True

    if "open camera" in text:
        open_camera()
        return True

    if "close camera" in text:
        close_camera()
        return True

    if "take selfie" in text:
        take_selfie(show_preview=True)
        return True

    if text.startswith("search "):
        web_search(text.replace("search ", "", 1).strip())
        return True

    if text.startswith("play ") or "play song" in text:
        q = text.replace("play ", "", 1).strip() if text.startswith("play ") else text.replace("play song", "", 1).strip()
        play_song(q)
        return True

    if "voice typing" in text or "dictate" in text or "start dictation" in text:
        threading.Thread(
            target=voice_typing,
            kwargs={"single_line": True, "press_enter": False},
            daemon=True,
        ).start()
        return True

    if "voice typing enter" in text or "voice typing press enter" in text or "dictation send" in text:
        threading.Thread(
            target=voice_typing,
            kwargs={"single_line": True, "press_enter": True},
            daemon=True,
        ).start()
        return True

    if "send whatsapp" in text or (text.startswith("send ") and "whatsapp" in text):
        try:
            m = re.search(r"(?:send whatsapp to|send whatsapp|send)[:\s]*([+\d\s-]+)\s*(?:message\s*(.+))?", text)
            if m and m.group(1):
                number = m.group(1).strip()
                message = m.group(2) if m.group(2) else None
                if not message:
                    speak("No message content detected. What should I send?")
                    try:
                        with sr.Microphone() as src:
                            recognizer.adjust_for_ambient_noise(src, duration=0.4)
                            a = recognizer.listen(src, timeout=7, phrase_time_limit=10)
                            message = recognizer.recognize_google(a, language="en_US")
                    except Exception:
                        message = None
                if number and message:
                    send_whatsapp(number, message)
                elif number and not message:
                    open_whatsapp(number)
                else:
                    speak("Please say: send whatsapp to plus number message your message.")
            else:
                speak("Please say: send whatsapp to plus sign and number message your message.")
        except Exception as e:
            speak("Failed to parse whatsapp command: " + str(e))
        return True

    speak(
        "Sorry, I didn't understand that command. "
        "Try opening programs, taking screenshots, camera, web search, playing songs, voice typing, or sending WhatsApp."
    )
    return True


def listen_for_wake_word(timeout=7):
    """
    Listen only for wake word; do NOT talk back when wake word is detected.
    """
    if not SR_AVAILABLE:
        print("SpeechRecognition not available; cannot listen for wake word.")
        return False
    with sr.Microphone() as source:
        try:
            recognizer.adjust_for_ambient_noise(source, duration=0.4)
            print("Listening for wake word...")
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=4)
            text = recognizer.recognize_google(audio, language="en_US").lower()
            print("Heard:", text)
            last_voice_text["text"] = text
            if WAKE_WORD in text:
                # NO speak('Yes?') here – silent wake
                return True
        except Exception:
            return False
    return False


def assistant_loop(gui=None):
    speak(f"Jarvis started. Say '{WAKE_WORD}' to wake me.")
    if gui:
        gui.show_message("Assistant ready")
        gui.set_status(
            listening=SR_AVAILABLE and assistant_should_run.is_set(),
            gestures_on=gesture_enabled.is_set(),
            last_text=last_voice_text.get("text", ""),
        )
    while not program_should_exit.is_set():
        try:
            if not assistant_should_run.is_set():
                time.sleep(0.2)
                continue
            if listen_for_wake_word(timeout=7):
                cmd_once(gui=gui)
                if gui:
                    gui.set_status(
                        listening=SR_AVAILABLE and assistant_should_run.is_set(),
                        gestures_on=gesture_enabled.is_set(),
                        last_text=last_voice_text.get("text", ""),
                    )
            time.sleep(0.1)
        except KeyboardInterrupt:
            break
        except Exception as e:
            print("Assistant loop error:", e)
            time.sleep(0.5)
    speak("Assistant loop stopping.")
    if gui:
        gui.show_message("Assistant stopped")


# ---------------------------
# Main: GUI + threads
# ---------------------------
def main():
    if not TK_AVAILABLE:
        speak("Tkinter not available. GUI mode is required for this assistant.")
        return
    root = tk.Tk()
    gui = RobotGUI(root)

    gesture_thread = threading.Thread(target=gesture_thread_fn, args=(gui,), daemon=True)
    gesture_thread.start()

    def auto_start_check():
        if gui.auto_start_var.get():
            gui.start_assistant()

    root.after(800, auto_start_check)

    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass
    program_should_exit.set()
    assistant_should_run.clear()
    time.sleep(0.3)
    speak("Goodbye.")


if __name__ == "__main__":
    main()
