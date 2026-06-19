#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOSS Skill Cooldown Timer
Cross-platform (Windows/macOS), zero external dependencies.
Windows EXE: pip install pyinstaller && pyinstaller --onefile --windowed boss_timer.pyw
"""
import tkinter as tk
from tkinter import messagebox, simpledialog
import json, os, time, threading, platform, sys

# ── Constants ───────────────────────────────────────────────────
APP_NAME    = "BOSS技能倒计时器"
APP_VERSION = "3.1"

FONT_NORMAL = ("Microsoft YaHei", 11) if platform.system() == "Windows" else ("PingFang SC", 13)
FONT_BOLD   = (FONT_NORMAL[0], FONT_NORMAL[1], "bold")
FONT_SMALL  = (FONT_NORMAL[0], FONT_NORMAL[1] - 1)

COLORS = {
    "bg":        "#0f0f1a",
    "panel":     "#161630",
    "header":    "#1a1a35",
    "row":       "#12122a",
    "row_warn":  "#1a0f0f",
    "border":    "#2a2a4a",
    "text":      "#e5e5e5",
    "muted":     "#a0a0b0",
    "dim":       "#555555",
    "green":     "#22c55e",
    "yellow":    "#f59e0b",
    "red":       "#ef4444",
    "blue":      "#3b82f6",
}

IS_WINDOWS = platform.system() == "Windows"
IS_MACOS   = platform.system() == "Darwin"

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRESETS_DIR = os.path.join(BASE_DIR, "presets")


# ── Skill Data ───────────────────────────────────────────────────
class Skill:
    __slots__ = ("name", "cooldown", "enabled", "remaining", "running")
    def __init__(self, name="技能", cooldown=30, enabled=True):
        self.name      = name
        self.cooldown  = max(1, min(60, int(cooldown)))
        self.enabled   = bool(enabled)
        self.remaining = float(self.cooldown)
        self.running   = False

    def reset(self):
        self.remaining = float(self.cooldown)
        self.running   = False

    def to_dict(self):
        return {"name": self.name, "cooldown": self.cooldown, "enabled": self.enabled}

    @classmethod
    def from_dict(cls, d):
        return cls(d.get("name", "技能"), d.get("cooldown", 30), d.get("enabled", True))


# ── Voice Engine ─────────────────────────────────────────────────
class VoiceEngine:
    def __init__(self):
        self.enabled       = True
        self.rate          = 2
        self.volume        = 100
        self.early_warning = 0
        self._voice        = None
        self._init_engine()

    def _init_engine(self):
        if IS_WINDOWS:
            try:
                import win32com.client
                self._voice = win32com.client.Dispatch("SAPI.SpVoice")
            except Exception:
                self._voice = None

    def speak(self, text):
        if not self.enabled or not text:
            return
        def _run():
            try:
                if IS_WINDOWS and self._voice:
                    self._voice.Rate   = self.rate
                    self._voice.Volume = self.volume
                    self._voice.Speak(text, 1)
                elif IS_MACOS:
                    os.system(f'say -v Tingting "{text}" &')
            except Exception:
                pass
        threading.Thread(target=_run, daemon=True).start()

    def speak_skill(self, name):
        self.speak(name)

    def speak_warning(self, name):
        self.speak(f"准备 {name}")

    def test(self):
        self.speak(f"{APP_NAME} 语音测试")


# ── Preset Manager ───────────────────────────────────────────────
class PresetManager:
    def __init__(self):
        os.makedirs(PRESETS_DIR, exist_ok=True)
        self._ensure_defaults()

    def _ensure_defaults(self):
        demo_path = os.path.join(PRESETS_DIR, "demo_boss.json")
        tpl_path  = os.path.join(PRESETS_DIR, "new_boss_template.json")
        if not os.path.exists(demo_path):
            self._write(demo_path, {
                "name": "炎狱之王", "game": "演示副本", "notes": "P1技能轴",
                "skills": [
                    {"name": "火焰吐息", "cooldown": 30},
                    {"name": "暗影冲锋", "cooldown": 20},
                    {"name": "陨石坠落", "cooldown": 90},
                    {"name": "冰霜禁锢", "cooldown": 20},
                ]
            })
        if not os.path.exists(tpl_path):
            self._write(tpl_path, {
                "name": "新BOSS",
                "skills": [{"name": "技能1", "cooldown": 30}, {"name": "技能2", "cooldown": 20}]
            })

    def _write(self, path, data):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def list_files(self):
        if not os.path.exists(PRESETS_DIR):
            return []
        return sorted(f for f in os.listdir(PRESETS_DIR) if f.endswith(".json"))

    def load(self, filename):
        path = os.path.join(PRESETS_DIR, filename)
        if not os.path.exists(path):
            return None
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        skills = [Skill.from_dict(s) for s in data.get("skills", [])]
        return {
            "filename": filename,
            "name":     data.get("name", ""),
            "game":     data.get("game", ""),
            "notes":    data.get("notes", ""),
            "skills":   skills,
        }

    def save(self, filename, name, skills, game="", notes=""):
        data = {
            "name":   name,
            "game":   game,
            "notes":  notes,
            "skills": [s.to_dict() for s in skills],
        }
        self._write(os.path.join(PRESETS_DIR, filename), data)

    def delete(self, filename):
        path = os.path.join(PRESETS_DIR, filename)
        if os.path.exists(path):
            os.remove(path)


# ── Timer Engine ─────────────────────────────────────────────────
class TimerEngine:
    def __init__(self, interval=0.1):
        self._skills       = []
        self._interval     = interval
        self._lock         = threading.Lock()
        self._thread_running = False
        self._thread       = None
        self._warn_tracker = {}
        self.state         = "idle"

        self.on_tick     = lambda skills: None
        self.on_complete = lambda idx, skill: None
        self.on_state    = lambda state: None

    # ── skill management ──
    def get_skills(self):
        with self._lock:
            return list(self._skills)

    def set_skills(self, skills):
        with self._lock:
            self._skills = [Skill(s.name, s.cooldown, s.enabled) for s in skills]

    def add_skill(self, name="新技能", cd=30):
        with self._lock:
            s = Skill(name, cd)
            self._skills.append(s)
            return len(self._skills) - 1

    def remove_skill(self, idx):
        with self._lock:
            if 0 <= idx < len(self._skills):
                self._skills.pop(idx)

    # ── global controls ──
    def start_all(self):
        with self._lock:
            if self.state == "running":
                return
            self.state = "running"
            for s in self._skills:
                if s.enabled:
                    s.running = True
                    if s.remaining <= 0:
                        s.remaining = float(s.cooldown)
            self._warn_tracker.clear()
            self._ensure_thread()
            self.on_state("running")

    def pause_all(self):
        with self._lock:
            if self.state != "running":
                return
            self.state = "paused"
            self._stop_thread()
            for s in self._skills:
                s.running = False
            self.on_state("paused")

    def resume_all(self):
        with self._lock:
            if self.state != "paused":
                return
            self.state = "running"
            for s in self._skills:
                if s.enabled:
                    s.running = True
            self._warn_tracker.clear()
            self._ensure_thread()
            self.on_state("running")

    def reset_all(self):
        with self._lock:
            self._stop_thread()
            self.state = "idle"
            self._warn_tracker.clear()
            for s in self._skills:
                s.reset()
            self.on_state("idle")
            self.on_tick(self._skills)

    # ── per-skill controls ──
    def start_skill(self, idx):
        """Start countdown for a single skill (works regardless of global state)."""
        with self._lock:
            if 0 <= idx < len(self._skills):
                s = self._skills[idx]
                if not s.enabled:
                    return
                s.running   = True
                s.remaining = float(s.cooldown)
                self._warn_tracker.pop(idx, None)
                self._ensure_thread()

    def reset_skill(self, idx):
        with self._lock:
            if 0 <= idx < len(self._skills):
                s = self._skills[idx]
                s.remaining = float(s.cooldown)
                s.running   = (self.state == "running" and s.enabled)
                self._warn_tracker.pop(idx, None)

    def toggle_enabled(self, idx):
        with self._lock:
            if 0 <= idx < len(self._skills):
                s = self._skills[idx]
                s.enabled = not s.enabled
                if not s.enabled:
                    s.running   = False
                    s.remaining = float(s.cooldown)
                elif self.state == "running":
                    s.running   = True
                    s.remaining = float(s.cooldown)

    def update_skill(self, idx, name=None, cooldown=None):
        with self._lock:
            if 0 <= idx < len(self._skills):
                s = self._skills[idx]
                if name is not None:
                    s.name = name
                if cooldown is not None:
                    s.cooldown  = max(1, min(60, int(cooldown)))
                    s.remaining = float(s.cooldown)

    # ── internal ──
    def _ensure_thread(self):
        if self._thread_running:
            return
        self._thread_running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _stop_thread(self):
        self._thread_running = False

    def _loop(self):
        while self._thread_running:
            t0 = time.time()
            with self._lock:
                completed = []
                for i, s in enumerate(self._skills):
                    if s.running and s.enabled:
                        s.remaining = max(0.0, s.remaining - self._interval)
                        if s.remaining <= 0:
                            s.remaining = 0.0
                            s.running   = False
                            completed.append((i, s))
                for i, s in completed:
                    self.on_complete(i, s)
                    s.remaining = float(s.cooldown)
                self.on_tick(self._skills)
                # Auto-stop thread if nothing is running
                if self.state != "running" and not any(s.running for s in self._skills):
                    self._thread_running = False
                    break
            elapsed = time.time() - t0
            time.sleep(max(0.0, self._interval - elapsed))
        self._thread_running = False

    def check_early_warning(self, warn_seconds):
        if warn_seconds <= 0:
            return None, None
        with self._lock:
            for i, s in enumerate(self._skills):
                if s.running and s.enabled and 0 < s.remaining <= warn_seconds:
                    last = self._warn_tracker.get(i, 999)
                    if last > warn_seconds:
                        self._warn_tracker[i] = s.remaining
                        return i, s
                    self._warn_tracker[i] = s.remaining
        return None, None

    def stop(self):
        self._stop_thread()


# ── Overlay Helpers (Windows only) ───────────────────────────────
class Overlay:
    @staticmethod
    def apply(window, topmost=True, click_through=False, opacity=0.85):
        if not IS_WINDOWS:
            return
        try:
            import ctypes
            hwnd = window.winfo_id()
            GWL_EXSTYLE = -20
            WS_EX_LAYERED     = 0x80000
            WS_EX_TRANSPARENT = 0x20
            WS_EX_TOPMOST     = 0x8
            user32 = ctypes.windll.user32
            style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            style |= WS_EX_LAYERED
            if topmost:
                style |= WS_EX_TOPMOST
            if click_through:
                style |= WS_EX_TRANSPARENT
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
            user32.SetLayeredWindowAttributes(hwnd, 0, int(opacity * 255), 2)
        except Exception:
            pass

    @staticmethod
    def set_click_through(window, enable):
        if not IS_WINDOWS:
            return
        try:
            import ctypes
            hwnd = window.winfo_id()
            user32 = ctypes.windll.user32
            style = user32.GetWindowLongW(hwnd, -20)
            style = (style | 0x20) if enable else (style & ~0x20)
            user32.SetWindowLongW(hwnd, -20, style)
        except Exception:
            pass

    @staticmethod
    def set_opacity(window, opacity):
        if not IS_WINDOWS:
            return
        try:
            import ctypes
            hwnd = window.winfo_id()
            ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, int(opacity * 255), 2)
        except Exception:
            pass


# ── Skill Row Widget ─────────────────────────────────────────────
class SkillRow(tk.Frame):
    def __init__(self, parent, index, skill, callbacks):
        super().__init__(parent, bg=COLORS["row"], height=52)
        self.idx   = index
        self.skill = skill
        self.cb    = callbacks
        self.pack_propagate(False)
        self.pack(fill=tk.X, padx=2, pady=1)
        self._build()

    def _build(self):
        s = self.skill
        row_bg = COLORS["row"]

        # Enable dot
        self.dot = tk.Label(self, text="●", font=FONT_NORMAL,
                            cursor="hand2", bg=row_bg, fg=COLORS["green"])
        self.dot.pack(side=tk.LEFT, padx=(6, 4), pady=10)
        self.dot.bind("<Button-1>", lambda e: self.cb["toggle_enabled"](self.idx))

        # Skill name
        self.lbl_name = tk.Label(self, text=s.name, font=FONT_BOLD,
                                 bg=row_bg, fg=COLORS["text"],
                                 width=12, anchor=tk.W)
        self.lbl_name.pack(side=tk.LEFT, padx=4, pady=10)
        self.lbl_name.bind("<Double-Button-1>", lambda e: self.cb["edit_name"](self.idx))

        # Cooldown label
        self.lbl_cd = tk.Label(self, text=f"{s.cooldown}s", font=FONT_NORMAL,
                               bg=row_bg, fg=COLORS["muted"], width=5)
        self.lbl_cd.pack(side=tk.LEFT, padx=4, pady=10)
        self.lbl_cd.bind("<Double-Button-1>", lambda e: self.cb["edit_cd"](self.idx))

        # Progress bar
        self.canvas = tk.Canvas(self, bg=COLORS["bg"], height=20,
                                highlightthickness=1, highlightbackground=COLORS["border"])
        self.canvas.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=10)

        # Time remaining
        self.lbl_time = tk.Label(self, text=f"{s.cooldown:.1f}s", font=FONT_BOLD,
                                 bg=row_bg, fg=COLORS["green"], width=7)
        self.lbl_time.pack(side=tk.LEFT, padx=4, pady=10)

        # Start button
        self.btn_start = tk.Label(self, text="▶ 开始", font=FONT_SMALL,
                                  cursor="hand2", bg=COLORS["green"], fg="#fff",
                                  padx=8, pady=2)
        self.btn_start.pack(side=tk.LEFT, padx=2, pady=10)
        self.btn_start.bind("<Button-1>", lambda e: self.cb["start_skill"](self.idx))

        # Reset button
        self.btn_reset = tk.Label(self, text="🔄 重置", font=FONT_SMALL,
                                  cursor="hand2", bg="#5c1a1a", fg=COLORS["red"],
                                  padx=8, pady=2)
        self.btn_reset.pack(side=tk.LEFT, padx=2, pady=10)
        self.btn_reset.bind("<Button-1>", lambda e: self.cb["reset_skill"](self.idx))

    def refresh(self, skill):
        self.skill = skill
        s = skill
        dim_color   = COLORS["dim"]
        muted_color = COLORS["muted"]

        # Dot
        self.dot.config(
            text="●" if s.enabled else "○",
            fg=COLORS["green"] if s.enabled else dim_color)

        # Name
        self.lbl_name.config(
            text=s.name,
            fg=dim_color if not s.enabled else COLORS["text"])

        # Cooldown
        self.lbl_cd.config(
            text=f"{s.cooldown}s",
            fg=dim_color if not s.enabled else muted_color)

        # Time + progress bar
        if not s.enabled:
            self.lbl_time.config(text="--", fg=dim_color)
            self._draw_bar(0.0, dim_color)
        elif not s.running:
            self.lbl_time.config(text=f"{s.cooldown:.1f}s", fg=COLORS["green"])
            self._draw_bar(1.0, COLORS["green"])
        else:
            r     = max(0.0, s.remaining)
            ratio = r / s.cooldown if s.cooldown > 0 else 0.0
            color = COLORS["red"] if r <= 5 else (COLORS["yellow"] if r <= 10 else COLORS["green"])
            self.lbl_time.config(text=f"{r:.1f}s", fg=color)
            self._draw_bar(ratio, color)

        # Start button
        if not s.enabled:
            self.btn_start.config(bg=COLORS["dim"], fg=COLORS["bg"], text="▶ 开始")
        elif s.running:
            self.btn_start.config(bg="#5c3d1a", fg=COLORS["yellow"], text="⏸ 暂停")
        else:
            self.btn_start.config(bg=COLORS["green"], fg="#fff", text="▶ 开始")

        # Row background
        warn   = s.running and s.enabled and s.remaining <= 5
        new_bg = COLORS["row_warn"] if warn else COLORS["row"]
        self.config(bg=new_bg)
        for w in (self.dot, self.lbl_name, self.lbl_time):
            try:
                w.config(bg=new_bg)
            except tk.TclError:
                pass

    def _draw_bar(self, ratio, color):
        self.canvas.delete("all")
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w < 5:
            return
        fw = int(w * max(0.0, min(1.0, ratio)))
        if fw > 0:
            self.canvas.create_rectangle(0, 0, fw, h, fill=color, outline="")
        if fw < w:
            self.canvas.create_rectangle(fw, 0, w, h, fill=COLORS["bg"], outline="")


# ── Settings Dialog ──────────────────────────────────────────────
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, voice, callbacks):
        super().__init__(parent)
        self.voice = voice
        self.cb    = callbacks
        self.title("设置")
        self.geometry("420x380")
        self.resizable(False, False)
        self.configure(bg=COLORS["bg"])
        self.transient(parent)
        self._build()
        self.update_idletasks()
        x = parent.winfo_x() + parent.winfo_width() // 2 - 210
        y = parent.winfo_y() + 30
        self.geometry(f"+{x}+{y}")

    def _build(self):
        pad = {"padx": 16, "pady": 6}
        tk.Label(self, text="⚙ 语音与显示设置", font=FONT_BOLD,
                 fg=COLORS["text"], bg=COLORS["bg"]).pack(anchor=tk.W, **pad)

        vf = tk.LabelFrame(self, text="🔊 语音播报", font=FONT_SMALL,
                           fg=COLORS["muted"], bg=COLORS["bg"], padx=12, pady=8)
        vf.pack(fill=tk.X, padx=14, pady=4)

        self.voice_var = tk.BooleanVar(value=self.voice.enabled)
        tk.Checkbutton(vf, text="启用语音播报", variable=self.voice_var,
                       font=FONT_SMALL, fg=COLORS["text"], bg=COLORS["bg"],
                       selectcolor=COLORS["bg"], activebackground=COLORS["bg"],
                       activeforeground=COLORS["text"],
                       command=lambda: setattr(self.voice, 'enabled', self.voice_var.get())
                       ).pack(anchor=tk.W)

        rf = tk.Frame(vf, bg=COLORS["bg"]); rf.pack(fill=tk.X, pady=2)
        tk.Label(rf, text="语速:", font=FONT_SMALL, fg=COLORS["muted"],
                 bg=COLORS["bg"], width=5).pack(side=tk.LEFT)
        self.rate_var = tk.IntVar(value=self.voice.rate)
        tk.Scale(rf, from_=-10, to=10, orient=tk.HORIZONTAL,
                 variable=self.rate_var, bg=COLORS["bg"], fg=COLORS["text"],
                 highlightthickness=0, troughcolor=COLORS["border"],
                 command=lambda v: setattr(self.voice, 'rate', int(float(v)))
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        wf = tk.Frame(vf, bg=COLORS["bg"]); wf.pack(fill=tk.X, pady=2)
        tk.Label(wf, text="预警:", font=FONT_SMALL, fg=COLORS["muted"],
                 bg=COLORS["bg"], width=5).pack(side=tk.LEFT)
        self.warn_var = tk.IntVar(value=self.voice.early_warning)
        for val, label in [(0, "关"), (3, "3s"), (5, "5s"), (10, "10s")]:
            tk.Radiobutton(wf, text=label, variable=self.warn_var, value=val,
                           font=FONT_SMALL, fg=COLORS["muted"], bg=COLORS["bg"],
                           selectcolor=COLORS["bg"], activebackground=COLORS["bg"],
                           activeforeground=COLORS["text"],
                           command=lambda v=val: setattr(self.voice, 'early_warning', v)
                           ).pack(side=tk.LEFT, padx=3)

        tk.Button(vf, text="📢 测试语音", font=FONT_SMALL,
                  bg=COLORS["border"], fg=COLORS["text"], relief=tk.FLAT,
                  cursor="hand2", command=self.voice.test,
                  padx=12, pady=3).pack(anchor=tk.W, pady=4)

        df = tk.LabelFrame(self, text="🖥 显示设置", font=FONT_SMALL,
                           fg=COLORS["muted"], bg=COLORS["bg"], padx=12, pady=8)
        df.pack(fill=tk.X, padx=14, pady=4)

        self.top_var = tk.BooleanVar(value=True)
        tk.Checkbutton(df, text="窗口置顶", variable=self.top_var,
                       font=FONT_SMALL, fg=COLORS["text"], bg=COLORS["bg"],
                       selectcolor=COLORS["bg"], activebackground=COLORS["bg"],
                       command=lambda: self.cb["topmost"](self.top_var.get())
                       ).pack(anchor=tk.W)

        self.ct_var = tk.BooleanVar(value=False)
        tk.Checkbutton(df, text="🖱 点击穿透", variable=self.ct_var,
                       font=FONT_SMALL, fg=COLORS["text"], bg=COLORS["bg"],
                       selectcolor=COLORS["bg"], activebackground=COLORS["bg"],
                       command=lambda: self.cb["click"](self.ct_var.get())
                       ).pack(anchor=tk.W)

        of = tk.Frame(df, bg=COLORS["bg"]); of.pack(fill=tk.X, pady=2)
        tk.Label(of, text="透明度:", font=FONT_SMALL, fg=COLORS["muted"],
                 bg=COLORS["bg"], width=5).pack(side=tk.LEFT)
        self.op_var = tk.IntVar(value=85)
        self.lbl_op = tk.Label(of, text="85%", font=FONT_SMALL,
                               fg=COLORS["muted"], bg=COLORS["bg"], width=4)
        self.lbl_op.pack(side=tk.RIGHT)
        tk.Scale(of, from_=30, to=100, orient=tk.HORIZONTAL,
                 variable=self.op_var, bg=COLORS["bg"], fg=COLORS["text"],
                 highlightthickness=0, troughcolor=COLORS["border"],
                 command=lambda v: (
                     self.lbl_op.config(text=f"{int(float(v))}%"),
                     self.cb.get("opacity", lambda x: None)(int(float(v)) / 100.0)
                 )).pack(side=tk.LEFT, fill=tk.X, expand=True)


# ── Main Application ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("780x520")
        self.minsize(600, 320)
        self.configure(bg=COLORS["bg"])

        self.timer   = TimerEngine(0.1)
        self.voice   = VoiceEngine()
        self.presets = PresetManager()

        self.timer.on_tick     = lambda skills: self.after(0, self._refresh_ui)
        self.timer.on_complete = lambda idx, s: self.after(0, lambda: self._on_skill_done(idx, s))
        self.timer.on_state    = lambda st: self.after(0, lambda: self._on_timer_state(st))

        self.rows         = []
        self.current_file = None
        self._topmost     = True
        self._click_thru  = False

        self._build_ui()
        self._load_preset("demo_boss.json")
        self.after(300, lambda: Overlay.apply(self, True, False, 0.85))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI ───────────────────────────────────────────────────
    def _build_ui(self):
        # Title bar
        tb = tk.Frame(self, bg="#1a1030", height=36)
        tb.pack(fill=tk.X); tb.pack_propagate(False)

        self.lbl_boss = tk.Label(tb, text=f"🔥 {APP_NAME}",
                                 font=FONT_BOLD, fg=COLORS["yellow"], bg="#1a1030")
        self.lbl_boss.pack(side=tk.LEFT, padx=12, pady=4)

        bf = tk.Frame(tb, bg="#1a1030"); bf.pack(side=tk.RIGHT, padx=8)
        self.btn_pin = tk.Label(bf, text="📌", font=("", 13),
                                cursor="hand2", bg="#1a1030", fg=COLORS["green"], width=3)
        self.btn_pin.pack(side=tk.LEFT)
        self.btn_pin.bind("<Button-1>", lambda e: self._toggle_topmost())

        self.btn_click = tk.Label(bf, text="🖱", font=("", 13),
                                  cursor="hand2", bg="#1a1030", fg=COLORS["muted"], width=3)
        self.btn_click.pack(side=tk.LEFT)
        self.btn_click.bind("<Button-1>", lambda e: self._toggle_click_through())

        self.btn_cfg = tk.Label(bf, text="⚙", font=("", 13),
                                cursor="hand2", bg="#1a1030", fg=COLORS["muted"], width=3)
        self.btn_cfg.pack(side=tk.LEFT)
        self.btn_cfg.bind("<Button-1>", lambda e: self._open_settings())

        # Header
        hf = tk.Frame(self, bg=COLORS["header"], height=26)
        hf.pack(fill=tk.X, padx=4, pady=(4, 0)); hf.pack_propagate(False)
        for t, w in [("启用", 5), ("技能名称", 14), ("冷却", 5),
                     ("倒计时进度", 38), ("剩余", 6), ("操作", 14)]:
            tk.Label(hf, text=t, font=FONT_SMALL, fg=COLORS["muted"],
                     bg=COLORS["header"], width=w).pack(side=tk.LEFT, padx=2, pady=3)

        # Scrollable skill list
        self.list_canvas = tk.Canvas(self, bg=COLORS["bg"], highlightthickness=0)
        self.scrollbar   = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.list_canvas.yview)
        self.skill_frame = tk.Frame(self.list_canvas, bg=COLORS["bg"])
        self.skill_frame.bind("<Configure>",
            lambda e: self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all")))
        self.list_canvas.create_window((0, 0), window=self.skill_frame, anchor=tk.NW)
        self.list_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_canvas.bind_all("<MouseWheel>",
            lambda e: self.list_canvas.yview_scroll(int(-e.delta / 120), "units"))

        # Bottom bar
        bb = tk.Frame(self, bg=COLORS["panel"], height=72)
        bb.pack(fill=tk.X, side=tk.BOTTOM, padx=4, pady=4); bb.pack_propagate(False)

        lf = tk.Frame(bb, bg=COLORS["panel"]); lf.pack(side=tk.LEFT, padx=8, pady=12)
        tk.Label(lf, text="+ 添加技能", font=FONT_SMALL, bg=COLORS["blue"],
                 fg=COLORS["text"], cursor="hand2", padx=12, pady=4
                 ).pack(side=tk.LEFT, padx=4)
        lf.winfo_children()[-1].bind("<Button-1>", lambda e: self._add_skill())

        files = self.presets.list_files()
        self.preset_var = tk.StringVar(value=files[0] if files else "无方案")
        self.preset_menu = tk.OptionMenu(lf, self.preset_var,
                                         files[0] if files else "无方案", *files,
                                         command=self._load_preset)
        self.preset_menu.config(font=FONT_SMALL, bg=COLORS["bg"],
                                fg=COLORS["muted"], highlightthickness=0, width=18)
        self.preset_menu.pack(side=tk.LEFT, padx=6)

        tk.Label(lf, text="💾 保存", font=FONT_SMALL, bg=COLORS["border"],
                 fg=COLORS["text"], cursor="hand2", padx=10, pady=4
                 ).pack(side=tk.LEFT, padx=4)
        lf.winfo_children()[-1].bind("<Button-1>", lambda e: self._save_preset())

        rf = tk.Frame(bb, bg=COLORS["panel"]); rf.pack(side=tk.RIGHT, padx=8, pady=12)
        self.btn_start = tk.Label(rf, text="▶ 全部开始", font=FONT_BOLD,
                                  bg=COLORS["green"], fg="#fff", cursor="hand2",
                                  padx=14, pady=5)
        self.btn_start.pack(side=tk.LEFT, padx=3)
        self.btn_start.bind("<Button-1>", lambda e: self._start_pause())

        tk.Label(rf, text="⏸ 暂停", font=FONT_BOLD, bg="#5c3d1a",
                 fg=COLORS["yellow"], cursor="hand2", padx=12, pady=5
                 ).pack(side=tk.LEFT, padx=3)
        rf.winfo_children()[-1].bind("<Button-1>", lambda e: self.timer.pause_all())

        tk.Label(rf, text="🔄 重置", font=FONT_BOLD, bg="#5c1a1a",
                 fg=COLORS["red"], cursor="hand2", padx=12, pady=5
                 ).pack(side=tk.LEFT, padx=3)
        rf.winfo_children()[-1].bind("<Button-1>", lambda e: self.timer.reset_all())

        self.lbl_voice = tk.Label(bb, text="🔊", font=("", 15),
                                  bg=COLORS["panel"], fg=COLORS["green"], cursor="hand2")
        self.lbl_voice.pack(side=tk.RIGHT, padx=6, pady=18)
        self.lbl_voice.bind("<Button-1>", lambda e: self._toggle_voice())

        self.lbl_state = tk.Label(bb, text="🟢 准备就绪", font=FONT_SMALL,
                                  bg=COLORS["panel"], fg=COLORS["muted"])
        self.lbl_state.pack(side=tk.RIGHT, padx=12, pady=18)

    # ── Skill Rows ───────────────────────────────────────────
    def _rebuild_rows(self):
        for row in self.rows:
            row.destroy()
        self.rows.clear()
        callbacks = {
            "toggle_enabled": self._on_toggle_enabled,
            "start_skill":    self._on_start_skill,
            "reset_skill":    self._on_reset_skill,
            "edit_name":      self._on_edit_name,
            "edit_cd":        self._on_edit_cd,
        }
        for i, s in enumerate(self.timer.get_skills()):
            self.rows.append(SkillRow(self.skill_frame, i, s, callbacks))
        self.skill_frame.update_idletasks()
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))

    def _refresh_ui(self):
        skills = self.timer.get_skills()
        for i, row in enumerate(self.rows):
            if i < len(skills):
                row.refresh(skills[i])
        self._check_early_warning()

    def _check_early_warning(self):
        ws = self.voice.early_warning
        if ws > 0:
            idx, s = self.timer.check_early_warning(ws)
            if s:
                self.voice.speak_warning(s.name)

    # ── Events ───────────────────────────────────────────────
    def _on_skill_done(self, idx, skill):
        self.voice.speak_skill(skill.name)

    def _on_timer_state(self, state):
        if state == "running":
            self.lbl_state.config(text="🔴 运行中", fg=COLORS["red"])
            self.btn_start.config(text="⏸ 全部暂停", bg="#5c3d1a", fg=COLORS["yellow"])
        elif state == "paused":
            self.lbl_state.config(text="🟡 已暂停", fg=COLORS["yellow"])
            self.btn_start.config(text="▶ 继续", bg=COLORS["green"], fg="#fff")
        else:
            self.lbl_state.config(text="🟢 准备就绪", fg=COLORS["green"])
            self.btn_start.config(text="▶ 全部开始", bg=COLORS["green"], fg="#fff")

    def _start_pause(self):
        s = self.timer.state
        if s == "running":   self.timer.pause_all()
        elif s == "paused":  self.timer.resume_all()
        else:                self.timer.start_all()

    def _on_toggle_enabled(self, idx):
        self.timer.toggle_enabled(idx); self._refresh_ui()

    def _on_start_skill(self, idx):
        """▶ button: start or pause a single skill."""
        skills = self.timer.get_skills()
        if idx < len(skills):
            s = skills[idx]
            if not s.enabled:
                return
            if s.running:
                # Pause this skill
                with self.timer._lock:
                    s.running = False
                self._refresh_ui()
            else:
                # Start this skill
                self.timer.start_skill(idx)
                self._refresh_ui()

    def _on_reset_skill(self, idx):
        self.timer.reset_skill(idx); self._refresh_ui()

    def _on_edit_name(self, idx):
        skills = self.timer.get_skills()
        if idx < len(skills):
            new = simpledialog.askstring("编辑技能名", "技能名称:",
                                         initialvalue=skills[idx].name, parent=self)
            if new and new.strip():
                self.timer.update_skill(idx, name=new.strip())
                self._refresh_ui()

    def _on_edit_cd(self, idx):
        skills = self.timer.get_skills()
        if idx < len(skills):
            new = simpledialog.askinteger("编辑冷却", "冷却时间 (1-60秒):",
                                          initialvalue=skills[idx].cooldown,
                                          minvalue=1, maxvalue=60, parent=self)
            if new:
                self.timer.update_skill(idx, cooldown=new)
                self._refresh_ui()

    def _add_skill(self):
        self.timer.add_skill("新技能", 30)
        self._rebuild_rows(); self._refresh_ui()

    def _load_preset(self, filename):
        if not filename or filename == "无方案":
            return
        data = self.presets.load(filename)
        if data:
            self.timer.reset_all()
            self.timer.set_skills(data["skills"])
            self.current_file = filename
            self.lbl_boss.config(text=f"🔥 {data['name']}")
            self._rebuild_rows(); self._refresh_ui()

    def _save_preset(self):
        skills = self.timer.get_skills()
        if not skills:
            messagebox.showwarning("提示", "没有技能可保存", parent=self)
            return
        name = simpledialog.askstring("保存方案", "BOSS 名称:", parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        self.presets.save(f"{name}.json", name, skills)
        self.current_file = f"{name}.json"
        self.lbl_boss.config(text=f"🔥 {name}")
        self._update_preset_menu()

    def _update_preset_menu(self):
        files = self.presets.list_files()
        menu  = self.preset_menu["menu"]
        menu.delete(0, "end")
        for f in files:
            menu.add_command(label=f, command=lambda v=f: self.preset_var.set(v) or self._load_preset(v))
        if files:
            self.preset_var.set(self.current_file or files[0])

    def _toggle_topmost(self):
        self._topmost = not self._topmost
        self.btn_pin.config(fg=COLORS["green"] if self._topmost else COLORS["muted"])
        try:
            self.attributes('-topmost', self._topmost)
        except tk.TclError:
            pass

    def _toggle_click_through(self):
        self._click_thru = not self._click_thru
        self.btn_click.config(fg=COLORS["green"] if self._click_thru else COLORS["muted"])
        Overlay.set_click_through(self, self._click_thru)

    def _toggle_voice(self):
        self.voice.enabled = not self.voice.enabled
        self.lbl_voice.config(
            fg=COLORS["green"] if self.voice.enabled else COLORS["dim"],
            text="🔊" if self.voice.enabled else "🔇")

    def _open_settings(self):
        callbacks = {
            "topmost": lambda v: (
                setattr(self, '_topmost', v),
                self.btn_pin.config(fg=COLORS["green"] if v else COLORS["muted"]),
                self.attributes('-topmost', v)
            ),
            "click": lambda v: (
                setattr(self, '_click_thru', v),
                self.btn_click.config(fg=COLORS["green"] if v else COLORS["muted"]),
                Overlay.set_click_through(self, v)
            ),
            "opacity": lambda v: Overlay.set_opacity(self, v),
        }
        SettingsDialog(self, self.voice, callbacks)

    def _on_close(self):
        self.timer.stop()
        self.destroy()


# ── Entry Point ──────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
